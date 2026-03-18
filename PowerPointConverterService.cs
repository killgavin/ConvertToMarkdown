using System.Security.Cryptography;
using System.Text;
using PptApp = Microsoft.Office.Interop.PowerPoint.Application;
using MsoTriState = Microsoft.Office.Core.MsoTriState;

namespace ConvertToMarkdown;

/// <summary>
/// PowerPoint 轉換服務實作 - 負責將 PowerPoint (.pptx / .ppt) 檔案透過
/// PowerPoint COM Interop 讀取投影片內容，並轉換為 GitHub Flavored Markdown (GFM) 格式。
/// 每張投影片以 Markdown 標題分隔，文字與表格擷取為 Markdown 格式，
/// 圖片類型的圖形則匯出為獨立圖檔。
/// </summary>
public class PowerPointConverterService : IPowerPointConverterService
{
    /// <summary>
    /// 非同步執行 PowerPoint 轉 Markdown 完整流程。
    /// 整個轉換工作會在背景執行緒執行，以避免 UI 凍結 (UI Freeze)。
    /// </summary>
    /// <param name="sourceFilePath">來源 PowerPoint 檔案的完整路徑。</param>
    /// <param name="progress">進度回報介面，用於向 UI 輸出執行日誌。</param>
    /// <returns>
    /// 非同步工作，完成後傳回 <see cref="ConversionResult"/>。
    /// </returns>
    public async Task<ConversionResult> ConvertAsync(string sourceFilePath, IProgress<string> progress)
    {
        return await Task.Run(() =>
        {
            PptApp? pptApp = null;
            Microsoft.Office.Interop.PowerPoint.Presentation? presentation = null;
            try
            {
                // === 步驟 1：驗證來源檔案是否存在 ===
                progress.Report("▶ [1/4] 驗證來源 PowerPoint 檔案...");
                if (!File.Exists(sourceFilePath))
                {
                    return new ConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"找不到指定的檔案：{sourceFilePath}"
                    };
                }

                // === 步驟 2：建立輸出資料夾 ===
                string sourceDir = Path.GetDirectoryName(sourceFilePath)!;
                string fileBaseName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputDir = Path.Combine(sourceDir, fileBaseName);
                Directory.CreateDirectory(outputDir);
                progress.Report($"▶ [2/4] 輸出資料夾：{outputDir}");

                // === 步驟 3：透過 PowerPoint COM Interop 開啟簡報 ===
                progress.Report("▶ [3/4] PowerPoint COM Interop 開啟簡報中...");

                try
                {
                    pptApp = new PptApp();
                    pptApp.DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone;

                    presentation = pptApp.Presentations.Open(
                        sourceFilePath,
                        ReadOnly: MsoTriState.msoTrue,
                        Untitled: MsoTriState.msoFalse,
                        WithWindow: MsoTriState.msoFalse);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return new ConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"無法啟動 Microsoft PowerPoint 進行轉換。請確認已安裝 Microsoft PowerPoint。\n詳細錯誤：{ex.Message}"
                    };
                }

                int slideCount = presentation.Slides.Count;
                progress.Report($"  ✔ 共偵測到 {slideCount} 張投影片");

                // === 步驟 4：逐張投影片擷取內容並組成 Markdown ===
                progress.Report($"▶ [4/4] 擷取投影片內容中...");

                var sb = new StringBuilder();
                int imageIndex = 0;

                for (int i = 1; i <= slideCount; i++)
                {
                    var slide = presentation.Slides[i];
                    progress.Report($"  ▷ 正在處理投影片 {i}/{slideCount}...");

                    // 投影片標題
                    sb.AppendLine($"## 投影片 {i}");
                    sb.AppendLine();

                    // 遍歷投影片上的所有圖形（Shape）
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                    {
                        try
                        {
                            // 處理表格
                            if (shape.HasTable == MsoTriState.msoTrue)
                            {
                                var table = shape.Table;
                                sb.AppendLine(ConvertTableToMarkdown(table));
                                continue;
                            }

                            // 處理含有文字的圖形
                            if (shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                var textFrame = shape.TextFrame;
                                if (textFrame.HasText == MsoTriState.msoTrue)
                                {
                                    string text = textFrame.TextRange.Text?.Trim() ?? string.Empty;
                                    if (!string.IsNullOrWhiteSpace(text))
                                    {
                                        // 若為標題預留位置，使用 Markdown 標題格式
                                        if (IsSlideTitle(shape))
                                        {
                                            sb.AppendLine($"### {text}");
                                        }
                                        else
                                        {
                                            sb.AppendLine(text);
                                        }
                                        sb.AppendLine();
                                    }
                                }
                            }

                            // 匯出圖片型態的圖形
                            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture
                                || shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                            {
                                imageIndex++;
                                string? imageFileName = ExportShapeAsImage(shape, outputDir, imageIndex, progress);
                                if (imageFileName != null)
                                {
                                    sb.AppendLine($"![圖片 {imageIndex}]({imageFileName})");
                                    sb.AppendLine();
                                }
                            }
                        }
                        catch (Exception shapeEx)
                        {
                            progress.Report($"    ⚠ 處理圖形時發生錯誤，跳過：{shapeEx.Message}");
                        }
                    }

                    sb.AppendLine("---");
                    sb.AppendLine();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(slide);
                }

                // 寫出 Markdown 檔案
                string outputFilePath = Path.Combine(outputDir, $"{fileBaseName}.md");
                File.WriteAllText(outputFilePath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                progress.Report($"  ✔ Markdown 檔案已產出：{outputFilePath}");

                return new ConversionResult
                {
                    IsSuccess = true,
                    OutputFilePath = outputFilePath
                };
            }
            catch (Exception ex)
            {
                return new ConversionResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"轉換過程發生例外：{ex.GetType().Name} - {ex.Message}"
                };
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp);
                }
            }
        });
    }

    /// <summary>
    /// 判斷 PowerPoint 圖形是否為投影片標題預留位置。
    /// </summary>
    /// <param name="shape">PowerPoint 圖形物件。</param>
    /// <returns>若為標題或置中標題預留位置，傳回 true。</returns>
    private static bool IsSlideTitle(Microsoft.Office.Interop.PowerPoint.Shape shape)
    {
        try
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
            {
                var phType = shape.PlaceholderFormat.Type;
                return phType == Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                    || phType == Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle;
            }
        }
        catch
        {
            // COM 存取 PlaceholderFormat 可能在非預留位置圖形上擲出例外，安全忽略
        }
        return false;
    }

    /// <summary>
    /// 將 PowerPoint 圖片圖形匯出為圖檔並儲存至輸出目錄。
    /// 使用圖片內容的 SHA256 雜湊值前 16 碼作為檔名，確保唯一且可去重。
    /// </summary>
    /// <param name="shape">含有圖片的 PowerPoint 圖形物件。</param>
    /// <param name="outputDir">輸出目錄的完整路徑。</param>
    /// <param name="imageIndex">圖片序號（用於暫存檔命名）。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>儲存成功時傳回圖片檔名；失敗時傳回 null。</returns>
    private static string? ExportShapeAsImage(
        Microsoft.Office.Interop.PowerPoint.Shape shape,
        string outputDir,
        int imageIndex,
        IProgress<string> progress)
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"ppt_img_{Guid.NewGuid():N}.png");
        try
        {
            // 使用 Shape.Export 匯出圖片
            shape.Export(tempPath, Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG);

            byte[] imageBytes = File.ReadAllBytes(tempPath);
            string hash = Convert.ToHexString(SHA256.HashData(imageBytes))[..16].ToLowerInvariant();
            string imageFileName = $"img_{hash}.png";
            string imageFilePath = Path.Combine(outputDir, imageFileName);

            if (!File.Exists(imageFilePath))
            {
                File.Copy(tempPath, imageFilePath);
                progress.Report($"  ✔ 圖片已匯出：{imageFileName}（{imageBytes.Length:N0} bytes）");
            }
            else
            {
                progress.Report($"  ✔ 圖片已存在（重複引用）：{imageFileName}");
            }

            return imageFileName;
        }
        catch
        {
            // Shape.Export 不一定在所有版本可用，匯出失敗時回報警告
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); }
                catch { /* 忽略清除失敗 */ }
            }
        }

        // 匯出失敗時標註圖片存在但無法匯出
        progress.Report($"  ⚠ 第 {imageIndex} 張圖片無法直接匯出，請手動擷取。");
        return null;
    }

    /// <summary>
    /// 將 PowerPoint 表格轉換為 GFM Markdown 表格字串。
    /// </summary>
    /// <param name="table">PowerPoint COM Interop 的表格物件。</param>
    /// <returns>符合 GFM 語法的 Markdown 表格字串。</returns>
    private static string ConvertTableToMarkdown(Microsoft.Office.Interop.PowerPoint.Table table)
    {
        int rowCount = table.Rows.Count;
        int colCount = table.Columns.Count;

        if (rowCount == 0 || colCount == 0) return string.Empty;

        var sb = new StringBuilder();

        // 表頭列（第 1 列）
        sb.Append('|');
        for (int c = 1; c <= colCount; c++)
        {
            string cellText = GetCellText(table, 1, c);
            sb.Append($" {EscapeMarkdownCell(cellText)} |");
        }
        sb.AppendLine();

        // 分隔線列
        sb.Append('|');
        for (int c = 1; c <= colCount; c++)
        {
            sb.Append("---|");
        }
        sb.AppendLine();

        // 資料列（第 2 列起）
        for (int r = 2; r <= rowCount; r++)
        {
            sb.Append('|');
            for (int c = 1; c <= colCount; c++)
            {
                string cellText = GetCellText(table, r, c);
                sb.Append($" {EscapeMarkdownCell(cellText)} |");
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// 安全地讀取 PowerPoint 表格儲存格的文字內容。
    /// </summary>
    /// <param name="table">PowerPoint 表格物件。</param>
    /// <param name="row">列索引（1-based）。</param>
    /// <param name="col">欄索引（1-based）。</param>
    /// <returns>儲存格文字；若存取失敗則傳回空字串。</returns>
    private static string GetCellText(Microsoft.Office.Interop.PowerPoint.Table table, int row, int col)
    {
        try
        {
            return table.Cell(row, col).Shape.TextFrame.TextRange.Text?.Trim() ?? " ";
        }
        catch
        {
            return " ";
        }
    }

    /// <summary>
    /// 跳脫 Markdown 表格儲存格中的特殊字元。
    /// </summary>
    /// <param name="value">儲存格原始字串值。</param>
    /// <returns>跳脫特殊字元後的字串。</returns>
    private static string EscapeMarkdownCell(string value)
    {
        if (string.IsNullOrEmpty(value)) return " ";

        return value
            .Replace("|", @"\|")
            .Replace("\r\n", " ")
            .Replace("\n", " ")
            .Replace("\r", " ");
    }
}
