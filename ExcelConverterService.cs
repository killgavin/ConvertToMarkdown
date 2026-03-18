using System.Text;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ConvertToMarkdown;

/// <summary>
/// Excel 轉換服務實作 - 負責將 Excel (.xlsx / .xls) 檔案的每個工作表
/// 透過 Excel COM Interop 讀取資料，並轉換為 GitHub Flavored Markdown (GFM) 標準表格格式，
/// 各工作表輸出為獨立 .md 檔案。
/// </summary>
public class ExcelConverterService : IExcelConverterService
{
    /// <summary>
    /// 非同步執行 Excel 轉 Markdown 完整流程。
    /// 整個轉換工作會在背景執行緒執行，以避免 UI 凍結 (UI Freeze)。
    /// </summary>
    /// <param name="sourceFilePath">來源 Excel 檔案的完整路徑。</param>
    /// <param name="progress">進度回報介面，用於向 UI 輸出執行日誌。</param>
    /// <returns>
    /// 非同步工作，完成後傳回每個工作表的 <see cref="ExcelConversionResult"/> 清單。
    /// </returns>
    public async Task<IReadOnlyList<ExcelConversionResult>> ConvertAsync(
        string sourceFilePath,
        IProgress<string> progress)
    {
        // 使用 Task.Run 將整個 I/O 密集工作移至執行緒集區，確保 UI 執行緒不被阻塞
        return await Task.Run(() =>
        {
            var results = new List<ExcelConversionResult>();

            ExcelApp? excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook? workbook = null;
            try
            {
                // === 步驟 1：驗證來源檔案是否存在 ===
                progress.Report("▶ [1/4] 驗證來源 Excel 檔案...");
                if (!File.Exists(sourceFilePath))
                {
                    results.Add(new ExcelConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"找不到指定的檔案：{sourceFilePath}"
                    });
                    return results;
                }

                // === 步驟 2：建立輸出資料夾（以來源檔案主檔名命名）===
                string sourceDir = Path.GetDirectoryName(sourceFilePath)!;
                string excelBaseName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputDir = Path.Combine(sourceDir, excelBaseName);
                Directory.CreateDirectory(outputDir);
                progress.Report($"▶ [2/4] 輸出資料夾：{outputDir}");

                // === 步驟 3：透過 Excel COM Interop 開啟活頁簿 ===
                progress.Report("▶ [3/4] Excel COM Interop 開啟活頁簿中...");

                try
                {
                    excelApp = new ExcelApp { Visible = false, DisplayAlerts = false };
                    workbook = excelApp.Workbooks.Open(
                        sourceFilePath,
                        ReadOnly: true,
                        UpdateLinks: 0);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    results.Add(new ExcelConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"無法啟動 Microsoft Excel 進行轉換。請確認已安裝 Microsoft Excel。\n詳細錯誤：{ex.Message}"
                    });
                    return results;
                }

                int sheetCount = workbook.Worksheets.Count;

                // === 步驟 4：逐一轉換每個工作表為 Markdown 表格 ===
                progress.Report($"▶ [4/4] 共偵測到 {sheetCount} 個工作表，開始逐一轉換...");

                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    string sheetName = worksheet.Name;
                    progress.Report($"  ▷ 正在處理工作表：【{sheetName}】");

                    try
                    {
                        // 讀取工作表使用範圍的資料
                        string markdownContent = ConvertWorksheetToMarkdown(worksheet, progress);

                        // 命名規則：{Excel主檔名}_{工作表名稱}.md
                        string safeSheetName = SanitizeFileName(sheetName);
                        string outputFileName = $"{excelBaseName}_{safeSheetName}.md";
                        string outputFilePath = Path.Combine(outputDir, outputFileName);

                        // 將 Markdown 寫出為 .md 檔案（UTF-8 含 BOM，確保各編輯器相容性）
                        File.WriteAllText(
                            outputFilePath,
                            markdownContent,
                            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

                        progress.Report($"  ✔ 工作表【{sheetName}】已輸出：{outputFilePath}");

                        results.Add(new ExcelConversionResult
                        {
                            IsSuccess = true,
                            SheetName = sheetName,
                            OutputFilePath = outputFilePath
                        });
                    }
                    catch (Exception sheetEx)
                    {
                        progress.Report($"  ✘ 工作表【{sheetName}】轉換失敗：{sheetEx.Message}");
                        results.Add(new ExcelConversionResult
                        {
                            IsSuccess = false,
                            SheetName = sheetName,
                            ErrorMessage = $"工作表【{sheetName}】轉換失敗：{sheetEx.Message}"
                        });
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                // 捕捉所有未預期的例外，確保程式不會崩潰，並回傳友善的繁體中文錯誤訊息
                results.Add(new ExcelConversionResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"轉換過程發生例外：{ex.GetType().Name} - {ex.Message}"
                });
                return results;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
        });
    }

    /// <summary>
    /// 將單一 Excel 工作表的使用範圍轉換為 GFM 標準 Markdown 表格字串。
    /// </summary>
    /// <param name="worksheet">Excel COM Interop 的工作表物件。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>符合 GFM 語法的 Markdown 表格字串；若工作表無資料則傳回空字串提示。</returns>
    private static string ConvertWorksheetToMarkdown(
        Microsoft.Office.Interop.Excel.Worksheet worksheet,
        IProgress<string> progress)
    {
        // 取得工作表的使用範圍（UsedRange）
        Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;

        int rowCount = usedRange.Rows.Count;
        int colCount = usedRange.Columns.Count;

        // 若工作表完全沒有資料，直接回傳提示文字
        if (rowCount == 0 || colCount == 0)
        {
            progress.Report($"    ⚠ 工作表【{worksheet.Name}】無任何資料列，輸出空白提示。");
            return $"*（工作表【{worksheet.Name}】無資料）*{Environment.NewLine}";
        }

        // 一次讀取整個使用範圍的值到二維陣列（效能最佳化，避免逐格呼叫 COM）
        // 若範圍只有一個儲存格，Value2 回傳單一值而非陣列，需特別處理
        object?[,] values;
        if (rowCount == 1 && colCount == 1)
        {
            values = new object?[2, 2]; // COM 陣列以 1 為起始索引
            values[1, 1] = usedRange.Value2;
        }
        else
        {
            values = (object?[,])usedRange.Value2;
        }

        // 檢查是否所有格子都是空白
        bool hasData = false;
        for (int r = 1; r <= rowCount && !hasData; r++)
            for (int c = 1; c <= colCount && !hasData; c++)
                if (values[r, c] != null)
                    hasData = true;

        if (!hasData)
        {
            progress.Report($"    ⚠ 工作表【{worksheet.Name}】無任何資料列，輸出空白提示。");
            return $"*（工作表【{worksheet.Name}】無資料）*{Environment.NewLine}";
        }

        var sb = new StringBuilder();

        // ── 表頭列處理（第 1 列）──
        sb.Append('|');
        for (int col = 1; col <= colCount; col++)
        {
            string cellValue = FormatCellValue(values[1, col]);
            sb.Append($" {EscapeMarkdownCell(cellValue)} |");
        }
        sb.AppendLine();

        // ── 分隔線列 ──
        sb.Append('|');
        for (int col = 1; col <= colCount; col++)
        {
            sb.Append("---|");
        }
        sb.AppendLine();

        // ── 資料列處理（第 2 列起）──
        for (int row = 2; row <= rowCount; row++)
        {
            sb.Append('|');
            for (int col = 1; col <= colCount; col++)
            {
                string cellValue = FormatCellValue(values[row, col]);
                sb.Append($" {EscapeMarkdownCell(cellValue)} |");
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// 將 COM Value2 傳回的儲存格值格式化為字串。
    /// </summary>
    /// <param name="value">儲存格值（可能為 null、double、string 等）。</param>
    /// <returns>格式化後的字串。</returns>
    internal static string FormatCellValue(object? value)
    {
        if (value == null) return " ";
        return value.ToString() ?? " ";
    }

    /// <summary>
    /// 跳脫 Markdown 表格儲存格中的特殊字元。
    /// </summary>
    /// <param name="value">儲存格原始字串值。</param>
    /// <returns>跳脫特殊字元後的字串；若輸入為空則傳回空白字串。</returns>
    internal static string EscapeMarkdownCell(string value)
    {
        if (string.IsNullOrEmpty(value)) return " ";

        // 管線符號（|）在 Markdown 表格中作為欄位分隔符，需以反斜線跳脫
        // 換行字元會破壞表格結構，以空格取代
        return value
            .Replace("|", @"\|")
            .Replace("\r\n", " ")
            .Replace("\n", " ")
            .Replace("\r", " ");
    }

    /// <summary>
    /// 移除字串中對 Windows/Linux 檔案系統不合法的字元，用於安全化檔案名稱。
    /// </summary>
    /// <param name="name">原始字串（通常為工作表名稱）。</param>
    /// <returns>移除不合法字元後的安全檔案名稱；若結果為空則傳回 "Sheet"。</returns>
    internal static string SanitizeFileName(string name)
    {
        // 取得系統定義的所有不合法路徑字元，並替換為底線
        char[] invalidChars = Path.GetInvalidFileNameChars();
        string sanitized = string.Concat(name.Select(c => invalidChars.Contains(c) ? '_' : c));
        return string.IsNullOrWhiteSpace(sanitized) ? "Sheet" : sanitized;
    }
}
