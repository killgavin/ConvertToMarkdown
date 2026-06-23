using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using HapDoc = HtmlAgilityPack.HtmlDocument;
using HapNode = HtmlAgilityPack.HtmlNode;
using ReverseMarkdown;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WdSaveFormat = Microsoft.Office.Interop.Word.WdSaveFormat;

namespace ConvertToMarkdown;

/// <summary>
/// 轉換服務實作 - 負責將 Word (.docx / .doc) 檔案轉換為 GitHub Flavored Markdown (GFM) 格式。
/// 轉換流程：Word COM Interop 開啟文件 → 匯出為篩選後 HTML → 搬移圖片 → 正規化表格 → ReverseMarkdown 產生 .md。
/// </summary>
public class ConverterService : IConverterService
{
    /// <summary>
    /// 非同步執行 Word 轉 Markdown 完整流程。
    /// 整個轉換工作會在背景執行緒執行，以避免 UI 凍結 (UI Freeze)。
    /// </summary>
    /// <param name="sourceFilePath">來源 Word 檔案的完整路徑。</param>
    /// <param name="progress">進度回報介面，用於向 UI 輸出執行日誌。</param>
    /// <returns>
    /// 非同步工作，完成後傳回 <see cref="ConversionResult"/>。
    /// </returns>
    public async Task<ConversionResult> ConvertAsync(string sourceFilePath, IProgress<string> progress)
    {
        // 使用 Task.Run 將整個 I/O 密集工作移至執行緒集區，確保 UI 執行緒不被阻塞
        return await Task.Run(() =>
        {
            // 暫存 Word Interop 匯出的 HTML 檔案路徑及對應圖片子資料夾路徑
            string? tempHtmlPath = null;
            string? tempHtmlFilesDir = null;
            try
            {
                // === 步驟 1：驗證來源檔案是否存在 ===
                progress.Report("▶ [1/6] 驗證來源檔案...");
                if (!File.Exists(sourceFilePath))
                {
                    return new ConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"找不到指定的檔案：{sourceFilePath}"
                    };
                }

                // === 步驟 2：建立輸出資料夾（依來源檔名命名） ===
                string sourceDir = Path.GetDirectoryName(sourceFilePath)!;
                string fileBaseName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputDir = Path.Combine(sourceDir, fileBaseName);
                Directory.CreateDirectory(outputDir);
                progress.Report($"▶ [2/6] 輸出資料夾：{outputDir}");

                // === 步驟 3：透過 Word COM Interop 將文件匯出為篩選後 HTML ===
                progress.Report("▶ [3/6] Word COM Interop 匯出 HTML 中...");
                tempHtmlPath = ConvertWordToHtml(sourceFilePath, progress);
                // Word 匯出 HTML 時，圖片會存放在 {檔名}_files 子資料夾中
                tempHtmlFilesDir = Path.Combine(
                    Path.GetDirectoryName(tempHtmlPath)!,
                    Path.GetFileNameWithoutExtension(tempHtmlPath) + "_files");

                string html = File.ReadAllText(tempHtmlPath, Encoding.UTF8);

                // === 步驟 4：搬移圖片資源至輸出資料夾 ===
                progress.Report("▶ [4/6] 搬移圖片資源...");
                int imageCount = 0;
                html = MoveAndRelinkImages(html, tempHtmlFilesDir, outputDir, ref imageCount, progress);
                progress.Report($"  ✔ 共搬移 {imageCount} 張圖片");

                // === 步驟 5：正規化 HTML 表格 ===
                progress.Report("▶ [5/6] 展開表格合併儲存格...");
                html = NormalizeTables(html, progress);

                // === 步驟 6：使用 ReverseMarkdown 轉換為 GFM 格式 ===
                progress.Report("▶ [6/6] ReverseMarkdown 產生 Markdown 內容...");
                var mdConfig = new Config
                {
                    GithubFlavored = true,
                    RemoveComments = true,
                    SmartHrefHandling = true
                };
                var mdConverter = new Converter(mdConfig);
                string markdown = mdConverter.Convert(html);

                // 將 Markdown 寫出為 .md 檔案（UTF-8 含 BOM，確保各編輯器相容性）
                string outputFilePath = Path.Combine(outputDir, $"{fileBaseName}.md");
                File.WriteAllText(outputFilePath, markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                progress.Report($"  ✔ Markdown 檔案已產出：{outputFilePath}");

                return new ConversionResult
                {
                    IsSuccess = true,
                    OutputFilePath = outputFilePath
                };
            }
            catch (Exception ex)
            {
                // 捕捉所有未預期的例外，確保程式不會崩潰，並回傳友善的錯誤訊息
                return new ConversionResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"轉換過程發生例外：{ex.GetType().Name} - {ex.Message}"
                };
            }
            finally
            {
                // 清除 Word Interop 匯出的暫存 HTML 檔案及其圖片子資料夾
                if (tempHtmlPath != null && File.Exists(tempHtmlPath))
                {
                    try { File.Delete(tempHtmlPath); }
                    catch { /* 忽略清除失敗 */ }
                }
                if (tempHtmlFilesDir != null && Directory.Exists(tempHtmlFilesDir))
                {
                    try { Directory.Delete(tempHtmlFilesDir, recursive: true); }
                    catch { /* 忽略清除失敗 */ }
                }
            }
        });
    }

    /// <summary>
    /// 透過 Microsoft Word COM Interop 將 .doc / .docx 檔案匯出為篩選後 HTML。
    /// 需要使用者電腦已安裝 Microsoft Word。
    /// </summary>
    /// <param name="wordPath">Word 檔案的完整路徑（.doc 或 .docx）。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>匯出的暫存 HTML 檔案路徑。</returns>
    /// <exception cref="InvalidOperationException">無法啟動 Microsoft Word 時擲出。</exception>
    private static string ConvertWordToHtml(string wordPath, IProgress<string> progress)
    {
        WordApp? wordApp = null;
        Microsoft.Office.Interop.Word.Document? wordDoc = null;
        try
        {
            wordApp = new WordApp { Visible = false, DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone };
            wordDoc = wordApp.Documents.Open(wordPath, ReadOnly: true);

            // 以篩選後 HTML 格式匯出，會產生較簡潔的 HTML，並將圖片存於 _files 子資料夾
            string tempHtmlPath = Path.Combine(
                Path.GetTempPath(),
                $"{Path.GetFileNameWithoutExtension(wordPath)}_{Guid.NewGuid():N}.html");

            wordDoc.SaveAs2(tempHtmlPath, WdSaveFormat.wdFormatFilteredHTML);

            progress.Report($"  ✔ Word → HTML 匯出完成：{Path.GetFileName(tempHtmlPath)}");
            return tempHtmlPath;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw new InvalidOperationException(
                $"無法啟動 Microsoft Word 進行轉換。請確認已安裝 Microsoft Word。\n詳細錯誤：{ex.Message}", ex);
        }
        finally
        {
            if (wordDoc != null)
            {
                wordDoc.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
            }
            if (wordApp != null)
            {
                wordApp.Quit(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }
    }

    /// <summary>
    /// 將 Word Interop 匯出 HTML 時產生的圖片檔案搬移至最終輸出目錄，
    /// 並更新 HTML 中 &lt;img&gt; 標籤的 src 為以內容雜湊命名的相對路徑。
    /// 同時處理仍以 base64 data URI 嵌入的圖片（部分版本的 Word 可能內嵌圖片）。
    /// </summary>
    /// <param name="html">Word Interop 匯出的原始 HTML 字串。</param>
    /// <param name="htmlFilesDir">Word 匯出 HTML 時產生的圖片子資料夾路徑。</param>
    /// <param name="outputDir">最終輸出目錄的完整路徑。</param>
    /// <param name="imageCount">圖片計數器（傳址參數）。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>圖片 src 已替換為相對路徑的 HTML 字串。</returns>
    private static string MoveAndRelinkImages(
        string html,
        string htmlFilesDir,
        string outputDir,
        ref int imageCount,
        IProgress<string> progress)
    {
        int localCount = imageCount;

        // ── 處理 Word 匯出的外部圖片檔案參照 ──
        // Word 匯出 HTML 時，<img src="XXX_files/image001.png"> 的格式
        if (Directory.Exists(htmlFilesDir))
        {
            string filesDirName = Path.GetFileName(htmlFilesDir);
            var fileRefPattern = new Regex(
                $@"src=""(?:{Regex.Escape(filesDirName)}/|(?:\./)?{Regex.Escape(filesDirName)}/)(?<filename>[^""]+)""",
                RegexOptions.IgnoreCase);

            html = fileRefPattern.Replace(html, match =>
            {
                string originalFileName = match.Groups["filename"].Value;
                string originalFilePath = Path.Combine(htmlFilesDir, originalFileName);

                if (!File.Exists(originalFilePath))
                {
                    progress.Report($"  ⚠ 圖片檔案不存在，跳過：{originalFileName}");
                    return match.Value;
                }

                try
                {
                    byte[] imageBytes = File.ReadAllBytes(originalFilePath);
                    string hash = Convert.ToHexString(SHA256.HashData(imageBytes))[..16].ToLowerInvariant();
                    string extension = Path.GetExtension(originalFileName).TrimStart('.').ToLowerInvariant();
                    if (string.IsNullOrEmpty(extension)) extension = "png";

                    string newFileName = $"img_{hash}.{extension}";
                    string newFilePath = Path.Combine(outputDir, newFileName);

                    if (!File.Exists(newFilePath))
                    {
                        File.Copy(originalFilePath, newFilePath);
                        progress.Report($"  ✔ 圖片已搬移：{originalFileName} → {newFileName}（{imageBytes.Length:N0} bytes）");
                    }
                    else
                    {
                        progress.Report($"  ✔ 圖片已存在（重複引用）：{newFileName}");
                    }

                    localCount++;
                    return $@"src=""{newFileName}""";
                }
                catch (Exception ex)
                {
                    progress.Report($"  ✘ 圖片搬移失敗：{originalFileName} - {ex.Message}");
                    return match.Value;
                }
            });
        }

        // ── 處理可能仍以 base64 data URI 嵌入的圖片 ──
        var base64Pattern = new Regex(
            @"src=""data:image/(?<type>[^;]+);base64,(?<data>[^""]+)""",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        html = base64Pattern.Replace(html, match =>
        {
            localCount++;
            string imageType = match.Groups["type"].Value.ToLowerInvariant();
            string base64Data = match.Groups["data"].Value;

            string extension = imageType switch
            {
                "jpeg" or "jpg" => "jpg",
                "png"           => "png",
                "gif"           => "gif",
                "webp"          => "webp",
                "svg+xml"       => "svg",
                _               => imageType.Replace("+", "_")
            };

            byte[] imageBytes;
            try
            {
                imageBytes = Convert.FromBase64String(base64Data);
            }
            catch (Exception ex)
            {
                progress.Report($"  ✘ 圖片解碼失敗（第 {localCount} 張）：{ex.Message}");
                return match.Value;
            }

            string hash = Convert.ToHexString(SHA256.HashData(imageBytes))[..16].ToLowerInvariant();
            string imageFileName = $"img_{hash}.{extension}";
            string imageFilePath = Path.Combine(outputDir, imageFileName);

            try
            {
                if (!File.Exists(imageFilePath))
                {
                    File.WriteAllBytes(imageFilePath, imageBytes);
                    progress.Report($"  ✔ 圖片已儲存：{imageFileName}（{imageBytes.Length:N0} bytes）");
                }
                else
                {
                    progress.Report($"  ✔ 圖片已存在（重複引用）：{imageFileName}");
                }
            }
            catch (Exception ex)
            {
                progress.Report($"  ✘ 圖片儲存失敗（第 {localCount} 張）：{ex.Message}");
                return match.Value;
            }

            return $@"src=""{imageFileName}""";
        });

        imageCount = localCount;
        return html;
    }

    /// <summary>
    /// 掃描 HTML 中所有 &lt;table&gt; 元素，將含有 colspan/rowspan 的合併儲存格展開，
    /// 使表格符合標準 Markdown 表格語法（每行欄數一致、無合併儲存格）。
    /// </summary>
    /// <param name="html">含表格的原始 HTML 字串。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>所有表格已展開合併儲存格的 HTML 字串。</returns>
    internal static string NormalizeTables(string html, IProgress<string> progress)
    {
        var doc = new HapDoc();
        doc.LoadHtml(html);

        // 使用 XPath 選取文件中所有 <table> 節點
        var tables = doc.DocumentNode.SelectNodes("//table");
        if (tables == null) return html;

        int tableCount = 0;
        // 注意：SelectNodes 傳回的是即時清單，展開操作可能影響遍歷，
        // 因此轉為陣列以確保遍歷穩定性
        foreach (var table in tables.ToArray())
        {
            FlattenTable(table);
            tableCount++;
        }

        if (tableCount > 0)
            progress.Report($"  ✔ 已正規化 {tableCount} 個表格");

        return doc.DocumentNode.OuterHtml;
    }

    /// <summary>
    /// 將單一 HTML 表格節點中的合併儲存格（colspan/rowspan）展開為標準格子。
    /// </summary>
    /// <remarks>
    /// 處理步驟：
    /// <list type="number">
    /// <item>掃描各行儲存格，以虛擬格子矩陣 (grid) 記錄每個邏輯位置的內容。</item>
    /// <item>colspan 展開為水平多欄；rowspan 展開為垂直多行。</item>
    /// <item>以 HtmlAgilityPack API 重建表格節點，移除所有 colspan/rowspan 屬性。</item>
    /// </list>
    /// </remarks>
    /// <param name="tableNode">待正規化的 HtmlAgilityPack 表格節點。</param>
    internal static void FlattenTable(HapNode tableNode)
    {
        // 取得表格內所有 <tr> 行節點（含 <thead>、<tbody> 下的 <tr>）
        var rows = tableNode.SelectNodes(".//tr");
        if (rows == null || rows.Count == 0) return;

        // grid：虛擬二維矩陣，索引鍵為 (列索引, 欄索引)，值為儲存格 InnerHtml
        var grid = new Dictionary<(int row, int col), string>();

        for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
        {
            // 取得該行的所有 <td> 與 <th> 儲存格節點
            var cells = rows[rowIdx].SelectNodes("td|th");
            if (cells == null) continue;

            int colIdx = 0;
            foreach (var cell in cells)
            {
                // 跳過已被上方列的 rowspan 佔用的欄位
                while (grid.ContainsKey((rowIdx, colIdx)))
                    colIdx++;

                // 解析 colspan（水平合併欄數）與 rowspan（垂直合併列數）
                int colspan = ParseSpanAttr(cell, "colspan");
                int rowspan = ParseSpanAttr(cell, "rowspan");

                // 取得儲存格的 HTML 內容（保留內部格式如粗體、連結等）
                string cellContent = cell.InnerHtml?.Trim() ?? string.Empty;

                // 將此儲存格涵蓋的所有 (列, 欄) 座標寫入虛擬矩陣
                // 規則：合併範圍的第一個格子保留內容，其餘補空白（Markdown 相容性）
                for (int r = 0; r < rowspan; r++)
                {
                    for (int c = 0; c < colspan; c++)
                    {
                        grid[(rowIdx + r, colIdx + c)] = (r == 0 && c == 0) ? cellContent : string.Empty;
                    }
                }

                colIdx += colspan;
            }
        }

        if (grid.Count == 0) return;

        // 計算重建表格所需的最大列數與欄數
        int maxRow = grid.Keys.Max(k => k.row);
        int maxCol = grid.Keys.Max(k => k.col);

        // 取得表格所屬的 HtmlDocument 實例，用於建立新節點
        var ownerDoc = tableNode.OwnerDocument;

        // 清除表格原有子節點，準備重建
        tableNode.RemoveAllChildren();

        for (int r = 0; r <= maxRow; r++)
        {
            // 建立 <tr> 行節點
            var trNode = ownerDoc.CreateElement("tr");

            for (int c = 0; c <= maxCol; c++)
            {
                // 第一行使用 <th>（表頭），其餘行使用 <td>（資料列）
                string tagName = r == 0 ? "th" : "td";
                var cellNode = ownerDoc.CreateElement(tagName);

                // 填入儲存格內容（展開後的單一內容，不含 colspan/rowspan）
                string content = grid.TryGetValue((r, c), out string? val) ? val : string.Empty;
                cellNode.InnerHtml = content;

                trNode.AppendChild(cellNode);
            }

            tableNode.AppendChild(trNode);
        }
    }

    /// <summary>
    /// 解析 HTML 節點屬性中的跨欄/跨列整數值。
    /// </summary>
    /// <param name="node">要讀取屬性的 HTML 節點。</param>
    /// <param name="attrName">屬性名稱，應為 <c>colspan</c> 或 <c>rowspan</c>。</param>
    /// <returns>
    /// 屬性的整數值（必須大於 0）；若屬性不存在或解析失敗，則傳回預設值 <c>1</c>。
    /// </returns>
    internal static int ParseSpanAttr(HapNode node, string attrName)
    {
        string attrValue = node.GetAttributeValue(attrName, "1");
        return int.TryParse(attrValue, out int result) && result > 0 ? result : 1;
    }
}
