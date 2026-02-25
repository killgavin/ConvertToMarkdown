using System.Text;
using System.Text.RegularExpressions;
using HapDoc = HtmlAgilityPack.HtmlDocument;
using HapNode = HtmlAgilityPack.HtmlNode;
using Mammoth;
using ReverseMarkdown;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WdSaveFormat = Microsoft.Office.Interop.Word.WdSaveFormat;

namespace ConvertToMarkdown;

/// <summary>
/// 轉換服務實作 - 負責將 Word (.docx / .doc) 檔案轉換為 GitHub Flavored Markdown (GFM) 格式。
/// 轉換流程：（若為 .doc 則先透過 Word COM Interop 轉為 .docx）→ Mammoth 解析 Docx → 提取圖片 → 正規化表格 → ReverseMarkdown 產生 .md。
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
            string? tempDocxPath = null;
            try
            {
                // === 步驟 0（條件性）：若為 .doc 格式，透過 Word COM Interop 轉為臨時 .docx ===
                bool isDoc = sourceFilePath.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)
                          && !sourceFilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);
                if (isDoc)
                {
                    progress.Report("▶ [0/6] 偵測到 .doc 格式，使用 Word COM Interop 轉換為 .docx...");
                    tempDocxPath = ConvertDocToDocx(sourceFilePath, progress);
                    sourceFilePath = tempDocxPath;
                }

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

                // === 步驟 2：建立輸出資料夾（以來源檔案主檔名命名）===
                // 在來源檔案所在目錄下建立與主檔名同名的子資料夾，用於存放 .md 及圖檔
                string sourceDir = Path.GetDirectoryName(sourceFilePath)!;
                string fileBaseName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputDir = Path.Combine(sourceDir, fileBaseName);
                Directory.CreateDirectory(outputDir);
                progress.Report($"▶ [2/6] 輸出資料夾：{outputDir}");

                // === 步驟 3：使用 Mammoth 將 Word 解析為 HTML ===
                // Mammoth 預設會將嵌入圖片轉為 base64 資料 URI 嵌入 HTML <img> 標籤
                progress.Report("▶ [3/6] Mammoth 解析 Word 文件中...");
                var mammothConverter = new DocumentConverter();
                var mammothResult = mammothConverter.ConvertToHtml(sourceFilePath);
                string html = mammothResult.Value;

                // 輸出 Mammoth 轉換過程中產生的所有警告
                foreach (var warning in mammothResult.Warnings)
                    progress.Report($"  ⚠ Mammoth 警告：{warning}");

                // === 步驟 4：從 HTML 提取 base64 圖片並儲存為獨立圖檔 ===
                // 將 base64 資料 URI 取代為相對路徑，供 Markdown 引用
                progress.Report("▶ [4/6] 提取並儲存圖片資源...");
                int imageCount = 0;
                html = ExtractAndSaveImages(html, outputDir, fileBaseName, ref imageCount, progress);
                progress.Report($"  ✔ 共提取 {imageCount} 張圖片");

                // === 步驟 5：正規化 HTML 表格 ===
                // Markdown 表格語法不支援合併儲存格 (colspan/rowspan)
                // 需先將合併儲存格展開為標準格子，再進行 Markdown 轉換
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
                // 清除預轉產生的臨時 .docx 檔案
                if (tempDocxPath != null && File.Exists(tempDocxPath))
                {
                    try { File.Delete(tempDocxPath); }
                    catch { /* 忽略清除失敗 */ }
                }
            }
        });
    }

    /// <summary>
    /// 透過 Microsoft Word COM Interop 將 .doc 檔案轉換為臨時 .docx 檔案。
    /// 需要使用者電腦已安裝 Microsoft Word。
    /// </summary>
    /// <param name="docPath">.doc 檔案的完整路徑。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>轉換產生的臨時 .docx 檔案路徑。</returns>
    /// <exception cref="InvalidOperationException">無法啟動 Microsoft Word 時擲出。</exception>
    private static string ConvertDocToDocx(string docPath, IProgress<string> progress)
    {
        WordApp? wordApp = null;
        Microsoft.Office.Interop.Word.Document? wordDoc = null;
        try
        {
            wordApp = new WordApp { Visible = false, DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone };
            wordDoc = wordApp.Documents.Open(docPath, ReadOnly: true);

            string tempDocxPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(docPath)}_{Guid.NewGuid():N}.docx");
            wordDoc.SaveAs2(tempDocxPath, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: (int)Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2013);

            progress.Report($"  ✔ .doc → .docx 轉換完成：{Path.GetFileName(tempDocxPath)}");
            return tempDocxPath;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw new InvalidOperationException(
                $"無法啟動 Microsoft Word 進行 .doc 轉換。請確認已安裝 Microsoft Word。\n詳細錯誤：{ex.Message}", ex);
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
    /// 掃描 HTML 字串中所有以 base64 資料 URI 嵌入的圖片，
    /// 將圖片二進位資料儲存為實體圖檔，並以相對路徑取代 src 屬性值。
    /// </summary>
    /// <param name="html">Mammoth 產生的原始 HTML 字串（含 base64 圖片）。</param>
    /// <param name="outputDir">圖片輸出目錄的完整路徑。</param>
    /// <param name="fileBaseName">來源 Word 檔案的主檔名，用於圖片命名前綴。</param>
    /// <param name="imageCount">圖片計數器（傳址參數），函式結束後記錄已儲存的圖片總數。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>圖片 src 已替換為相對路徑的 HTML 字串。</returns>
    private static string ExtractAndSaveImages(
        string html,
        string outputDir,
        string fileBaseName,
        ref int imageCount,
        IProgress<string> progress)
    {
        // 正規表示式比對格式：src="data:image/TYPE;base64,DATA"
        var imagePattern = new Regex(
            @"src=""data:image/(?<type>[^;]+);base64,(?<data>[^""]+)""",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // 使用暫存變數處理 ref 參數在 Lambda 中的限制
        int localCount = imageCount;

        string result = imagePattern.Replace(html, match =>
        {
            localCount++;
            string imageType = match.Groups["type"].Value.ToLowerInvariant();
            string base64Data = match.Groups["data"].Value;

            // 根據 MIME 子類型決定圖片副檔名
            string extension = imageType switch
            {
                "jpeg" or "jpg" => "jpg",
                "png"           => "png",
                "gif"           => "gif",
                "webp"          => "webp",
                "svg+xml"       => "svg",
                _               => imageType.Replace("+", "_")
            };

            // 命名規則：{主檔名}_圖片_{序號三位數}.{副檔名}
            string imageFileName = $"{fileBaseName}_圖片_{localCount:D3}.{extension}";
            string imageFilePath = Path.Combine(outputDir, imageFileName);

            try
            {
                // 解碼 base64 並以高品質（原始二進位）寫入磁碟
                byte[] imageBytes = Convert.FromBase64String(base64Data);
                File.WriteAllBytes(imageFilePath, imageBytes);
                progress.Report($"  ✔ 圖片已儲存：{imageFileName}（{imageBytes.Length:N0} bytes）");
            }
            catch (Exception ex)
            {
                progress.Report($"  ✘ 圖片儲存失敗（第 {localCount} 張）：{ex.Message}");
                // 儲存失敗時，保留原始 base64 src，避免圖片遺失
                return match.Value;
            }

            // 以相對路徑取代 base64 資料 URI（Markdown 引用同目錄圖檔）
            return $@"src=""{imageFileName}""";
        });

        imageCount = localCount;
        return result;
    }

    /// <summary>
    /// 掃描 HTML 中所有 &lt;table&gt; 元素，將含有 colspan/rowspan 的合併儲存格展開，
    /// 使表格符合標準 Markdown 表格語法（每行欄數一致、無合併儲存格）。
    /// </summary>
    /// <param name="html">含表格的原始 HTML 字串。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>所有表格已展開合併儲存格的 HTML 字串。</returns>
    private static string NormalizeTables(string html, IProgress<string> progress)
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
    private static void FlattenTable(HapNode tableNode)
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
    private static int ParseSpanAttr(HapNode node, string attrName)
    {
        string attrValue = node.GetAttributeValue(attrName, "1");
        return int.TryParse(attrValue, out int result) && result > 0 ? result : 1;
    }
}
