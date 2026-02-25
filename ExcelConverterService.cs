using System.Data;
using System.Text;
using ExcelDataReader;

namespace ConvertToMarkdown;

/// <summary>
/// Excel 轉換服務實作 - 負責將 Excel (.xlsx / .xls) 檔案的每個工作表
/// 轉換為 GitHub Flavored Markdown (GFM) 標準表格格式，並各自輸出為獨立 .md 檔案。
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

                // === 步驟 2：建立輸出資料夾 MD ===
                // 在來源檔案所在目錄下建立 MD 子資料夾，用於存放各工作表的 .md 檔案
                string sourceDir = Path.GetDirectoryName(sourceFilePath)!;
                string excelBaseName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputDir = Path.Combine(sourceDir, "MD");
                Directory.CreateDirectory(outputDir);
                progress.Report($"▶ [2/4] 輸出資料夾：{outputDir}");

                // === 步驟 3：使用 ExcelDataReader 讀取所有工作表 ===
                // ExcelDataReader 要求先註冊字碼頁編碼提供者（支援 .xls 的 BIFF 格式）
                // 此呼叫幂等，重複呼叫無副作用
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                progress.Report("▶ [3/4] ExcelDataReader 讀取工作表中...");

                DataSet dataSet;
                try
                {
                    // 以共用讀取模式開啟 Excel 檔案，允許其他程式同時讀取（但若被獨占寫入鎖定則仍會失敗）
                    using var stream = new FileStream(
                        sourceFilePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite);

                    // 根據副檔名選擇對應的讀取器：
                    // .xlsx → OpenXml 格式；.xls → BIFF (Binary Interchange File Format) 格式
                    using IExcelDataReader reader = Path.GetExtension(sourceFilePath)
                        .Equals(".xls", StringComparison.OrdinalIgnoreCase)
                        ? ExcelReaderFactory.CreateBinaryReader(stream)
                        : ExcelReaderFactory.CreateOpenXmlReader(stream);

                    // 將 Excel 全部工作表讀入 DataSet，
                    // UseHeaderRow = false 表示不自動以第一列作為欄位名稱，
                    // 改由後續邏輯手動處理表頭（第一列視為表頭）
                    var config = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false
                        }
                    };
                    dataSet = reader.AsDataSet(config);
                }
                catch (IOException ioEx)
                {
                    // 檔案被其他程式開啟中（共用違規）或其他 I/O 錯誤
                    results.Add(new ExcelConversionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"無法開啟 Excel 檔案，請確認檔案未被其他程式佔用。\n詳細錯誤：{ioEx.Message}"
                    });
                    return results;
                }

                // === 步驟 4：逐一轉換每個工作表為 Markdown 表格 ===
                progress.Report($"▶ [4/4] 共偵測到 {dataSet.Tables.Count} 個工作表，開始逐一轉換...");

                // 遍歷 DataSet 中的每一個 DataTable（對應 Excel 的每一個工作表）
                foreach (DataTable table in dataSet.Tables)
                {
                    // 取得工作表名稱（即 DataTable.TableName，由 ExcelDataReader 自動帶入）
                    string sheetName = table.TableName;
                    progress.Report($"  ▷ 正在處理工作表：【{sheetName}】");

                    try
                    {
                        // 將 DataTable 轉換為 GFM Markdown 表格字串
                        string markdownContent = ConvertDataTableToMarkdown(table, progress);

                        // 命名規則：{Excel主檔名}_{工作表名稱}.md
                        // 移除工作表名稱中對檔案系統不合法的字元，避免建檔失敗
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
        });
    }

    /// <summary>
    /// 將單一 DataTable（對應一個工作表）的內容轉換為 GFM 標準 Markdown 表格字串。
    /// </summary>
    /// <remarks>
    /// 轉換規則：
    /// <list type="bullet">
    /// <item>
    ///   <term>表頭列（Header Row）</term>
    ///   <description>
    ///     預設以 DataTable 中索引為 0 的列（即 Excel 第一列）作為 Markdown 表頭（Header）。
    ///     表頭列輸出後，緊接一行分隔線（由 <c>|---|</c> 組成），符合 GFM 表格語法。
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>資料列（Data Rows）</term>
    ///   <description>
    ///     從索引 1 開始遍歷，每列輸出為一行 <c>| 欄位1 | 欄位2 | ... |</c> 格式。
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>空值處理</term>
    ///   <description>
    ///     若儲存格（Cell）為 null 或空值，輸出空白字串（<c>" "</c>），
    ///     確保 Markdown 表格每列欄數一致、對齊正確。
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>特殊字元跳脫</term>
    ///   <description>
    ///     儲存格內容中若含有管線符號（<c>|</c>），需以 <c>\|</c> 跳脫，
    ///     避免破壞 Markdown 表格結構。
    ///   </description>
    /// </item>
    /// </list>
    /// </remarks>
    /// <param name="table">要轉換的工作表資料（DataTable）。</param>
    /// <param name="progress">進度回報介面。</param>
    /// <returns>符合 GFM 語法的 Markdown 表格字串；若工作表無資料則傳回空字串提示。</returns>
    private static string ConvertDataTableToMarkdown(DataTable table, IProgress<string> progress)
    {
        // 若工作表完全沒有列（Row），直接回傳提示文字
        if (table.Rows.Count == 0)
        {
            progress.Report($"    ⚠ 工作表【{table.TableName}】無任何資料列，輸出空白提示。");
            return $"*（工作表【{table.TableName}】無資料）*{Environment.NewLine}";
        }

        // 取得工作表的欄位數量（Column count）
        // DataTable 的欄位數量由讀取到的最大欄數決定
        int columnCount = table.Columns.Count;

        var sb = new StringBuilder();

        // ── 表頭列處理 ──────────────────────────────────────────────────────
        // 取出 DataTable 中的第一列（Row index = 0）作為 Markdown 表頭
        DataRow headerRow = table.Rows[0];

        // 逐欄讀取表頭儲存格內容，並進行特殊字元跳脫後，組成表頭行
        sb.Append('|');
        for (int col = 0; col < columnCount; col++)
        {
            // 取得儲存格值；若為 DBNull（空儲存格）則以空白字串代替，確保欄位對齊
            string cellValue = headerRow[col] == DBNull.Value
                ? " "
                : EscapeMarkdownCell(headerRow[col]?.ToString() ?? " ");

            sb.Append($" {cellValue} |");
        }
        sb.AppendLine();

        // ── 分隔線列（Separator Row）────────────────────────────────────────
        // GFM 規範要求在表頭列與資料列之間插入分隔線，格式為 |---|---|...
        sb.Append('|');
        for (int col = 0; col < columnCount; col++)
        {
            sb.Append("---|");
        }
        sb.AppendLine();

        // ── 資料列處理 ──────────────────────────────────────────────────────
        // 從索引 1 開始遍歷，跳過已作為表頭的第一列
        // DataTable.Rows 是以索引存取的集合，foreach 無法直接跳過第一列，故使用 for 迴圈
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            DataRow dataRow = table.Rows[rowIdx];

            // 逐欄讀取每個儲存格（Cell）的值，組成 Markdown 表格的一行
            sb.Append('|');
            for (int col = 0; col < columnCount; col++)
            {
                // 儲存格為 DBNull（即 Excel 空白儲存格）時，填入空白字串
                // 確保輸出的 Markdown 表格在視覺上對齊，且符合 GFM 語法規範
                string cellValue = dataRow[col] == DBNull.Value
                    ? " "
                    : EscapeMarkdownCell(dataRow[col]?.ToString() ?? " ");

                sb.Append($" {cellValue} |");
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// 跳脫 Markdown 表格儲存格中的特殊字元。
    /// </summary>
    /// <param name="value">儲存格原始字串值。</param>
    /// <returns>跳脫特殊字元後的字串；若輸入為空則傳回空白字串。</returns>
    private static string EscapeMarkdownCell(string value)
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
    private static string SanitizeFileName(string name)
    {
        // 取得系統定義的所有不合法路徑字元，並替換為底線
        char[] invalidChars = Path.GetInvalidFileNameChars();
        string sanitized = string.Concat(name.Select(c => invalidChars.Contains(c) ? '_' : c));
        return string.IsNullOrWhiteSpace(sanitized) ? "Sheet" : sanitized;
    }
}
