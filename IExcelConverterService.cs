namespace ConvertToMarkdown;

/// <summary>
/// Excel 轉換服務介面 - 定義將 Excel 檔案各工作表轉換為 Markdown 的行為合約。
/// </summary>
public interface IExcelConverterService
{
    /// <summary>
    /// 非同步將指定的 Excel (.xlsx / .xls) 檔案的每個工作表轉換為獨立的 Markdown 檔案。
    /// </summary>
    /// <param name="sourceFilePath">來源 Excel 檔案的完整路徑。</param>
    /// <param name="progress">
    /// 進度回報介面，用於將執行日誌訊息回傳給呼叫端（通常是 UI 執行緒）。
    /// </param>
    /// <returns>
    /// 非同步工作，完成後傳回 <see cref="ExcelConversionResult"/> 清單，
    /// 每個工作表對應一筆結果記錄。
    /// </returns>
    Task<IReadOnlyList<ExcelConversionResult>> ConvertAsync(string sourceFilePath, IProgress<string> progress);
}

/// <summary>
/// Excel 工作表轉換結果 - 封裝單一工作表的轉換執行結果。
/// </summary>
public class ExcelConversionResult
{
    /// <summary>取得或設定轉換是否成功完成。</summary>
    public bool IsSuccess { get; init; }

    /// <summary>取得或設定工作表名稱。</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>取得或設定輸出的 Markdown 檔案完整路徑。</summary>
    public string OutputFilePath { get; init; } = string.Empty;

    /// <summary>取得或設定轉換失敗時的錯誤訊息。</summary>
    public string ErrorMessage { get; init; } = string.Empty;
}
