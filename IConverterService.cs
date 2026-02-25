namespace ConvertToMarkdown;

/// <summary>
/// 轉換服務介面 - 定義將 Word 檔案轉換為 Markdown 的行為合約。
/// 實作此介面的類別需處理圖片提取、表格正規化等轉換工作。
/// </summary>
public interface IConverterService
{
    /// <summary>
    /// 非同步將指定的 Word (.docx) 檔案轉換為 Markdown 格式。
    /// </summary>
    /// <param name="sourceFilePath">來源 Word 檔案的完整路徑。</param>
    /// <param name="progress">
    /// 進度回報介面，用於將執行日誌訊息回傳給呼叫端（通常是 UI 執行緒）。
    /// </param>
    /// <returns>
    /// 非同步工作，完成後傳回 <see cref="ConversionResult"/>，
    /// 包含是否成功、輸出路徑及錯誤訊息等資訊。
    /// </returns>
    Task<ConversionResult> ConvertAsync(string sourceFilePath, IProgress<string> progress);
}

/// <summary>
/// 轉換結果 - 封裝單次轉換操作的執行結果。
/// </summary>
public class ConversionResult
{
    /// <summary>取得或設定轉換是否成功完成。</summary>
    public bool IsSuccess { get; init; }

    /// <summary>取得或設定輸出的 Markdown 檔案完整路徑。</summary>
    public string OutputFilePath { get; init; } = string.Empty;

    /// <summary>取得或設定轉換失敗時的錯誤訊息。</summary>
    public string ErrorMessage { get; init; } = string.Empty;
}
