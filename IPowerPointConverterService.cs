namespace ConvertToMarkdown;

/// <summary>
/// PowerPoint 轉換服務介面 - 定義將 PowerPoint 檔案各投影片轉換為 Markdown 的行為合約。
/// </summary>
public interface IPowerPointConverterService
{
    /// <summary>
    /// 非同步將指定的 PowerPoint (.pptx / .ppt) 檔案轉換為 Markdown 檔案。
    /// </summary>
    /// <param name="sourceFilePath">來源 PowerPoint 檔案的完整路徑。</param>
    /// <param name="progress">
    /// 進度回報介面，用於將執行日誌訊息回傳給呼叫端（通常是 UI 執行緒）。
    /// </param>
    /// <returns>
    /// 非同步工作，完成後傳回 <see cref="ConversionResult"/>，
    /// 包含是否成功、輸出路徑及錯誤訊息等資訊。
    /// </returns>
    Task<ConversionResult> ConvertAsync(string sourceFilePath, IProgress<string> progress);
}
