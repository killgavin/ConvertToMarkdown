namespace ConvertToMarkdown;

/// <summary>
/// 程式進入點 - 初始化 WinForms 應用程式並啟動主視窗。
/// </summary>
static class Program
{
    /// <summary>
    /// 應用程式主要進入點（單執行緒 Apartment，WinForms 必要設定）。
    /// </summary>
    [STAThread]
    static void Main()
    {
        // 註冊字碼頁編碼提供者，供 ExcelDataReader 讀取 .xls (BIFF) 格式時使用
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 套用應用程式組態（高 DPI、視覺樣式等預設設定）
        ApplicationConfiguration.Initialize();

        // 啟動主視窗（Word 轉 Markdown 工具）
        Application.Run(new MainForm());
    }
}