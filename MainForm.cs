namespace ConvertToMarkdown;

/// <summary>
/// 主視窗 - Word / Excel / PowerPoint 轉 Markdown 工具的使用者介面。
/// 提供「瀏覽檔案」、「開始轉換」功能及執行日誌顯示區塊。
/// </summary>
public partial class MainForm : Form
{
    /// <summary>Word 轉換服務實例，負責處理 Word 轉換邏輯。</summary>
    private readonly IConverterService _converterService;

    /// <summary>Excel 轉換服務實例，負責處理 Excel 多工作表轉換邏輯。</summary>
    private readonly IExcelConverterService _excelConverterService;

    /// <summary>PowerPoint 轉換服務實例，負責處理 PowerPoint 投影片轉換邏輯。</summary>
    private readonly IPowerPointConverterService _powerPointConverterService;

    /// <summary>
    /// 初始化主視窗，並建立轉換服務實例。
    /// </summary>
    public MainForm()
    {
        InitializeComponent();
        _converterService = new ConverterService();
        _excelConverterService = new ExcelConverterService();
        _powerPointConverterService = new PowerPointConverterService();
    }

    /// <summary>
    /// 「瀏覽...」按鈕的點擊事件處理方法。
    /// 開啟檔案選取對話方塊，讓使用者選取 Word (.docx) 檔案。
    /// </summary>
    private void BtnBrowse_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "選取 Word 檔案",
            Filter = "所有 Word 文件 (*.docx;*.doc)|*.docx;*.doc|Word 文件 (*.docx)|*.docx|Word 97-2003 文件 (*.doc)|*.doc",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtFilePath.Text = dialog.FileName;
            // 選取檔案後啟用「開始轉換」按鈕
            btnConvert.Enabled = true;
            AppendLog($"已選取檔案：{dialog.FileName}");
        }
    }

    /// <summary>
    /// 「開始轉換」按鈕的點擊事件處理方法（非同步）。
    /// 驗證輸入後，呼叫轉換服務執行 Word 轉 Markdown 工作。
    /// 整個轉換過程為非同步執行，UI 不會凍結。
    /// </summary>
    private async void BtnConvert_Click(object sender, EventArgs e)
    {
        string sourceFilePath = txtFilePath.Text.Trim();

        // 驗證使用者是否已選取檔案
        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            MessageBox.Show("請先選取要轉換的 Word 檔案。", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 驗證副檔名必須為 .docx 或 .doc
        string extension = Path.GetExtension(sourceFilePath);
        if (!extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
         && !extension.Equals(".doc", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("僅支援 .docx 及 .doc 格式的 Word 檔案。", "格式錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 轉換期間停用按鈕，避免重複觸發
        SetControlsEnabled(false);
        rtbLog.Clear();
        AppendLog("═══════════════════════════════════════");
        AppendLog("  Word 轉 Markdown 轉換工具  開始執行");
        AppendLog("═══════════════════════════════════════");

        // 建立進度回報物件，確保訊息在 UI 執行緒上更新
        var progress = new Progress<string>(AppendLog);

        // 非同步執行轉換，await 確保 UI 執行緒不被阻塞
        var result = await _converterService.ConvertAsync(sourceFilePath, progress);

        // 轉換完成後處理結果
        if (result.IsSuccess)
        {
            AppendLog("───────────────────────────────────────");
            AppendLog($"✔ 轉換成功！");
            AppendLog($"  輸出路徑：{result.OutputFilePath}");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"轉換完成！\n\n輸出路徑：\n{result.OutputFilePath}",
                "轉換成功",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        else
        {
            AppendLog("───────────────────────────────────────");
            AppendLog($"✘ 轉換失敗：{result.ErrorMessage}");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"轉換失敗：\n{result.ErrorMessage}",
                "轉換失敗",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        // 恢復按鈕狀態
        SetControlsEnabled(true);
    }

    /// <summary>
    /// 「瀏覽...」按鈕的點擊事件處理方法（Excel）。
    /// 開啟檔案選取對話方塊，讓使用者選取 Excel (.xlsx / .xls) 檔案。
    /// </summary>
    private void BtnBrowseExcel_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "選取 Excel 檔案",
            Filter = "所有 Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 活頁簿 (*.xls)|*.xls",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtExcelFilePath.Text = dialog.FileName;
            // 選取檔案後啟用「轉換 Excel」按鈕
            btnConvertExcel.Enabled = true;
            AppendLog($"已選取 Excel 檔案：{dialog.FileName}");
        }
    }

    /// <summary>
    /// 「轉換 Excel」按鈕的點擊事件處理方法（非同步）。
    /// 驗證輸入後，呼叫 Excel 轉換服務執行各工作表轉 Markdown 工作。
    /// 整個轉換過程為非同步執行，UI 不會凍結。
    /// </summary>
    private async void BtnConvertExcel_Click(object sender, EventArgs e)
    {
        string sourceFilePath = txtExcelFilePath.Text.Trim();

        // 驗證使用者是否已選取 Excel 檔案
        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            MessageBox.Show("請先選取要轉換的 Excel 檔案。", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 驗證副檔名必須為 .xlsx 或 .xls
        string extension = Path.GetExtension(sourceFilePath);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
         && !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("僅支援 .xlsx 及 .xls 格式的 Excel 檔案。", "格式錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 轉換期間停用按鈕，避免重複觸發
        SetControlsEnabled(false);
        rtbLog.Clear();
        AppendLog("═══════════════════════════════════════");
        AppendLog("  Excel 轉 Markdown 轉換工具  開始執行");
        AppendLog("═══════════════════════════════════════");

        // 建立進度回報物件，確保訊息在 UI 執行緒上更新
        var progress = new Progress<string>(AppendLog);

        // 非同步執行 Excel 轉換，await 確保 UI 執行緒不被阻塞
        var results = await _excelConverterService.ConvertAsync(sourceFilePath, progress);

        // 統計成功與失敗的工作表數量
        int successCount = results.Count(r => r.IsSuccess);
        int failCount = results.Count - successCount;

        AppendLog("───────────────────────────────────────");

        if (failCount == 0 && successCount > 0)
        {
            AppendLog($"✔ 全部 {successCount} 個工作表轉換成功！");
            AppendLog("═══════════════════════════════════════");
            string outputList = string.Join("\n", results.Select(r => r.OutputFilePath));
            MessageBox.Show(
                $"Excel 轉換完成！\n共 {successCount} 個工作表已轉換。\n\n輸出路徑：\n{outputList}",
                "轉換成功",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        else if (successCount == 0)
        {
            string errorMsg = results.FirstOrDefault()?.ErrorMessage ?? "未知錯誤";
            AppendLog($"✘ 轉換失敗：{errorMsg}");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"Excel 轉換失敗：\n{errorMsg}",
                "轉換失敗",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        else
        {
            AppendLog($"⚠ 部分完成：{successCount} 個工作表成功，{failCount} 個工作表失敗。");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"Excel 部分轉換完成。\n成功：{successCount} 個工作表\n失敗：{failCount} 個工作表\n\n請查看執行日誌了解詳情。",
                "部分完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        // 恢復按鈕狀態
        SetControlsEnabled(true);
    }

    /// <summary>
    /// 「瀏覽...」按鈕的點擊事件處理方法（PowerPoint）。
    /// 開啟檔案選取對話方塊，讓使用者選取 PowerPoint (.pptx / .ppt) 檔案。
    /// </summary>
    private void BtnBrowsePpt_Click(object sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "選取 PowerPoint 檔案",
            Filter = "所有 PowerPoint 文件 (*.pptx;*.ppt)|*.pptx;*.ppt|PowerPoint 簡報 (*.pptx)|*.pptx|PowerPoint 97-2003 簡報 (*.ppt)|*.ppt",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            txtPptFilePath.Text = dialog.FileName;
            btnConvertPpt.Enabled = true;
            AppendLog($"已選取 PowerPoint 檔案：{dialog.FileName}");
        }
    }

    /// <summary>
    /// 「轉換 PowerPoint」按鈕的點擊事件處理方法（非同步）。
    /// 驗證輸入後，呼叫 PowerPoint 轉換服務執行投影片轉 Markdown 工作。
    /// </summary>
    private async void BtnConvertPpt_Click(object sender, EventArgs e)
    {
        string sourceFilePath = txtPptFilePath.Text.Trim();

        // 驗證使用者是否已選取 PowerPoint 檔案
        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            MessageBox.Show("請先選取要轉換的 PowerPoint 檔案。", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // 驗證副檔名必須為 .pptx 或 .ppt
        string extension = Path.GetExtension(sourceFilePath);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
         && !extension.Equals(".ppt", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("僅支援 .pptx 及 .ppt 格式的 PowerPoint 檔案。", "格式錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 轉換期間停用按鈕，避免重複觸發
        SetControlsEnabled(false);
        rtbLog.Clear();
        AppendLog("═══════════════════════════════════════");
        AppendLog("  PowerPoint 轉 Markdown 轉換工具  開始執行");
        AppendLog("═══════════════════════════════════════");

        var progress = new Progress<string>(AppendLog);

        var result = await _powerPointConverterService.ConvertAsync(sourceFilePath, progress);

        if (result.IsSuccess)
        {
            AppendLog("───────────────────────────────────────");
            AppendLog($"✔ 轉換成功！");
            AppendLog($"  輸出路徑：{result.OutputFilePath}");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"轉換完成！\n\n輸出路徑：\n{result.OutputFilePath}",
                "轉換成功",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        else
        {
            AppendLog("───────────────────────────────────────");
            AppendLog($"✘ 轉換失敗：{result.ErrorMessage}");
            AppendLog("═══════════════════════════════════════");
            MessageBox.Show(
                $"轉換失敗：\n{result.ErrorMessage}",
                "轉換失敗",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        SetControlsEnabled(true);
    }

    /// <summary>
    /// 將訊息文字附加至執行日誌 RichTextBox，並自動捲動至最新行。
    /// </summary>
    /// <param name="message">要附加的日誌訊息文字。</param>
    private void AppendLog(string message)
    {
        // 確保在 UI 執行緒上執行（來自 Progress<T> 的回呼已在 UI 執行緒，
        // 此處仍加入保護以應對直接呼叫的情境）
        if (rtbLog.InvokeRequired)
        {
            rtbLog.Invoke(() => AppendLog(message));
            return;
        }

        rtbLog.AppendText(message + Environment.NewLine);
        // 自動捲動至最新一行
        rtbLog.ScrollToCaret();
    }

    /// <summary>
    /// 控制主要操作控制項的啟用/停用狀態。
    /// 轉換進行中時停用輸入，轉換完成後恢復。
    /// </summary>
    /// <param name="enabled">
    /// <c>true</c> 表示啟用控制項；<c>false</c> 表示停用控制項。
    /// </param>
    private void SetControlsEnabled(bool enabled)
    {
        btnBrowse.Enabled = enabled;
        btnConvert.Enabled = enabled;
        txtFilePath.Enabled = enabled;
        btnBrowseExcel.Enabled = enabled;
        btnConvertExcel.Enabled = enabled;
        txtExcelFilePath.Enabled = enabled;
        btnBrowsePpt.Enabled = enabled;
        btnConvertPpt.Enabled = enabled;
        txtPptFilePath.Enabled = enabled;
    }
}
