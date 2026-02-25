namespace ConvertToMarkdown;

partial class MainForm
{
    /// <summary>
    /// 設計工具所需的私有變數。
    /// </summary>
    private System.ComponentModel.IContainer components = null!;

    // 控制項宣告
    private System.Windows.Forms.Label lblFilePath;
    private System.Windows.Forms.TextBox txtFilePath;
    private System.Windows.Forms.Button btnBrowse;
    private System.Windows.Forms.Button btnConvert;
    private System.Windows.Forms.Label lblExcelFilePath;
    private System.Windows.Forms.TextBox txtExcelFilePath;
    private System.Windows.Forms.Button btnBrowseExcel;
    private System.Windows.Forms.Button btnConvertExcel;
    private System.Windows.Forms.Label lblLog;
    private System.Windows.Forms.RichTextBox rtbLog;
    private System.Windows.Forms.Panel panelTop;
    private System.Windows.Forms.Panel panelBottom;

    /// <summary>
    /// 清除所有使用中的資源。
    /// </summary>
    /// <param name="disposing">若受控資源應被清除，則為 true；否則為 false。</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form 設計工具產生的程式碼

    /// <summary>
    /// 設計工具支援所需的方法，請勿使用程式碼編輯器修改此方法的內容。
    /// </summary>
    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        // ── 控制項初始化 ──────────────────────────────────────────────────

        lblFilePath      = new System.Windows.Forms.Label();
        txtFilePath      = new System.Windows.Forms.TextBox();
        btnBrowse        = new System.Windows.Forms.Button();
        btnConvert       = new System.Windows.Forms.Button();
        lblExcelFilePath = new System.Windows.Forms.Label();
        txtExcelFilePath = new System.Windows.Forms.TextBox();
        btnBrowseExcel   = new System.Windows.Forms.Button();
        btnConvertExcel  = new System.Windows.Forms.Button();
        lblLog           = new System.Windows.Forms.Label();
        rtbLog           = new System.Windows.Forms.RichTextBox();
        panelTop         = new System.Windows.Forms.Panel();
        panelBottom      = new System.Windows.Forms.Panel();

        panelTop.SuspendLayout();
        panelBottom.SuspendLayout();
        SuspendLayout();

        // ── panelTop（來源檔案選取區）─────────────────────────────────────

        panelTop.Dock     = System.Windows.Forms.DockStyle.Top;
        panelTop.Height   = 200;
        panelTop.Padding  = new Padding(12, 12, 12, 0);
        panelTop.Controls.Add(lblFilePath);
        panelTop.Controls.Add(txtFilePath);
        panelTop.Controls.Add(btnBrowse);
        panelTop.Controls.Add(btnConvert);
        panelTop.Controls.Add(lblExcelFilePath);
        panelTop.Controls.Add(txtExcelFilePath);
        panelTop.Controls.Add(btnBrowseExcel);
        panelTop.Controls.Add(btnConvertExcel);

        // lblFilePath：「來源 Word 檔案：」標籤
        lblFilePath.AutoSize = true;
        lblFilePath.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F);
        lblFilePath.Location = new System.Drawing.Point(12, 14);
        lblFilePath.Text     = "來源 Word 檔案：";

        // txtFilePath：顯示選取的檔案路徑（唯讀，由瀏覽對話方塊填入）
        txtFilePath.Font      = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        txtFilePath.Location  = new System.Drawing.Point(12, 38);
        txtFilePath.Size      = new System.Drawing.Size(430, 23);
        txtFilePath.ReadOnly  = true;
        txtFilePath.BackColor = System.Drawing.SystemColors.Window;
        txtFilePath.TabIndex  = 0;

        // btnBrowse：瀏覽並選取 Word 檔案的按鈕
        btnBrowse.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        btnBrowse.Location = new System.Drawing.Point(450, 37);
        btnBrowse.Size     = new System.Drawing.Size(80, 26);
        btnBrowse.Text     = "瀏覽...";
        btnBrowse.TabIndex = 1;
        btnBrowse.Click   += BtnBrowse_Click;

        // btnConvert：觸發 Word 轉換流程的主要按鈕
        btnConvert.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold);
        btnConvert.Location = new System.Drawing.Point(12, 70);
        btnConvert.Size     = new System.Drawing.Size(120, 32);
        btnConvert.Text     = "開始轉換";
        btnConvert.TabIndex = 2;
        btnConvert.Enabled  = false;
        btnConvert.Click   += BtnConvert_Click;

        // lblExcelFilePath：「來源 Excel 檔案：」標籤
        lblExcelFilePath.AutoSize = true;
        lblExcelFilePath.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F);
        lblExcelFilePath.Location = new System.Drawing.Point(12, 114);
        lblExcelFilePath.Text     = "來源 Excel 檔案：";

        // txtExcelFilePath：顯示選取的 Excel 檔案路徑（唯讀，由瀏覽對話方塊填入）
        txtExcelFilePath.Font      = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        txtExcelFilePath.Location  = new System.Drawing.Point(12, 138);
        txtExcelFilePath.Size      = new System.Drawing.Size(430, 23);
        txtExcelFilePath.ReadOnly  = true;
        txtExcelFilePath.BackColor = System.Drawing.SystemColors.Window;
        txtExcelFilePath.TabIndex  = 3;

        // btnBrowseExcel：瀏覽並選取 Excel 檔案的按鈕
        btnBrowseExcel.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        btnBrowseExcel.Location = new System.Drawing.Point(450, 137);
        btnBrowseExcel.Size     = new System.Drawing.Size(80, 26);
        btnBrowseExcel.Text     = "瀏覽...";
        btnBrowseExcel.TabIndex = 4;
        btnBrowseExcel.Click   += BtnBrowseExcel_Click;

        // btnConvertExcel：觸發 Excel 轉換流程的按鈕
        btnConvertExcel.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold);
        btnConvertExcel.Location = new System.Drawing.Point(12, 170);
        btnConvertExcel.Size     = new System.Drawing.Size(140, 32);
        btnConvertExcel.Text     = "轉換 Excel";
        btnConvertExcel.TabIndex = 5;
        btnConvertExcel.Enabled  = false;
        btnConvertExcel.Click   += BtnConvertExcel_Click;

        // ── panelBottom（執行日誌區）──────────────────────────────────────

        panelBottom.Dock    = System.Windows.Forms.DockStyle.Fill;
        panelBottom.Padding = new Padding(12, 4, 12, 12);
        panelBottom.Controls.Add(lblLog);
        panelBottom.Controls.Add(rtbLog);

        // lblLog：執行日誌標籤
        lblLog.AutoSize = true;
        lblLog.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F);
        lblLog.Location = new System.Drawing.Point(12, 4);
        lblLog.Text     = "執行日誌：";

        // rtbLog：顯示轉換過程中各步驟執行狀態的多行文字方塊
        rtbLog.Dock        = System.Windows.Forms.DockStyle.Bottom;
        rtbLog.Height      = 270;
        rtbLog.Font        = new System.Drawing.Font("Consolas", 9F);
        rtbLog.BackColor   = System.Drawing.Color.FromArgb(30, 30, 30);
        rtbLog.ForeColor   = System.Drawing.Color.FromArgb(220, 220, 220);
        rtbLog.ReadOnly    = true;
        rtbLog.ScrollBars  = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
        rtbLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        rtbLog.TabIndex    = 3;

        // ── MainForm 主視窗屬性 ────────────────────────────────────────────

        AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
        ClientSize          = new System.Drawing.Size(560, 520);
        Controls.Add(panelBottom);
        Controls.Add(panelTop);
        Font                = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        FormBorderStyle     = System.Windows.Forms.FormBorderStyle.FixedSingle;
        MaximizeBox         = false;
        MinimumSize         = new System.Drawing.Size(576, 558);
        Name                = "MainForm";
        Text                = "Word / Excel 轉 Markdown 工具";
        StartPosition       = System.Windows.Forms.FormStartPosition.CenterScreen;

        panelTop.ResumeLayout(false);
        panelTop.PerformLayout();
        panelBottom.ResumeLayout(false);
        panelBottom.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
}
