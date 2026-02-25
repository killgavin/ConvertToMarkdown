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

        lblFilePath  = new System.Windows.Forms.Label();
        txtFilePath  = new System.Windows.Forms.TextBox();
        btnBrowse    = new System.Windows.Forms.Button();
        btnConvert   = new System.Windows.Forms.Button();
        lblLog       = new System.Windows.Forms.Label();
        rtbLog       = new System.Windows.Forms.RichTextBox();
        panelTop     = new System.Windows.Forms.Panel();
        panelBottom  = new System.Windows.Forms.Panel();

        panelTop.SuspendLayout();
        panelBottom.SuspendLayout();
        SuspendLayout();

        // ── panelTop（來源檔案選取區）─────────────────────────────────────

        panelTop.Dock     = System.Windows.Forms.DockStyle.Top;
        panelTop.Height   = 100;
        panelTop.Padding  = new Padding(12, 12, 12, 0);
        panelTop.Controls.Add(lblFilePath);
        panelTop.Controls.Add(txtFilePath);
        panelTop.Controls.Add(btnBrowse);
        panelTop.Controls.Add(btnConvert);

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

        // btnConvert：觸發轉換流程的主要按鈕
        btnConvert.Font     = new System.Drawing.Font("Microsoft JhengHei UI", 10F, System.Drawing.FontStyle.Bold);
        btnConvert.Location = new System.Drawing.Point(12, 70);
        btnConvert.Size     = new System.Drawing.Size(120, 32);
        btnConvert.Text     = "開始轉換";
        btnConvert.TabIndex = 2;
        btnConvert.Enabled  = false;
        btnConvert.Click   += BtnConvert_Click;

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
        ClientSize          = new System.Drawing.Size(560, 420);
        Controls.Add(panelBottom);
        Controls.Add(panelTop);
        Font                = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
        FormBorderStyle     = System.Windows.Forms.FormBorderStyle.FixedSingle;
        MaximizeBox         = false;
        MinimumSize         = new System.Drawing.Size(576, 458);
        Name                = "MainForm";
        Text                = "Word 轉 Markdown 工具";
        StartPosition       = System.Windows.Forms.FormStartPosition.CenterScreen;

        panelTop.ResumeLayout(false);
        panelTop.PerformLayout();
        panelBottom.ResumeLayout(false);
        panelBottom.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
}
