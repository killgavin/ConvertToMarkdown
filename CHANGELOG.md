# 變更日誌 (Changelog)

本專案遵循 [Conventional Commits](https://www.conventionalcommits.org/zh-hant/) 規範撰寫提交訊息。

---

## [1.1.0] — 2026-02-25

### 提交資訊

```
feat: 支援 .doc (Word 97-2003) 格式轉換（透過 Word COM Interop 預轉為 .docx）
```

**分支**：`copilot/add-winform-file-conversion-tool`

---

### 新增功能 (Features)

#### 📄 .doc 格式支援

- 新增 **步驟 0（條件性）**：偵測到 `.doc` 檔案時，自動透過 Microsoft Word COM Interop 轉換為臨時 `.docx` 再進入現有轉換管線
- 新增 `ConverterService.ConvertDocToDocx()` 私有方法，負責 COM Interop 轉換遏輯
- 臨時檔案於轉換完成後自動清除（`finally` 區塊）
- 當未安裝 Microsoft Word 時，捕捉 `COMException` 並給出友善提示

#### 🖥 UI 調整

- `OpenFileDialog` 篩選器新增 `.doc` 格式支援
- 副檔名驗證放寬為接受 `.docx` 及 `.doc`

#### 📦 專案設定

- `ConvertToMarkdown.csproj` 新增 `Microsoft.Office.Interop.Word` NuGet 套件參考

#### 📝 系統需求變更

- 轉換 `.doc` 檔案時，使用者電腦需安裝 Microsoft Word
- 轉換 `.docx` 檔案時無額外需求（行為與 v1.0.0 完全相同）

---

## [1.0.0] — 2026-02-25

### 提交資訊

```
feat: 實作 Word 轉 Markdown WinForms 工具 (階段 0-3)
```

**提交雜湊**：`c01487b`  
**分支**：`copilot/add-winform-file-conversion-tool`  
**作者**：copilot-swe-agent  

---

### 新增功能 (Features)

本次提交從零建置整個 .NET 8.0 WinForms 應用程式，為首次功能性提交，包含以下所有模組：

#### 🏗 專案基礎結構

| 新增檔案 | 說明 |
|---|---|
| `ConvertToMarkdown.csproj` | .NET 8.0-windows WinForms 專案設定檔，包含三個 NuGet 套件相依性 |
| `.gitignore` | 排除 `bin/`、`obj/`、`*.user` 等建置產出物，避免污染版本控制 |

#### 📐 介面合約層 (`IConverterService.cs`)

定義轉換服務行為合約，解耦 UI 層與轉換邏輯層：

- **`IConverterService` 介面**：宣告 `ConvertAsync(string, IProgress<string>)` 非同步方法簽章
- **`ConversionResult` 類別**：封裝轉換結果，包含：
  - `IsSuccess`：是否成功完成
  - `OutputFilePath`：輸出 `.md` 檔案的完整路徑
  - `ErrorMessage`：失敗時的友善錯誤訊息

#### ⚙️ 核心轉換邏輯 (`ConverterService.cs`)

實作 `IConverterService`，完整的六步驟轉換流程：

**步驟 1 — 驗證來源檔案**
- 確認指定路徑的 `.docx` 檔案確實存在，提前回傳友善錯誤

**步驟 2 — 建立輸出資料夾**
- 於來源檔案所在目錄下自動建立 `MD/` 子資料夾

**步驟 3 — Mammoth 解析 Word**
- 使用 `Mammoth.DocumentConverter.ConvertToHtml(path)` 將 `.docx` 轉為 HTML 字串
- 輸出 Mammoth 產生的所有轉換警告至執行日誌

**步驟 4 — 圖片提取與儲存** (`ExtractAndSaveImages`)
- 以 Compiled Regex 掃描 `src="data:image/TYPE;base64,DATA"` 格式的嵌入圖片
- 解碼 base64 → 以原始二進位寫入 `MD/` 資料夾（命名規則：`{主檔名}_圖片_{序號三位數}.{副檔名}`）
- 將 `src` 屬性值替換為相對路徑，確保 Markdown 可正確引用
- 支援副檔名：`jpg`、`png`、`gif`、`webp`、`svg`

**步驟 5 — 表格合併儲存格平鋪化** (`NormalizeTables` / `FlattenTable`)
- Markdown 表格語法不支援 `colspan`/`rowspan`，需先展開
- 以 HtmlAgilityPack 解析 HTML DOM，XPath 選取所有 `<table>` 節點
- 核心演算法：以虛擬格子矩陣 `Dictionary<(row, col), string>` 記錄每個邏輯座標的內容
  - 合併範圍首格保留原始內容，其餘格填入空字串
  - 展開完成後重建 `<table>`，第 0 行使用 `<th>`，其餘行使用 `<td>`
  - 移除所有 `colspan`/`rowspan` 屬性

**步驟 6 — ReverseMarkdown 轉換**
- 使用 `ReverseMarkdown.Converter` 以 GFM (GitHub Flavored Markdown) 設定轉換 HTML
- 開啟 `GithubFlavored`、`RemoveComments`、`SmartHrefHandling` 選項
- 輸出 UTF-8 with BOM 的 `.md` 檔案（確保各種編輯器相容性）

#### 🖥 使用者介面層

**`MainForm.cs` — 事件處理邏輯**

| 方法 | 說明 |
|---|---|
| `BtnBrowse_Click` | 開啟 `OpenFileDialog`，篩選 `*.docx`，將路徑填入文字方塊並啟用轉換按鈕 |
| `BtnConvert_Click` | 非同步事件處理（`async void`）；轉換中停用全部控制項防止重複觸發；以 `Progress<string>` 串接日誌輸出 |
| `AppendLog` | 安全地將訊息附加至 `RichTextBox`，包含 `InvokeRequired` 跨執行緒保護，並自動捲動至最新行 |
| `SetControlsEnabled` | 統一管理「瀏覽」、「開始轉換」按鈕與路徑輸入框的啟用狀態 |

**`MainForm.Designer.cs` — UI 佈局（正體中文標籤）**

```
┌──────────────────────────────────────┐
│ 來源 Word 檔案：                      │
│ [─── 路徑文字方塊 ───────] [瀏覽...]  │
│ [開始轉換]                            │
├──────────────────────────────────────┤
│ 執行日誌：                            │
│ ┌──────────────────────────────────┐ │
│ │  (深色背景 Consolas 字型日誌區)  │ │
│ └──────────────────────────────────┘ │
└──────────────────────────────────────┘
```

- 視窗固定大小（560 × 420），不可最大化
- 全程使用「微軟正黑體 UI」字型，確保正體中文顯示正確
- 日誌區使用深色背景（RGB 30,30,30）搭配淺色文字，提升可讀性

**`Program.cs` — 程式進入點**
- `[STAThread]` 必要屬性（WinForms COM 呼叫需求）
- `ApplicationConfiguration.Initialize()` 套用高 DPI 與視覺樣式設定
- `Application.Run(new MainForm())` 啟動主視窗

---

### 依賴套件 (Dependencies)

| 套件名稱 | 版本 | 用途 |
|---|---|---|
| `Mammoth` | 1.3.1（由 1.0.0 解析） | DOCX → HTML 轉換（無需安裝 Microsoft Office） |
| `ReverseMarkdown` | 4.6.0 | HTML → GitHub Flavored Markdown 轉換 |
| `HtmlAgilityPack` | 1.11.72 | HTML DOM 解析與表格節點操作 |

> **注意**：Mammoth 套件以 .NET Framework 4.x 為目標，在 net8.0-windows 下透過相容性模式運作，功能正常。

---

### 建置方式 (Build)

```bash
# 還原套件並建置（Release 組態）
dotnet build -c Release

# 在非 Windows 環境（如 Linux CI）需設定 EnableWindowsTargeting=true（已於 .csproj 中設定）
```

---

### 輸出結構範例

```
📁 my-docs/
├── 報告.docx                   ← 來源檔案
└── 📁 MD/
    ├── 報告.md                 ← 轉換產出的 Markdown 檔案
    ├── 報告_圖片_001.png        ← 從文件提取的圖片
    ├── 報告_圖片_002.png
    └── 報告_圖片_003.jpg
```

---

### 技術決策紀錄 (ADR)

| 決策 | 理由 |
|---|---|
| 禁用 Office Interop | 需要安裝 Microsoft Office，不符合輕量部署需求 |
| 使用 Mammoth | 純 .NET 實作，可讀取 `.docx` 內部 XML，無外部依賴 |
| 使用 HtmlAgilityPack 正規化表格 | Mammoth 輸出的 HTML 保留原始 `colspan`/`rowspan`，ReverseMarkdown 無法正確處理，需前置處理 |
| `Task.Run` 包裝同步 Mammoth API | Mammoth API 為同步設計，以 `Task.Run` 移至執行緒集區避免 UI 凍結 |
| UTF-8 with BOM 輸出 | 確保 Windows 環境中各種文字編輯器（記事本、VS Code）均可正確識別編碼 |
