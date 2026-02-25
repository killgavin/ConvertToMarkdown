# ConvertToMarkdown

> 將 Word (.docx / .doc) 檔案轉換為 Markdown 文件的 WinForms 桌面工具

[![.NET](https://img.shields.io/badge/.NET-8.0--windows-512BD4)](https://dotnet.microsoft.com/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

---

## 功能特色

- 📄 **一鍵轉換**：選取 `.docx` 或 `.doc` 檔案後，點擊「開始轉換」即可產出 Markdown 文件
- 🖼 **圖片自動提取**：將 Word 內嵌圖片儲存為獨立圖檔，Markdown 使用相對路徑引用
- 📊 **表格合併儲存格處理**：自動展開 `colspan`/`rowspan`，確保符合 Markdown 表格語法
- ⚡ **非同步執行**：轉換過程在背景執行緒進行，UI 不凍結
- 📋 **即時執行日誌**：轉換每個步驟的狀態即時顯示於深色日誌區

---

## 系統需求

| 項目 | 需求 |
|---|---|
| 作業系統 | Windows 10 / 11（64-bit） |
| 執行環境 | [.NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) |
| 輸入格式 | Word 2007 以上 (`.docx`) / Word 97-2003 (`.doc`) |
| .doc 支援 | 需安裝 Microsoft Word（透過 COM Interop 轉換） |

---

## 使用方式

1. 啟動程式 (`ConvertToMarkdown.exe`)
2. 點擊「**瀏覽...**」選取要轉換的 `.docx` 或 `.doc` 檔案
3. 點擊「**開始轉換**」
4. 轉換完成後，至來源檔案所在目錄的 `MD/` 資料夾查看結果

### 輸出結構

```
📁 your-folder/
├── document.docx          ← 來源 Word 檔案
└── 📁 MD/
    ├── document.md        ← 轉換產出的 Markdown 文件
    ├── document_圖片_001.png
    └── document_圖片_002.jpg
```

---

## 建置方式

```bash
# 還原套件並建置
dotnet build -c Release

# 執行
dotnet run
```

> 在非 Windows 環境建置時，`.csproj` 已設定 `EnableWindowsTargeting=true`，可正常完成編譯。

---

## 相依套件

| 套件 | 版本 | 用途 |
|---|---|---|
| [Mammoth](https://github.com/mwilliamson/dotnet-mammoth) | 1.3.1 | DOCX → HTML（無需安裝 Microsoft Office）|
| [ReverseMarkdown](https://github.com/mysticmind/reversemarkdown-net) | 4.6.0 | HTML → GitHub Flavored Markdown |
| [HtmlAgilityPack](https://html-agility-pack.net/) | 1.11.72 | HTML DOM 解析，用於表格正規化 |
| Microsoft.Office.Interop.Word | COM | .doc → .docx 預轉換（需安裝 Word） |

---

## 變更日誌

詳見 [CHANGELOG.md](CHANGELOG.md)。
