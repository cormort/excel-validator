# SumCheck ✓

> 智能 Excel 數值驗算工具，自動檢查合計與小計錯誤

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

## ✨ 功能特色

- 🤖 **智能模式推薦** - 自動分析表格結構，推薦最適合的驗算模式
- 🧙 **引導式精靈** - 三步驟輕鬆完成設定
- 📊 **多種驗算模式** - 支援縱向/橫向、關鍵字分組、縮排分層、手動指定
- 🌙 **明暗主題** - 支援 Light/Dark 模式切換
- 📑 **稽核報告** - 一鍵產生完整的 HTML 錯誤報告
- 🚀 **零後端依賴** - 純前端運作，資料不會上傳

## 🎯 使用方式

### 線上使用
直接開啟 [GitHub Pages](https://你的帳號.github.io/excel-validator-main/)

### 本地使用
1. Clone 此專案
2. 用瀏覽器開啟 `index.html`

## 📖 驗算模式說明

| 模式 | 適用場景 | 操作方式 |
|------|---------|---------|
| **縱向 (關鍵字分組)** | 預算表、財務報表 | 選擇名稱欄，系統依關鍵字自動分組 |
| **縱向 (縮排分層)** | 階層結構表格 | 選擇名稱欄，系統依縮排自動辨識 |
| **橫向 (關鍵字分組)** | 一般報表 | 系統依標題關鍵字自動結算 |
| **橫向 (手動指定)** | 自訂公式 | 點選欄位設定 A±B=C |
| **縱向 (手動指定)** | 自訂公式 | 點選列設定 A±B=C |

## 🏗️ 專案結構

```
excel-validator-main/
├── index.html              # 主頁面
├── css/
│   ├── variables.css       # CSS 變數與主題
│   ├── components.css      # UI 元件樣式
│   └── main.css           # 主樣式整合
├── js/
│   ├── app.js             # 應用程式入口
│   ├── excel-parser.js    # Excel 解析模組
│   ├── validator.js       # 驗算邏輯模組
│   ├── smart-detect.js    # 智能偵測模組
│   ├── ui-controller.js   # UI 控制模組
│   └── report.js          # 報告產生模組
├── README.md
└── LICENSE
```

## 🛠️ 技術棧

- **Excel 解析**: [SheetJS](https://sheetjs.com/) (xlsx)
- **UI**: 原生 HTML/CSS/JavaScript
- **設計系統**: CSS Variables + 響應式設計
- **主題**: Light/Dark 模式

## 📝 授權條款

MIT License - 詳見 [LICENSE](LICENSE)

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！
