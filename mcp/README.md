# SumCheck MCP Server

把 SumCheck 的驗算引擎（`js/validator.js`、`js/smart-detect.js`）包成 MCP server，
讓 AI agent 只傳檔案路徑或 TSV 字串就能驗算合計/小計，表格本身不進 context。

## 安裝

```bash
cd mcp && npm install
```

## 註冊到 Claude Code

```bash
claude mcp add sumcheck -- node /Users/hsiehminchieh/Dev/Work/excel-validator/mcp/server.js
```

## Tools

| Tool | 用途 | 必要參數 |
|---|---|---|
| `list_sheets` | 列出工作表 | `file` |
| `detect_mode` | 智能偵測驗算模式（模式、信心度、理由） | `file` 或 `tsv` |
| `validate` | 執行驗算，回傳錯誤清單與差額 | `file` 或 `tsv`（`mode` 省略則自動偵測） |

`validate` 回傳範例：

```json
{
  "mode": "vertical_group",
  "autoDetected": true,
  "confidence": 50,
  "errorCount": 1,
  "totalDiff": -50,
  "truncated": false,
  "errors": [{ "cell": "B2", "expected": 300, "actual": 250, "diff": -50 }]
}
```

手動模式（`horizontal` / `vertical_row`）需傳 `indices`（1-based，最後一個為結果欄/列），
減項用 `signs` 對應標 `-1`。關鍵字模式預設觸發詞 `主管, 小計, 核定`、排除詞 `合計, 總計`，
可用 `trigger_keywords` / `exclude_keywords` 覆寫。
