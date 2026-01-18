/**
 * 導出管理組件
 * 支持多種格式導出：Excel、CSV、PDF、HTML
 */

const ExportManager = {
  elements: {},

  init() {
    this._cacheElements();
    this._bindEvents();
  },

  _cacheElements() {
    this.elements = {
      exportButtons: document.getElementById("exportButtons"),
      btnExportExcel: document.getElementById("btnExportExcel"),
      btnExportCSV: document.getElementById("btnExportCSV"),
      btnExportPDF: document.getElementById("btnExportPDF"),
      btnExportHTML: document.getElementById("btnExportHTML"),
    };
  },

  _bindEvents() {
    // 導出按鈕
    this.elements.btnExportExcel?.addEventListener("click", () => {
      this.exportExcel();
    });

    this.elements.btnExportCSV?.addEventListener("click", () => {
      this.exportCSV();
    });

    this.elements.btnExportPDF?.addEventListener("click", () => {
      this.exportPDF();
    });

    this.elements.btnExportHTML?.addEventListener("click", () => {
      this.exportHTML();
    });

    // 監聽錯誤修正事件
    document.addEventListener("errorFixed", (e) => {
      this._addFixedErrorToExport(e.detail.error);
    });
  },

  exportExcel() {
    const sheetData = Store.getState("sheetData");
    const errors = Store.getState("errors");
    const workbook = XLSX.utils.book_new();

    if (!sheetData || sheetData.length === 0) {
      UIController.showToast("error", "沒有數據可導出");
      return;
    }

    // 創建包含錯誤註釋的工作表
    const sheetDataWithErrors = this._addErrorAnnotations(sheetData, errors);
    const worksheet = XLSX.utils.aoa_to_sheet(sheetDataWithErrors);

    // 添加錯誤報告作為單獨工作表
    if (errors.length > 0) {
      const errorReport = this._createErrorReportSheet(errors);
      XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
      XLSX.utils.book_append_sheet(workbook, errorReport, "Error Report");
    } else {
      XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    }

    // 導出文件
    const timestamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `sumcheck-export-${timestamp}.xlsx`);

    UIController.showToast("success", "Excel 導出成功");
  },

  exportCSV() {
    const sheetData = Store.getState("sheetData");
    const errors = Store.getState("errors");

    if (!sheetData || sheetData.length === 0) {
      UIController.showToast("error", "沒有數據可導出");
      return;
    }

    // 創建CSV內容
    const csvContent = this._createCSVContent(sheetData, errors);

    // 導出文件
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `sumcheck-export-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);

    UIController.showToast("success", "CSV 導出成功");
  },

  exportPDF() {
    const errors = Store.getState("errors");

    // 使用瀏覽器打印功能生成PDF
    const printWindow = window.open("", "_blank");
    if (printWindow) {
      printWindow.document.write(`
                <!DOCTYPE html>
                <html>
                <head>
                    <title>驗算報告</title>
                    <style>
                        body {
                            font-family: Arial, sans-serif;
                            padding: 20px;
                            color: #333;
                        }
                        .header {
                            text-align: center;
                            margin-bottom: 30px;
                        }
                        .error-count {
                            background: #f8d7da;
                            color: #721c24;
                            padding: 15px;
                            border-radius: 5px;
                            font-size: 18px;
                            font-weight: bold;
                        }
                        .error-list {
                            margin-top: 20px;
                        }
                        .error-item {
                            padding: 10px;
                            border: 1px solid #ddd;
                            margin-bottom: 10px;
                            border-radius: 3px;
                        }
                        .error-type {
                            color: #d9534f;
                            font-weight: bold;
                        }
                        .error-location {
                            color: #666;
                            margin-bottom: 5px;
                        }
                    </style>
                </head>
                <body>
                    <div class="header">
                        <h1>SumCheck 驗算報告</h1>
                        <p>生成時間：${new Date().toLocaleString("zh-TW")}</p>
                    </div>
                    <div class="error-count">
                        共發現 ${errors.length} 個錯誤
                    </div>
                    ${errors.length > 0 ? this._generatePDFErrorList(errors) : "<p>未發現錯誤！</p>"}
                </body>
                </html>
            `);
      printWindow.document.close();
      printWindow.print();

      UIController.showToast("success", "PDF 導出準備完成");
    }
  },

  exportHTML() {
    const sheetData = Store.getState("sheetData");
    const errors = Store.getState("errors");

    if (!sheetData || sheetData.length === 0) {
      UIController.showToast("error", "沒有數據可導出");
      return;
    }

    // 創建HTML報告
    const htmlContent = this._generateHTMLReport(sheetData, errors);

    // 導出文件
    const blob = new Blob([htmlContent], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `sumcheck-report-${new Date().toISOString().slice(0, 10)}.html`;
    a.click();
    URL.revokeObjectURL(url);

    UIController.showToast("success", "HTML 報告導出成功");
  },

  _addErrorAnnotations(sheetData, errors) {
    if (errors.length === 0) return sheetData;

    // 為每個錯誤添加註釋
    return sheetData.map((row, rowIndex) => {
      return row.map((cell, colIndex) => {
        const actualRow = rowIndex + 1;
        const actualCol = colIndex + 1;
        const error = errors.find(
          (err) => err.row === actualRow && err.col === actualCol,
        );

        if (error) {
          return `${cell} [錯誤：${error.type}]`;
        }
        return cell;
      });
    });
  },

  _createErrorReportSheet(errors) {
    const headers = [
      "錯誤類型",
      "行",
      "列",
      "預期值",
      "實際值",
      "差異",
      "描述",
    ];
    const rows = errors.map((err) => [
      this._getErrorTypeName(err.type),
      err.row,
      err.col,
      err.expected || "-",
      err.actual || "-",
      err.diff || "-",
      err.description || "",
    ]);

    return XLSX.utils.aoa_to_sheet([headers, ...rows]);
  },

  _createCSVContent(sheetData, errors) {
    // CSV 表頭
    const headers = sheetData[0]
      ? sheetData[0].map((_, i) => `Column_${i + 1}`)
      : [];

    // 添加錯誤信息列
    let csv = headers.join(",") + "\n";

    // 添加數據行
    sheetData.forEach((row) => {
      csv +=
        row
          .map((cell) => {
            if (cell === null || cell === undefined) return "";
            // 處理包含逗號和換行的值
            const str = String(cell);
            if (str.includes(",") || str.includes("\n") || str.includes('"')) {
              return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
          })
          .join(",") + "\n";
    });

    // 添加錯誤報告
    if (errors.length > 0) {
      csv += "\n\n=== 錯誤報告 ===\n";
      csv += "錯誤類型,行,列,預期值,實際值,差異\n";
      errors.forEach((err) => {
        csv += `${this._getErrorTypeName(err.type)},${err.row},${err.col},${err.expected || "-"},${err.actual || "-"},${err.diff || "-"}\n`;
      });
    }

    return csv;
  },

  _generatePDFErrorList(errors) {
    return `
            <div class="error-list">
                ${errors
                  .map(
                    (err) => `
                    <div class="error-item">
                        <div class="error-type">${this._getErrorTypeName(err.type)}</div>
                        <div class="error-location">位置：行 ${err.row}, 列 ${err.col}</div>
                        <div>預期值：${err.expected || "無"}</div>
                        <div>實際值：${err.actual || "無"}</div>
                        <div>差異：${err.diff || "無"}</div>
                    </div>
                `,
                  )
                  .join("")}
            </div>
        `;
  },

  _generateHTMLReport(sheetData, errors) {
    const timestamp = new Date().toLocaleString("zh-TW");

    return `
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SumCheck 驗算報告</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            background: #f5f5f5;
            color: #333;
            padding: 20px;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            padding: 30px 20px;
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            border-radius: 10px 10px 0 0;
        }
        .header h1 {
            margin: 0 0 10px 0;
            font-size: 28px;
        }
        .header p {
            margin: 0;
            opacity: 0.9;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px 20px;
            margin-bottom: 30px;
        }
        .summary-card {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        .summary-card.success {
            border-left-color: #28a745;
        }
        .summary-card.error {
            border-left-color: #dc3545;
        }
        .summary-card h3 {
            margin: 0 0 10px 0;
            font-size: 24px;
        }
        .summary-card p {
            margin: 0;
            font-size: 18px;
            font-weight: 600;
        }
        .table-container {
            overflow-x: auto;
            padding: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }
        th {
            background: #f8f9fa;
            padding: 12px;
            font-weight: 600;
            border: 1px solid #dee2e6;
            position: sticky;
            top: 0;
        }
        td {
            padding: 8px 12px;
            border: 1px solid #dee2e6;
        }
        .error-cell {
            background: #fff3cd;
            color: #856404;
        }
        .footer {
            text-align: center;
            padding: 20px;
            color: #6c757d;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 SumCheck 驗算報告</h1>
            <p>生成時間：${timestamp}</p>
        </div>

        <div class="summary">
            <div class="summary-card ${errors.length === 0 ? "success" : "error"}">
                <h3>錯誤數量</h3>
                <p>${errors.length}</p>
            </div>
            <div class="summary-card">
                <h3>總行數</h3>
                <p>${sheetData.length}</p>
            </div>
            <div class="summary-card">
                <h3>總列數</h3>
                <p>${Math.max(...sheetData.map((row) => row.length))}</p>
            </div>
        </div>

        ${
          errors.length > 0
            ? `
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>類型</th>
                        <th>行</th>
                        <th>列</th>
                        <th>預期值</th>
                        <th>實際值</th>
                        <th>差異</th>
                    </tr>
                </thead>
                <tbody>
                    ${errors
                      .map(
                        (err) => `
                        <tr>
                            <td>${this._getErrorTypeName(err.type)}</td>
                            <td>${err.row}</td>
                            <td>${err.col}</td>
                            <td class="error-cell">${err.expected || "-"}</td>
                            <td class="error-cell">${err.actual || "-"}</td>
                            <td>${err.diff || "-"}</td>
                        </tr>
                    `,
                      )
                      .join("")}
                </tbody>
            </table>
        </div>
        `
            : `
        <div style="text-align: center; padding: 40px; color: #28a745; font-size: 18px;">
            ✅ 未發現錯誤！
        </div>
        `
        }

        <div class="footer">
            <p>由 SumCheck 自動生成 | 生成時間：${timestamp}</p>
        </div>
    </div>
</body>
</html>
        `;
  },

  _getErrorTypeName(type) {
    const typeNames = {
      sum_mismatch: "加總不符",
      missing_value: "遺漏值",
      formula_error: "公式錯誤",
      data_inconsistency: "數據不一致",
    };
    return typeNames[type] || type;
  },

  _addFixedErrorToExport(error) {
    // 如果需要，可以添加到導出數據中
    console.log("錯誤已修正:", error);
  },

  exportAll() {
    // 一鍵導出所有格式
    const fileName = `sumcheck-all-${new Date().toISOString().slice(0, 10)}`;

    // 導出Excel
    this.exportExcel();

    // 導出HTML
    setTimeout(() => {
      this.exportHTML();
    }, 500);

    UIController.showToast("success", "正在導出所有格式...");
  },
};
