/**
 * Excel 驗算大師 - 報告產生模組
 * 產生 HTML 稽核報告
 */

const ReportGenerator = {
  /**
   * 產生完整 HTML 報告
   * @param {Object} options - 報告選項
   * @returns {string} HTML 內容
   */
  generateHTML(options) {
    const {
      sheetData,
      headers,
      errors,
      corrections,
      headerRowIndex,
      startColIndex,
      endRowIndex,
      endColIndex,
    } = options;

    // 建立修正後的資料副本
    const correctedData = JSON.parse(JSON.stringify(sheetData));
    corrections.forEach((value, key) => {
      const row = Math.floor(key / 10000);
      const col = key % 10000;
      if (correctedData[row]) correctedData[row][col] = value;
    });

    const errorList = Array.from(errors.values());
    const totalDiff = errorList.reduce((sum, e) => sum + e.diff, 0);

    return `
<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel 驗算稽核報告</title>
  <style>
    :root {
      --primary: #6366f1;
      --success: #10b981;
      --error: #ef4444;
      --bg: #f8fafc;
      --text: #1e293b;
      --border: #e2e8f0;
    }
    
    * { box-sizing: border-box; margin: 0; padding: 0; }
    
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
      padding: 40px;
    }
    
    .container { max-width: 1200px; margin: 0 auto; }
    
    h1 {
      font-size: 28px;
      margin-bottom: 24px;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    
    h2 {
      font-size: 20px;
      margin: 32px 0 16px;
      padding-bottom: 8px;
      border-bottom: 2px solid var(--border);
      display: flex;
      align-items: center;
      gap: 8px;
    }
    
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 16px;
      margin-bottom: 32px;
    }
    
    .stat-card {
      background: white;
      padding: 20px;
      border-radius: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .stat-label { font-size: 13px; color: #64748b; margin-bottom: 4px; }
    .stat-value { font-size: 28px; font-weight: 700; }
    .stat-value.error { color: var(--error); }
    .stat-value.success { color: var(--success); }
    
    .section {
      background: white;
      border-radius: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      overflow: hidden;
      margin-bottom: 24px;
    }
    
    .section-body { padding: 20px; overflow-x: auto; }
    
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td {
      border: 1px solid var(--border);
      padding: 8px 12px;
      text-align: left;
      white-space: nowrap;
    }
    th { background: #f1f5f9; font-weight: 600; }
    
    .err { background: #fee2e2; color: #991b1b; font-weight: 600; border: 2px solid var(--error); }
    .fixed { background: #d1fae5; color: #065f46; font-weight: 600; }
    .diff { color: var(--error); font-weight: 700; }
    
    .error-list { list-style: none; }
    .error-item {
      display: flex;
      align-items: center;
      gap: 16px;
      padding: 12px 16px;
      border-bottom: 1px solid var(--border);
    }
    .error-item:last-child { border-bottom: none; }
    
    .error-loc {
      background: var(--primary);
      color: white;
      padding: 4px 10px;
      border-radius: 6px;
      font-weight: 600;
      font-size: 12px;
    }
    
    .error-detail { flex: 1; }
    .error-expected { font-weight: 600; color: var(--success); }
    .error-actual { font-weight: 600; color: var(--error); }
    
    .footer {
      text-align: center;
      margin-top: 40px;
      color: #94a3b8;
      font-size: 13px;
    }
    
    @media print {
      body { padding: 20px; }
      .section { break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>📑 Excel 驗算稽核報告</h1>
    
    <div class="summary">
      <div class="stat-card">
        <div class="stat-label">發現錯誤數</div>
        <div class="stat-value error">${errorList.length}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">差異總額</div>
        <div class="stat-value ${totalDiff >= 0 ? 'error' : 'success'}">
          ${totalDiff >= 0 ? '+' : ''}${totalDiff.toLocaleString()}
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-label">產生時間</div>
        <div class="stat-value" style="font-size: 16px;">${new Date().toLocaleString()}</div>
      </div>
    </div>
    
    <h2>📋 錯誤摘要</h2>
    <div class="section">
      <ul class="error-list">
        ${errorList.map(e => `
          <li class="error-item">
            <span class="error-loc">${this._getCellName(e.row, e.col)}</span>
            <div class="error-detail">
              原始值：<span class="error-actual">${e.actual.toLocaleString()}</span>
              → 應為：<span class="error-expected">${e.expected.toLocaleString()}</span>
              <span class="diff">（差異 ${e.diff >= 0 ? '+' : ''}${e.diff.toLocaleString()}）</span>
            </div>
          </li>
        `).join('')}
      </ul>
    </div>
    
    <h2>🔴 原始資料 (錯誤標示)</h2>
    <div class="section">
      <div class="section-body">
        ${this._generateTableHTML(sheetData, headers, errors, null, 'original', headerRowIndex, startColIndex, endRowIndex, endColIndex)}
      </div>
    </div>
    
    <h2>🟢 建議修正</h2>
    <div class="section">
      <div class="section-body">
        ${this._generateTableHTML(correctedData, headers, null, corrections, 'corrected', headerRowIndex, startColIndex, endRowIndex, endColIndex)}
      </div>
    </div>
    
    <h2>🔵 差異分析</h2>
    <div class="section">
      <div class="section-body">
        ${this._generateDiffTableHTML(sheetData, correctedData, headers, corrections, headerRowIndex, startColIndex, endRowIndex, endColIndex)}
      </div>
    </div>
    
    <div class="footer">
      由 SumCheck 自動產生 | ${new Date().toLocaleDateString()}
    </div>
  </div>
</body>
</html>`;
  },

  /**
   * 產生表格 HTML
   */
  _generateTableHTML(data, headers, errors, corrections, type, hRowIdx, sCol, eRow, eCol) {
    let html = '<table><thead><tr><th>列</th>';

    for (let c = sCol; c <= eCol && c < headers.length; c++) {
      html += `<th>${headers[c] || ''}</th>`;
    }
    html += '</tr></thead><tbody>';

    for (let r = hRowIdx + 1; r <= eRow && r < data.length; r++) {
      const row = data[r];
      if (!row) continue;

      html += `<tr><td>${r + 1}</td>`;

      for (let c = sCol; c <= eCol && c < (row.length || headers.length); c++) {
        const key = r * 10000 + c;
        const val = row[c];
        const display = typeof val === 'number' ? val.toLocaleString() : (val || '');

        let cls = '';
        if (type === 'original' && errors?.has(key)) cls = 'err';
        if (type === 'corrected' && corrections?.has(key)) cls = 'fixed';

        html += `<td class="${cls}">${display}</td>`;
      }
      html += '</tr>';
    }

    html += '</tbody></table>';
    return html;
  },

  /**
   * 產生差異表格 HTML
   */
  _generateDiffTableHTML(orig, fixed, headers, corrections, hRowIdx, sCol, eRow, eCol) {
    let html = '<table><thead><tr><th>列</th>';

    for (let c = sCol; c <= eCol && c < headers.length; c++) {
      html += `<th>${headers[c] || ''}</th>`;
    }
    html += '</tr></thead><tbody>';

    for (let r = hRowIdx + 1; r <= eRow && r < orig.length; r++) {
      html += `<tr><td>${r + 1}</td>`;

      for (let c = sCol; c <= eCol; c++) {
        const key = r * 10000 + c;
        let content = '-';
        let cls = '';

        if (corrections.has(key)) {
          const v1 = parseFloat(orig[r]?.[c]) || 0;
          const v2 = parseFloat(fixed[r]?.[c]) || 0;
          const diff = v1 - v2;
          content = diff.toLocaleString(undefined, { minimumFractionDigits: 2 });
          cls = 'diff';
        }

        html += `<td class="${cls}">${content}</td>`;
      }
      html += '</tr>';
    }

    html += '</tbody></table>';
    return html;
  },

  /**
   * 取得儲存格名稱 (如 A1, B2)
   */
  _getCellName(row, col) {
    const colName = this._getColumnName(col);
    return `${colName}${row + 1}`;
  },

  /**
   * 數字轉欄名 (0->A, 1->B, 26->AA)
   */
  _getColumnName(num) {
    let name = '';
    while (num >= 0) {
      name = String.fromCharCode(65 + (num % 26)) + name;
      num = Math.floor(num / 26) - 1;
    }
    return name;
  },

  /**
   * 下載報告
   */
  download(options) {
    const html = this.generateHTML(options);
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `稽核報告_${new Date().toISOString().slice(0, 10)}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  },
};

// 匯出模組
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ReportGenerator;
}
