/**
 * Excel 驗算大師 - 智能偵測模組
 * 自動分析表格結構並推薦最適合的驗算模式
 */

const SmartDetect = {
  // 關鍵字定義
  KEYWORDS: {
    trigger: ['主管', '小計', '核定', '結轉', '合計', 'Total', 'Subtotal', 'Sum'],
    exclude: ['總計', '總合計', 'Grand Total'],
    indent: ['　', '  ', '    '], // 全形空格、半形空格
  },

  /**
   * 分析表格結構並推薦驗算模式
   * @param {Array} sheetData - 二維陣列的表格資料
   * @param {number} headerRow - 標題列索引 (0-based)
   * @returns {Object} 推薦結果
   */
  analyze(sheetData, headerRow = 0) {
    if (!sheetData || sheetData.length < 3) {
      return { mode: null, confidence: 0, reasons: ['資料不足'] };
    }

    const scores = {
      vertical_group: 0,
      vertical_indent: 0,
      vertical_row: 0,
      horizontal: 0,
      horizontal_group: 0,
    };

    const analysis = {
      keywordRows: [],
      indentedRows: [],
      numericCols: [],
      textCols: [],
      headers: sheetData[headerRow] || [],
    };

    // 1. 分析欄位類型
    this._analyzeColumns(sheetData, headerRow, analysis);

    // 2. 分析關鍵字分布
    this._analyzeKeywords(sheetData, headerRow, analysis);

    // 3. 分析縮排結構
    this._analyzeIndentation(sheetData, headerRow, analysis);

    // 4. 計算各模式分數
    this._calculateScores(analysis, scores);

    // 5. 找出最高分模式
    const sortedModes = Object.entries(scores)
      .sort((a, b) => b[1] - a[1]);
    
    const topMode = sortedModes[0];
    const confidence = Math.min(100, Math.round(topMode[1]));

    return {
      mode: topMode[0],
      confidence,
      reasons: this._getReasons(analysis, topMode[0]),
      analysis,
      allScores: Object.fromEntries(sortedModes),
    };
  },

  /**
   * 分析欄位類型（數值 vs 文字）
   */
  _analyzeColumns(data, headerRow, result) {
    const headers = data[headerRow] || [];
    const sampleRows = data.slice(headerRow + 1, headerRow + 20);

    headers.forEach((header, colIdx) => {
      let numCount = 0;
      let textCount = 0;

      sampleRows.forEach(row => {
        const val = row?.[colIdx];
        if (val === null || val === undefined || val === '') return;

        if (typeof val === 'number' || this._isNumericString(val)) {
          numCount++;
        } else if (typeof val === 'string' && val.trim()) {
          textCount++;
        }
      });

      if (numCount > textCount * 2) {
        result.numericCols.push(colIdx);
      } else if (textCount > 0) {
        result.textCols.push(colIdx);
      }
    });
  },

  /**
   * 分析關鍵字分布
   */
  _analyzeKeywords(data, headerRow, result) {
    const allTriggers = this.KEYWORDS.trigger;
    
    // 掃描資料列
    for (let r = headerRow + 1; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;

      for (let c = 0; c < row.length; c++) {
        const cellText = String(row[c] || '');
        if (allTriggers.some(kw => cellText.includes(kw))) {
          result.keywordRows.push({ row: r, col: c, text: cellText });
        }
      }
    }

    // 掃描標題列
    result.headers.forEach((header, idx) => {
      const headerText = String(header || '');
      if (allTriggers.some(kw => headerText.includes(kw))) {
        result.keywordCols = result.keywordCols || [];
        result.keywordCols.push(idx);
      }
    });
  },

  /**
   * 分析縮排結構
   */
  _analyzeIndentation(data, headerRow, result) {
    // 找出第一個文字欄位
    const textCol = result.textCols[0];
    if (textCol === undefined) return;

    let lastIndent = 0;
    let hasHierarchy = false;

    for (let r = headerRow + 1; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;

      const cellText = String(row[textCol] || '');
      const indent = this._countLeadingSpaces(cellText);

      if (indent > 0) {
        result.indentedRows.push({ row: r, indent, text: cellText });
      }

      if (indent !== lastIndent && lastIndent !== 0) {
        hasHierarchy = true;
      }
      lastIndent = indent;
    }

    result.hasHierarchy = hasHierarchy;
  },

  /**
   * 計算各模式分數
   */
  _calculateScores(analysis, scores) {
    const { keywordRows, indentedRows, numericCols, textCols, keywordCols } = analysis;

    // 縱向關鍵字分組：有關鍵字列 + 文字欄位
    if (keywordRows.length > 0 && textCols.length > 0) {
      scores.vertical_group += 30 + Math.min(keywordRows.length * 5, 40);
    }

    // 縱向縮排：有縮排結構
    if (indentedRows.length > 3 && analysis.hasHierarchy) {
      scores.vertical_indent += 40 + Math.min(indentedRows.length * 3, 40);
    }

    // 橫向關鍵字分組：標題含關鍵字
    if (keywordCols && keywordCols.length > 0) {
      scores.horizontal_group += 30 + Math.min(keywordCols.length * 10, 30);
    }

    // 數值欄位多 → 可能需要手動模式
    if (numericCols.length > 5 && keywordRows.length < 3) {
      scores.horizontal += 20;
      scores.vertical_row += 20;
    }

    // 預設給縱向關鍵字分組一點基礎分（最常用）
    scores.vertical_group += 10;
  },

  /**
   * 產生推薦原因說明
   */
  _getReasons(analysis, mode) {
    const reasons = [];

    switch (mode) {
      case 'vertical_group':
        if (analysis.keywordRows.length > 0) {
          reasons.push(`發現 ${analysis.keywordRows.length} 個分組關鍵字列`);
        }
        if (analysis.textCols.length > 0) {
          reasons.push(`第 ${analysis.textCols[0] + 1} 欄可作為名稱欄`);
        }
        break;

      case 'vertical_indent':
        if (analysis.indentedRows.length > 0) {
          reasons.push(`發現 ${analysis.indentedRows.length} 個縮排層級`);
        }
        reasons.push('表格具有階層結構');
        break;

      case 'horizontal_group':
        if (analysis.keywordCols?.length > 0) {
          reasons.push(`標題列含有 ${analysis.keywordCols.length} 個結算關鍵字`);
        }
        break;

      default:
        reasons.push('建議手動選取欄位設定公式');
    }

    return reasons;
  },

  /**
   * 輔助函數：計算前導空白數
   */
  _countLeadingSpaces(str) {
    if (!str || typeof str !== 'string') return 0;
    const match = str.match(/^[\s\u3000]*/);
    return match ? match[0].length : 0;
  },

  /**
   * 輔助函數：判斷是否為數值字串
   */
  _isNumericString(val) {
    if (typeof val !== 'string') return false;
    const cleaned = val.replace(/[,，$%\s]/g, '');
    return !isNaN(parseFloat(cleaned)) && isFinite(cleaned);
  },

  /**
   * 取得模式的顯示資訊
   */
  getModeInfo(mode) {
    const modes = {
      vertical_group: {
        icon: '↕️',
        name: '縱向 (關鍵字分組)',
        desc: '適用於預算表、財務報表',
      },
      vertical_row: {
        icon: '↕️',
        name: '縱向 (指定列)',
        desc: '手動選取列來設定 A±B=C 公式',
      },
      vertical_indent: {
        icon: '📊',
        name: '縱向 (縮排分層)',
        desc: '適用於有階層結構的表格',
      },
      horizontal: {
        icon: '↔️',
        name: '橫向 (指定欄)',
        desc: '手動選取欄來設定 A±B=C 公式',
      },
      horizontal_group: {
        icon: '↔️',
        name: '橫向 (關鍵字分組)',
        desc: '適用於一般報表、橫向加總',
      },
    };
    return modes[mode] || { icon: '❓', name: '未知模式', desc: '' };
  },
};

// 匯出模組（支援 ES Module 和傳統 script 標籤）
if (typeof module !== 'undefined' && module.exports) {
  module.exports = SmartDetect;
}
