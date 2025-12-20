/**
 * Excel 驗算大師 - 驗算邏輯模組
 * 核心驗算引擎
 */

const Validator = {
    errors: new Map(),
    corrections: new Map(),

    /**
     * 執行驗算
     * @param {Object} options - 驗算選項
     * @returns {Object} 驗算結果
     */
    validate(options) {
        this.errors.clear();
        this.corrections.clear();

        const {
            mode,
            sheetData,
            headerRow,
            endRow,
            startCol,
            endCol,
            selectedIndices,
            selectedSigns,
            keywords,
            sumDirection,
        } = options;

        const hRowIdx = headerRow - 1;
        const sColIdx = startCol - 1;
        const maxRow = endRow || sheetData.length;
        const maxCol = endCol || (sheetData[hRowIdx]?.length || 0);
        const headers = sheetData[hRowIdx] || [];

        switch (mode) {
            case 'vertical_group':
                this._validateVerticalGroup({
                    sheetData, hRowIdx, maxRow, sColIdx, maxCol,
                    nameColIndex: selectedIndices[0],
                    triggerKeywords: keywords.trigger,
                    excludeKeywords: keywords.exclude,
                    sumDirection,
                });
                break;

            case 'horizontal_group':
                this._validateHorizontalGroup({
                    sheetData, headers, hRowIdx, maxRow, sColIdx, maxCol,
                    triggerKeywords: keywords.trigger,
                    excludeKeywords: keywords.exclude,
                });
                break;

            case 'vertical_indent':
                this._validateVerticalIndent({
                    sheetData, hRowIdx, maxRow, sColIdx, maxCol,
                    nameColIndex: selectedIndices[0],
                });
                break;

            case 'horizontal':
                this._validateHorizontalManual({
                    sheetData, hRowIdx, maxRow,
                    selectedIndices, selectedSigns,
                });
                break;

            case 'vertical_row':
                this._validateVerticalManual({
                    sheetData, sColIdx, maxCol,
                    selectedIndices, selectedSigns,
                });
                break;
        }

        return this.getResults();
    },

    /**
     * 縱向關鍵字分組驗算
     */
    _validateVerticalGroup(opts) {
        const { sheetData, hRowIdx, maxRow, sColIdx, maxCol, nameColIndex,
            triggerKeywords, excludeKeywords, sumDirection } = opts;

        for (let c = sColIdx; c < maxCol; c++) {
            if (c === nameColIndex) continue;

            let tempSum = 0;
            let pendingCheck = null;

            for (let r = hRowIdx + 1; r < maxRow; r++) {
                const row = sheetData[r];
                if (!row) continue;

                const name = String(row[nameColIndex] || '');
                const val = this._parseNumber(row[c]);
                const isNum = !isNaN(val);

                const isTrigger = triggerKeywords.some(k => name.includes(k));
                const isExclude = excludeKeywords.some(k => name.includes(k));

                if (isExclude) continue;

                if (isTrigger) {
                    if (sumDirection === 'bottom') {
                        // 加總在下方：先累計，遇到關鍵字時驗算
                        if (isNum) this._checkValue(r, c, tempSum, val);
                        tempSum = 0;
                    } else {
                        // 加總在上方：遇到關鍵字時記錄待驗算
                        if (pendingCheck !== null) {
                            this._checkValue(pendingCheck.row, c, tempSum, pendingCheck.target);
                        }
                        tempSum = 0;
                        pendingCheck = isNum ? { row: r, target: val } : null;
                    }
                } else {
                    if (isNum) tempSum += val;
                }
            }

            // 處理最後一組（加總在上方）
            if (sumDirection === 'top' && pendingCheck !== null) {
                this._checkValue(pendingCheck.row, c, tempSum, pendingCheck.target);
            }
        }
    },

    /**
     * 橫向關鍵字分組驗算
     */
    _validateHorizontalGroup(opts) {
        const { sheetData, headers, hRowIdx, maxRow, sColIdx, maxCol,
            triggerKeywords, excludeKeywords } = opts;

        for (let r = hRowIdx + 1; r < maxRow; r++) {
            const row = sheetData[r];
            if (!row) continue;

            let tempSum = 0;

            for (let c = sColIdx; c < maxCol; c++) {
                const headerName = String(headers[c] || '');
                const val = this._parseNumber(row[c]);
                const isNum = !isNaN(val);

                if (triggerKeywords.some(k => headerName.includes(k))) {
                    if (isNum) this._checkValue(r, c, tempSum, val);
                    tempSum = 0;
                } else if (!excludeKeywords.some(k => headerName.includes(k))) {
                    if (isNum) tempSum += val;
                }
            }
        }
    },

    /**
     * 縱向縮排驗算
     */
    _validateVerticalIndent(opts) {
        const { sheetData, hRowIdx, maxRow, sColIdx, maxCol, nameColIndex } = opts;

        // 計算每列的縮排層級
        const rowLevels = [];
        for (let r = hRowIdx + 1; r < maxRow; r++) {
            const name = sheetData[r] ? String(sheetData[r][nameColIndex] || '') : '';
            rowLevels[r] = this._countLeadingSpaces(name);
        }

        // 對每個數值欄位進行驗算
        for (let c = sColIdx; c < maxCol; c++) {
            if (c === nameColIndex) continue;

            for (let r = hRowIdx + 1; r < maxRow - 1; r++) {
                // 如果下一列縮排更深，表示這是一個父節點
                if (rowLevels[r + 1] > rowLevels[r]) {
                    const targetLevel = rowLevels[r + 1];
                    let childrenSum = 0;
                    let hasChild = false;

                    // 加總同層級的子項目
                    for (let k = r + 1; k < maxRow; k++) {
                        if (rowLevels[k] <= rowLevels[r]) break;
                        if (rowLevels[k] === targetLevel) {
                            const val = this._parseNumber(sheetData[k]?.[c]);
                            if (!isNaN(val)) {
                                childrenSum += val;
                                hasChild = true;
                            }
                        }
                    }

                    if (hasChild) {
                        const targetVal = this._parseNumber(sheetData[r]?.[c]);
                        if (!isNaN(targetVal)) {
                            this._checkValue(r, c, childrenSum, targetVal);
                        }
                    }
                }
            }
        }
    },

    /**
     * 橫向手動驗算 (A±B=C)
     */
    _validateHorizontalManual(opts) {
        const { sheetData, hRowIdx, maxRow, selectedIndices, selectedSigns } = opts;

        if (selectedIndices.length < 2) return;

        const targetIdx = selectedIndices[selectedIndices.length - 1];
        const inputs = selectedIndices.slice(0, -1);

        for (let r = hRowIdx + 1; r < maxRow; r++) {
            const row = sheetData[r];
            if (!row) continue;

            const targetVal = this._parseNumber(row[targetIdx]);
            if (isNaN(targetVal)) continue;

            let sum = 0;
            inputs.forEach(i => {
                const v = this._parseNumber(row[i]);
                sum += (isNaN(v) ? 0 : v) * (selectedSigns.get(i) || 1);
            });

            this._checkValue(r, targetIdx, sum, targetVal);
        }
    },

    /**
     * 縱向手動驗算 (A±B=C)
     */
    _validateVerticalManual(opts) {
        const { sheetData, sColIdx, maxCol, selectedIndices, selectedSigns } = opts;

        if (selectedIndices.length < 2) return;

        const targetIdx = selectedIndices[selectedIndices.length - 1];
        const inputs = selectedIndices.slice(0, -1);

        for (let c = sColIdx; c < maxCol; c++) {
            const targetVal = this._parseNumber(sheetData[targetIdx]?.[c]);
            if (isNaN(targetVal)) continue;

            let sum = 0;
            inputs.forEach(r => {
                const v = this._parseNumber(sheetData[r]?.[c]);
                sum += (isNaN(v) ? 0 : v) * (selectedSigns.get(r) || 1);
            });

            this._checkValue(targetIdx, c, sum, targetVal);
        }
    },

    /**
     * 檢查數值並記錄錯誤
     */
    _checkValue(row, col, calculated, actual) {
        const tolerance = 1.0; // 容差
        if (Math.abs(calculated - actual) > tolerance) {
            const key = row * 10000 + col;
            const diff = actual - calculated;
            this.errors.set(key, {
                row,
                col,
                expected: calculated,
                actual,
                diff,
                message: `應為 ${calculated.toLocaleString()} (差 ${diff >= 0 ? '+' : ''}${diff.toLocaleString()})`,
            });
            this.corrections.set(key, calculated);
        }
    },

    /**
     * 取得驗算結果
     */
    getResults() {
        const errorList = Array.from(this.errors.values());
        const totalDiff = errorList.reduce((sum, e) => sum + e.diff, 0);

        return {
            errorCount: this.errors.size,
            errors: this.errors,
            corrections: this.corrections,
            errorList,
            totalDiff,
            hasErrors: this.errors.size > 0,
        };
    },

    /**
     * 取得錯誤鍵值
     */
    getErrorKey(row, col) {
        return row * 10000 + col;
    },

    /**
     * 檢查儲存格是否有錯誤
     */
    hasError(row, col) {
        return this.errors.has(this.getErrorKey(row, col));
    },

    /**
     * 取得儲存格錯誤訊息
     */
    getErrorMessage(row, col) {
        const error = this.errors.get(this.getErrorKey(row, col));
        return error?.message || null;
    },

    /**
     * 數值解析
     */
    _parseNumber(val) {
        if (val === null || val === undefined || val === '') return NaN;
        if (typeof val === 'number') return isNaN(val) ? NaN : val;

        const str = val.toString().trim();
        if (str === '') return NaN;
        if (/[a-zA-Z\u4e00-\u9fa5]/.test(str)) return NaN;

        const validPattern = /^[-\d\.,，\(\)\$%\uFF05\s]+$/;
        if (!validPattern.test(str)) return NaN;

        let cleanStr = str.replace(/[,，\$\s]/g, '').replace(/[%％]/g, '');
        if (cleanStr.startsWith('(') && cleanStr.endsWith(')')) {
            cleanStr = '-' + cleanStr.replace(/[\(\)]/g, '');
        }

        const num = parseFloat(cleanStr);
        return isNaN(num) ? NaN : num;
    },

    /**
     * 計算前導空白數
     */
    _countLeadingSpaces(str) {
        if (!str || typeof str !== 'string') return 0;
        const match = str.match(/^[\s\u00a0\u3000]*/);
        return match ? match[0].length : 0;
    },

    /**
     * 清除狀態
     */
    reset() {
        this.errors.clear();
        this.corrections.clear();
    },
};

// 匯出模組
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Validator;
}
