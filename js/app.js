/**
 * SumCheck - 應用程式入口 (Flat Toolbar Version)
 * 整合所有模組並管理應用程式狀態
 */

const App = {
    // 狀態
    isLoaded: false,
    currentMode: 'vertical_group',

    /**
     * 初始化應用程式
     */
    init() {
        console.log('✓ SumCheck v2.1 (Flat Toolbar)');

        // 初始化 UI
        UIController.init();

        // 設定預設模式
        this.setMode('vertical_group');

        console.log('✅ 應用程式初始化完成');
    },

    /**
     * 載入 Excel 檔案
     */
    async loadFile(file) {
        const result = await ExcelParser.parseFile(file);

        if (result.success) {
            this.isLoaded = true;

            // 更新工作表選擇器
            UIController.updateSheetSelector(result.sheetNames, result.currentSheet);

            // 自動偵測資料範圍
            const range = ExcelParser.detectDataRange();
            if (range) {
                document.getElementById('headerRow').value = range.headerRow;
                document.getElementById('endRow').value = range.endRow;
                document.getElementById('startCol').value = range.startCol;
                document.getElementById('endCol').value = range.endCol;
            }

            // 執行智能偵測
            this._runSmartDetection();

            // 渲染表格
            this.renderGrid();
        }

        return result;
    },

    /**
     * 切換工作表
     */
    switchSheet(sheetName) {
        ExcelParser.loadSheet(sheetName);
        Validator.reset();
        UIController.reset();

        const range = ExcelParser.detectDataRange();
        if (range) {
            document.getElementById('headerRow').value = range.headerRow;
            document.getElementById('endRow').value = range.endRow;
        }

        this._runSmartDetection();
        this.renderGrid();
    },

    /**
     * 執行智能偵測
     */
    _runSmartDetection() {
        const data = ExcelParser.getData();
        const headerRow = parseInt(document.getElementById('headerRow')?.value) || 1;

        const result = SmartDetect.analyze(data, headerRow - 1);
        UIController.showSmartDetection(result);
    },

    /**
     * 設定驗算模式
     */
    setMode(mode) {
        this.currentMode = mode;
        UIController.selectedMode = mode;

        // 單選模式需要清除選取
        if (mode === 'vertical_group' || mode === 'vertical_indent') {
            UIController.selectedIndices = [];
        }

        this.renderGrid();
    },

    /**
     * 渲染表格
     */
    renderGrid() {
        const container = document.getElementById('gridContainer');
        if (!container) return;

        const data = ExcelParser.getData();
        if (!data || data.length === 0) {
            container.innerHTML = `
                <div class="grid-placeholder">
                    <div class="grid-placeholder-icon">📂</div>
                    <div>請先載入 Excel 檔案</div>
                </div>
            `;
            return;
        }

        const settings = UIController.getSettings();
        const { headerRow, endRow, startCol, endCol } = settings;

        const hRowIdx = headerRow - 1;
        const sColIdx = startCol - 1;
        const maxRow = endRow || data.length;
        const maxCol = endCol || (data[hRowIdx]?.length || 0);
        const headers = data[hRowIdx] || [];

        let html = '<table class="data-grid"><thead><tr class="sticky-header">';
        html += '<th class="row-header">列</th>';

        // 標題列
        for (let c = sColIdx; c < maxCol; c++) {
            const h = headers[c];
            const isSelected = UIController.selectedIndices.includes(c);
            const isClickable = ['horizontal', 'vertical_group', 'vertical_indent'].includes(this.currentMode);

            let badge = '';
            let className = isClickable ? 'clickable' : '';
            if (isSelected) {
                className += ' selected';
                if (this.currentMode === 'horizontal') {
                    const pos = UIController.selectedIndices.indexOf(c);
                    const isTarget = pos === UIController.selectedIndices.length - 1 && UIController.selectedIndices.length > 1;
                    if (isTarget) {
                        badge = '<span class="logic-tag" style="background:var(--badge-target)">=</span>';
                    } else {
                        const sign = UIController.selectedSigns.get(c) || 1;
                        const isPlus = sign === 1;
                        badge = `<span class="logic-tag interactive" style="background:${isPlus ? 'var(--badge-input)' : 'var(--badge-minus)'}" data-col="${c}" data-action="sign">${isPlus ? '+' : '-'}</span>`;
                    }
                } else {
                    badge = '<span class="logic-tag" style="background:#f59e0b">鍵</span>';
                }
            }

            const clickHandler = isClickable ? `data-col="${c}" data-action="toggle"` : '';
            html += `<th class="${className}" ${clickHandler}>${badge}${h || 'Col ' + (c + 1)}</th>`;
        }
        html += '</tr></thead><tbody>';

        // 資料列
        for (let r = hRowIdx + 1; r < maxRow; r++) {
            const row = data[r] || [];
            const isRowSelected = UIController.selectedIndices.includes(r);
            const isRowClickable = this.currentMode === 'vertical_row';

            let rowBadge = '';
            let rowClass = 'row-header' + (isRowClickable ? ' clickable' : '');

            if (isRowSelected && this.currentMode === 'vertical_row') {
                rowClass += ' selected';
                const pos = UIController.selectedIndices.indexOf(r);
                const isTarget = pos === UIController.selectedIndices.length - 1 && UIController.selectedIndices.length > 1;
                if (isTarget) {
                    rowBadge = '<span class="logic-tag" style="background:var(--badge-target)">=</span>';
                } else {
                    const sign = UIController.selectedSigns.get(r) || 1;
                    const isPlus = sign === 1;
                    rowBadge = `<span class="logic-tag interactive" style="background:${isPlus ? 'var(--badge-input)' : 'var(--badge-minus)'}" data-row="${r}" data-action="sign">${isPlus ? '+' : '-'}</span>`;
                }
            }

            const rowClickHandler = isRowClickable ? `data-row="${r}" data-action="toggle"` : '';
            html += `<tr><td class="${rowClass}" ${rowClickHandler}>${r + 1} ${rowBadge}</td>`;

            for (let c = sColIdx; c < maxCol; c++) {
                const val = row[c];
                const display = ExcelParser.formatNumber(val);
                const hasError = Validator.hasError(r, c);
                const errorMsg = Validator.getErrorMessage(r, c);

                let cellClass = '';
                let style = '';

                if (hasError) {
                    cellClass = 'err-cell';
                } else if (isRowSelected || UIController.selectedIndices.includes(c)) {
                    style = 'background: var(--primary-light);';
                }

                html += `<td class="${cellClass}" style="${style}" data-row="${r}" data-col="${c}">`;
                html += display;
                if (errorMsg) html += `<span class="err-msg">${errorMsg}</span>`;
                html += '</td>';
            }
            html += '</tr>';
        }

        html += '</tbody></table>';
        container.innerHTML = html;

        // 綁定表格事件
        this._bindGridEvents(container);
    },

    /**
     * 綁定表格事件
     */
    _bindGridEvents(container) {
        container.addEventListener('click', (e) => {
            const el = e.target.closest('[data-action]');
            if (!el) return;

            const action = el.dataset.action;
            const col = parseInt(el.dataset.col);
            const row = parseInt(el.dataset.row);

            if (action === 'toggle') {
                const idx = !isNaN(col) ? col : row;

                // 單選模式
                if (this.currentMode === 'vertical_group' || this.currentMode === 'vertical_indent') {
                    UIController.selectedIndices = [idx];
                } else {
                    UIController.toggleSelection(idx);
                }
                this.renderGrid();
            } else if (action === 'sign') {
                const idx = !isNaN(col) ? col : row;
                UIController.toggleSign(idx, e);
                this.renderGrid();
            }
        });
    },

    /**
     * 執行驗算
     */
    runValidation() {
        const data = ExcelParser.getData();
        if (!data || data.length === 0) {
            UIController.showToast('error', '請先載入 Excel 檔案');
            return;
        }

        const settings = UIController.getSettings();

        // 驗證設定
        if (['vertical_group', 'vertical_indent'].includes(this.currentMode)) {
            if (UIController.selectedIndices.length !== 1) {
                UIController.showToast('error', '請先選取名稱欄');
                return;
            }
        }

        if (['horizontal', 'vertical_row'].includes(this.currentMode)) {
            if (UIController.selectedIndices.length < 2) {
                UIController.showToast('error', '請至少選取 2 個欄位/列');
                return;
            }
        }

        UIController.showLoading(true);

        // 延遲執行，讓 loading 顯示
        setTimeout(() => {
            try {
                const results = Validator.validate({
                    mode: this.currentMode,
                    sheetData: data,
                    headerRow: settings.headerRow,
                    endRow: settings.endRow,
                    startCol: settings.startCol,
                    endCol: settings.endCol,
                    selectedIndices: UIController.selectedIndices,
                    selectedSigns: UIController.selectedSigns,
                    keywords: settings.keywords,
                    sumDirection: settings.sumDirection,
                });

                this.renderGrid();
                UIController.updateErrorPanel(results);

                if (results.hasErrors) {
                    UIController.showToast('error', `發現 ${results.errorCount} 個錯誤`);
                } else {
                    UIController.showToast('success', '✅ 驗算成功！無錯誤');
                }
            } catch (err) {
                UIController.showToast('error', '驗算失敗：' + err.message);
            } finally {
                UIController.showLoading(false);
            }
        }, 100);
    },

    /**
     * 下載報告
     */
    downloadReport() {
        if (Validator.errors.size === 0) {
            UIController.showToast('error', '無錯誤可匯出');
            return;
        }

        const data = ExcelParser.getData();
        const settings = UIController.getSettings();
        const hRowIdx = settings.headerRow - 1;

        ReportGenerator.download({
            sheetData: data,
            headers: data[hRowIdx] || [],
            errors: Validator.errors,
            corrections: Validator.corrections,
            headerRowIndex: hRowIdx,
            startColIndex: settings.startCol - 1,
            endRowIndex: (settings.endRow || data.length) - 1,
            endColIndex: (settings.endCol || data[hRowIdx]?.length || 0) - 1,
        });

        UIController.showToast('success', '報告已下載');
    },

    /**
     * 重置應用程式狀態
     */
    reset() {
        Validator.reset();
        UIController.reset();
        this.renderGrid();
        UIController.showToast('success', '已重置');
    },
};

// 等待 DOM 載入後初始化
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// 匯出模組
if (typeof module !== 'undefined' && module.exports) {
    module.exports = App;
}
