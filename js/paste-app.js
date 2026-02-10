/**
 * SumCheck - 貼上頁應用程式 (Paste Tab)
 * 獨立運作的貼上內容驗算模組
 */

// 建立獨立的 Validator 實例，避免與 Tab 1 互相汙染
const PasteValidator = Object.create(Validator);
PasteValidator.errors = new Map();
PasteValidator.corrections = new Map();

const PasteApp = {
    isLoaded: false,
    currentMode: 'vertical_group',
    sheetData: [],

    /**
     * 初始化
     */
    init() {
        PasteUIController.init();
        this.setMode('vertical_group');
        console.log('✓ PasteApp initialized');
    },

    /**
     * 解析貼上的文字為 2D 陣列
     * Excel 複製的內容格式：tab 分隔欄、換行分隔列
     */
    parseClipboardText(text) {
        if (!text || !text.trim()) return null;

        const lines = text.split(/\r\n|\n|\r/).filter(line => line.trim() !== '');
        if (lines.length === 0) return null;

        const data = lines.map(line => {
            return line.split('\t').map(cell => {
                const trimmed = cell.trim();
                // 嘗試轉數字（移除千分位逗號）
                const cleaned = trimmed.replace(/,/g, '');
                const num = parseFloat(cleaned);
                // 如果是純數字（含負數、小數），轉為 number
                if (!isNaN(num) && /^[-\d.,()$%\s]+$/.test(trimmed)) {
                    return num;
                }
                return trimmed;
            });
        });

        return data;
    },

    /**
     * 載入貼上的資料
     */
    loadPastedData(text) {
        const data = this.parseClipboardText(text);
        if (!data || data.length === 0) {
            PasteUIController.showToast('error', '無法解析貼上的內容');
            return false;
        }

        this.sheetData = data;
        this.isLoaded = true;

        // 自動偵測資料範圍
        const range = this._detectDataRange(data);
        if (range) {
            document.getElementById('p-headerRow').value = range.headerRow;
            document.getElementById('p-endRow').value = range.endRow;
            document.getElementById('p-startCol').value = range.startCol;
            document.getElementById('p-endCol').value = range.endCol;
        }

        // 智能偵測
        this._runSmartDetection();

        // 渲染表格
        this.renderGrid();

        PasteUIController.showToast('success', `成功載入 ${data.length} 列 × ${data[0]?.length || 0} 欄`);
        return true;
    },

    /**
     * 自動偵測資料範圍
     */
    _detectDataRange(data) {
        if (!data || !data.length) return null;

        let headerRow = 1;
        let endRow = data.length;
        let endCol = 0;

        data.forEach(row => {
            if (row?.length > endCol) endCol = row.length;
        });

        // 找標題列
        for (let r = 0; r < Math.min(10, data.length); r++) {
            const row = data[r];
            if (!row) continue;
            const nonEmpty = row.filter(c => c !== null && c !== undefined && c !== '').length;
            if (nonEmpty >= 3) {
                headerRow = r + 1;
                break;
            }
        }

        return { headerRow, endRow, startCol: 1, endCol };
    },

    /**
     * 執行智能偵測
     */
    _runSmartDetection() {
        const headerRow = parseInt(document.getElementById('p-headerRow')?.value) || 1;
        if (typeof SmartDetect !== 'undefined') {
            const result = SmartDetect.analyze(this.sheetData, headerRow - 1);
            PasteUIController.showSmartDetection(result);
        }
    },

    /**
     * 設定驗算模式
     */
    setMode(mode) {
        this.currentMode = mode;
        PasteUIController.selectedMode = mode;
        if (mode === 'vertical_group' || mode === 'vertical_indent') {
            PasteUIController.selectedIndices = [];
        }
        this.renderGrid();
    },

    /**
     * 渲染表格
     */
    renderGrid() {
        const container = document.getElementById('p-gridContainer');
        if (!container) return;

        const data = this.sheetData;
        if (!data || data.length === 0) {
            container.innerHTML = `
                <div class="grid-placeholder">
                    <div class="grid-placeholder-icon">📋</div>
                    <div>請先貼上 Excel 內容</div>
                </div>
            `;
            return;
        }

        const settings = PasteUIController.getSettings();
        const { headerRow, endRow, startCol, endCol } = settings;
        const hRowIdx = headerRow - 1;
        const sColIdx = startCol - 1;
        const maxRow = endRow || data.length;
        const maxCol = endCol || (data[hRowIdx]?.length || 0);
        const headers = data[hRowIdx] || [];

        let html = '<table class="data-grid"><thead><tr class="sticky-header">';
        html += '<th class="row-header">列</th>';

        for (let c = sColIdx; c < maxCol; c++) {
            const h = headers[c];
            const isSelected = PasteUIController.selectedIndices.includes(c);
            const isClickable = ['horizontal', 'vertical_group', 'vertical_indent'].includes(this.currentMode);

            let badge = '';
            let className = isClickable ? 'clickable' : '';
            if (isSelected) {
                className += ' selected';
                if (this.currentMode === 'horizontal') {
                    const pos = PasteUIController.selectedIndices.indexOf(c);
                    const isTarget = pos === PasteUIController.selectedIndices.length - 1 && PasteUIController.selectedIndices.length > 1;
                    if (isTarget) {
                        badge = '<span class="logic-tag" style="background:var(--badge-target)">=</span>';
                    } else {
                        const sign = PasteUIController.selectedSigns.get(c) || 1;
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

        for (let r = hRowIdx + 1; r < maxRow; r++) {
            const row = data[r] || [];
            const isRowSelected = PasteUIController.selectedIndices.includes(r);
            const isRowClickable = this.currentMode === 'vertical_row';

            let rowBadge = '';
            let rowClass = 'row-header' + (isRowClickable ? ' clickable' : '');

            if (isRowSelected && this.currentMode === 'vertical_row') {
                rowClass += ' selected';
                const pos = PasteUIController.selectedIndices.indexOf(r);
                const isTarget = pos === PasteUIController.selectedIndices.length - 1 && PasteUIController.selectedIndices.length > 1;
                if (isTarget) {
                    rowBadge = '<span class="logic-tag" style="background:var(--badge-target)">=</span>';
                } else {
                    const sign = PasteUIController.selectedSigns.get(r) || 1;
                    const isPlus = sign === 1;
                    rowBadge = `<span class="logic-tag interactive" style="background:${isPlus ? 'var(--badge-input)' : 'var(--badge-minus)'}" data-row="${r}" data-action="sign">${isPlus ? '+' : '-'}</span>`;
                }
            }

            const rowClickHandler = isRowClickable ? `data-row="${r}" data-action="toggle"` : '';
            html += `<tr><td class="${rowClass}" ${rowClickHandler}>${r + 1} ${rowBadge}</td>`;

            for (let c = sColIdx; c < maxCol; c++) {
                const val = row[c];
                const display = ExcelParser.formatNumber(val);
                const hasError = PasteValidator.hasError(r, c);
                const errorMsg = PasteValidator.getErrorMessage(r, c);

                let cellClass = '';
                let style = '';

                if (hasError) {
                    cellClass = 'err-cell';
                } else if (isRowSelected || PasteUIController.selectedIndices.includes(c)) {
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
                if (this.currentMode === 'vertical_group' || this.currentMode === 'vertical_indent') {
                    PasteUIController.selectedIndices = [idx];
                } else {
                    PasteUIController.toggleSelection(idx);
                }
                this.renderGrid();
            } else if (action === 'sign') {
                const idx = !isNaN(col) ? col : row;
                PasteUIController.toggleSign(idx, e);
                this.renderGrid();
            }
        });
    },

    /**
     * 執行驗算
     */
    runValidation() {
        if (!this.sheetData || this.sheetData.length === 0) {
            PasteUIController.showToast('error', '請先貼上資料');
            return;
        }

        const settings = PasteUIController.getSettings();

        if (['vertical_group', 'vertical_indent'].includes(this.currentMode)) {
            if (PasteUIController.selectedIndices.length !== 1) {
                PasteUIController.showToast('error', '請先選取名稱欄');
                return;
            }
        }

        if (['horizontal', 'vertical_row'].includes(this.currentMode)) {
            if (PasteUIController.selectedIndices.length < 2) {
                PasteUIController.showToast('error', '請至少選取 2 個欄位/列');
                return;
            }
        }

        PasteUIController.showLoading(true);

        setTimeout(() => {
            try {
                const results = PasteValidator.validate({
                    mode: this.currentMode,
                    sheetData: this.sheetData,
                    headerRow: settings.headerRow,
                    endRow: settings.endRow,
                    startCol: settings.startCol,
                    endCol: settings.endCol,
                    selectedIndices: PasteUIController.selectedIndices,
                    selectedSigns: PasteUIController.selectedSigns,
                    keywords: settings.keywords,
                    sumDirection: settings.sumDirection,
                });

                this.renderGrid();
                PasteUIController.updateErrorPanel(results);

                if (results.hasErrors) {
                    PasteUIController.showToast('error', `發現 ${results.errorCount} 個錯誤`);
                } else {
                    PasteUIController.showToast('success', '✅ 驗算成功！無錯誤');
                }
            } catch (err) {
                PasteUIController.showToast('error', '驗算失敗：' + err.message);
            } finally {
                PasteUIController.showLoading(false);
            }
        }, 100);
    },

    /**
     * 下載報告
     */
    downloadReport() {
        if (PasteValidator.errors.size === 0) {
            PasteUIController.showToast('error', '無錯誤可匯出');
            return;
        }

        const settings = PasteUIController.getSettings();
        const hRowIdx = settings.headerRow - 1;

        if (typeof ReportGenerator !== 'undefined') {
            ReportGenerator.download({
                sheetData: this.sheetData,
                headers: this.sheetData[hRowIdx] || [],
                errors: PasteValidator.errors,
                corrections: PasteValidator.corrections,
                headerRowIndex: hRowIdx,
                startColIndex: settings.startCol - 1,
                endRowIndex: (settings.endRow || this.sheetData.length) - 1,
                endColIndex: (settings.endCol || this.sheetData[hRowIdx]?.length || 0) - 1,
            });
            PasteUIController.showToast('success', '報告已下載');
        }
    },

    /**
     * 重置
     */
    reset() {
        PasteValidator.reset();
        PasteUIController.reset();
        this.sheetData = [];
        this.isLoaded = false;
        this.renderGrid();
        PasteUIController.showToast('success', '已重置');
    },
};
