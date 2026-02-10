/**
 * SumCheck - 貼上頁 UI 控制器 (Paste Tab)
 * 獨立於 UIController，管理貼上頁的 UI 互動
 */

const PasteUIController = {
    selectedMode: 'vertical_group',
    selectedIndices: [],
    selectedSigns: new Map(),
    currentErrorIndex: 0,
    recommendedMode: null,

    elements: {},

    /**
     * 初始化
     */
    init() {
        this._cacheElements();
        this._bindEvents();
        this._updateUIForMode(this.selectedMode);
    },

    /**
     * 快取 DOM
     */
    _cacheElements() {
        this.elements = {
            // 貼上區
            pasteArea: document.getElementById('p-pasteArea'),
            pasteInput: document.getElementById('p-pasteInput'),
            btnRepaste: document.getElementById('p-btnRepaste'),

            // 工具列
            calcMode: document.getElementById('p-calcMode'),
            btnRangeConfig: document.getElementById('p-btnRangeConfig'),
            btnLogicConfig: document.getElementById('p-btnLogicConfig'),
            btnValidate: document.getElementById('p-btnValidate'),
            btnDownload: document.getElementById('p-btnDownload'),
            btnReset: document.getElementById('p-btnReset'),
            btnModeHelp: document.getElementById('p-btnModeHelp'),

            // 配置面板
            rangeConfigPanel: document.getElementById('p-rangeConfigPanel'),
            logicConfigPanel: document.getElementById('p-logicConfigPanel'),

            // 模式說明
            modeHelpPanel: document.getElementById('p-modeHelpPanel'),
            btnCloseModeHelp: document.getElementById('p-btnCloseModeHelp'),
            modeCards: document.querySelectorAll('#paste-page .mode-card'),

            // 範圍設定
            headerRow: document.getElementById('p-headerRow'),
            endRow: document.getElementById('p-endRow'),
            startCol: document.getElementById('p-startCol'),
            endCol: document.getElementById('p-endCol'),

            // 關鍵字設定
            keyword1: document.getElementById('p-keyword1'),
            keyword2: document.getElementById('p-keyword2'),
            sumDirection: document.getElementById('p-sumDirection'),
            sumDirectionGroup: document.getElementById('p-sumDirectionGroup'),

            // 智能偵測
            smartDetectPanel: document.getElementById('p-smartDetectPanel'),
            smartDetectResult: document.getElementById('p-smartDetectResult'),
            btnApplyRecommend: document.getElementById('p-btnApplyRecommend'),

            // 表格
            gridContainer: document.getElementById('p-gridContainer'),

            // 操作提示
            logicHint: document.getElementById('p-logicHint'),
            logicHintText: document.getElementById('p-logicHintText'),

            // 錯誤面板
            errorPanel: document.getElementById('p-errorPanel'),
            errorCount: document.getElementById('p-errorCount'),
            errorDiff: document.getElementById('p-errorDiff'),
            btnPrevError: document.getElementById('p-btnPrevError'),
            btnNextError: document.getElementById('p-btnNextError'),

            // 載入中
            loadingOverlay: document.getElementById('p-loadingOverlay'),

            // Toast (共用主頁的 toast container)
            toastContainer: document.getElementById('toastContainer'),
        };
    },

    /**
     * 綁定事件
     */
    _bindEvents() {
        // 貼上事件
        this.elements.pasteInput?.addEventListener('paste', (e) => {
            setTimeout(() => {
                const text = this.elements.pasteInput.value;
                if (text && text.trim()) {
                    const success = PasteApp.loadPastedData(text);
                    if (success) {
                        // 隱藏貼上區，顯示表格
                        this.elements.pasteArea?.classList.add('hidden');
                        this.elements.gridContainer?.classList.remove('hidden');
                    }
                }
            }, 50);
        });

        // 重新貼上
        this.elements.btnRepaste?.addEventListener('click', () => {
            PasteApp.reset();
            this.elements.pasteArea?.classList.remove('hidden');
            this.elements.gridContainer?.classList.add('hidden');
            this.elements.pasteInput.value = '';
            this.elements.pasteInput.focus();
        });

        // 模式選擇
        this.elements.calcMode?.addEventListener('change', (e) => {
            this.selectMode(e.target.value);
        });

        // Toggle 配置面板
        this.elements.btnRangeConfig?.addEventListener('click', () => {
            this._togglePanel('p-rangeConfigPanel');
            this.elements.btnRangeConfig.classList.toggle('active');
        });

        this.elements.btnLogicConfig?.addEventListener('click', () => {
            this._togglePanel('p-logicConfigPanel');
            this.elements.btnLogicConfig.classList.toggle('active');
        });

        // 模式說明
        this.elements.btnModeHelp?.addEventListener('click', () => {
            this.elements.modeHelpPanel?.classList.toggle('hidden');
            this.elements.modeCards?.forEach(card => {
                card.classList.toggle('selected', card.dataset.mode === this.selectedMode);
            });
        });

        this.elements.btnCloseModeHelp?.addEventListener('click', () => {
            this.elements.modeHelpPanel?.classList.add('hidden');
        });

        // 模式卡片點擊
        this.elements.modeCards?.forEach(card => {
            card.addEventListener('click', () => {
                const mode = card.dataset.mode;
                if (mode) {
                    this.selectMode(mode);
                    this.elements.calcMode.value = mode;
                    this.elements.modeHelpPanel?.classList.add('hidden');
                    this.showToast('success', `已切換至：${card.querySelector('.mode-card-title').textContent}`);
                }
            });
        });

        // 驗算
        this.elements.btnValidate?.addEventListener('click', () => {
            PasteApp.runValidation();
        });

        // 下載報告
        this.elements.btnDownload?.addEventListener('click', () => {
            PasteApp.downloadReport();
        });

        // 重置
        this.elements.btnReset?.addEventListener('click', () => {
            PasteApp.reset();
            this.elements.pasteArea?.classList.remove('hidden');
            this.elements.gridContainer?.classList.add('hidden');
            this.elements.pasteInput.value = '';
        });

        // 錯誤導航
        this.elements.btnPrevError?.addEventListener('click', () => {
            this.navigateError('prev');
        });
        this.elements.btnNextError?.addEventListener('click', () => {
            this.navigateError('next');
        });

        // 設定變更時重新渲染
        ['p-headerRow', 'p-endRow', 'p-startCol', 'p-endCol'].forEach(id => {
            document.getElementById(id)?.addEventListener('change', () => {
                PasteApp.renderGrid();
            });
        });

        // 套用推薦
        this.elements.btnApplyRecommend?.addEventListener('click', () => {
            if (this.recommendedMode) {
                this.selectMode(this.recommendedMode);
                this.elements.calcMode.value = this.recommendedMode;
            }
        });
    },

    /**
     * Toggle 面板
     */
    _togglePanel(panelId) {
        const panel = document.getElementById(panelId);
        if (panel) panel.classList.toggle('hidden');
    },

    /**
     * 選擇模式
     */
    selectMode(mode) {
        this.selectedMode = mode;
        this._updateUIForMode(mode);
        if (mode === 'vertical_group' || mode === 'vertical_indent') {
            this.selectedIndices = [];
        }
        PasteApp.setMode(mode);
    },

    /**
     * 根據模式更新 UI
     */
    _updateUIForMode(mode) {
        const isKeywordMode = ['vertical_group', 'horizontal_group'].includes(mode);
        const isManualMode = ['horizontal', 'vertical_row'].includes(mode);

        if (this.elements.btnLogicConfig) {
            this.elements.btnLogicConfig.style.display = isKeywordMode ? '' : 'none';
        }
        if (this.elements.sumDirectionGroup) {
            this.elements.sumDirectionGroup.style.display = mode === 'vertical_group' ? '' : 'none';
        }
        if (this.elements.logicHint) {
            if (isManualMode && PasteApp.isLoaded) {
                this.elements.logicHint.classList.remove('hidden');
                const hintText = mode === 'horizontal'
                    ? '請點選欄位標題以設定驗算邏輯 (第一個點選 = 結果欄)'
                    : '請點選列號以設定驗算邏輯 (第一個點選 = 結果列)';
                if (this.elements.logicHintText) this.elements.logicHintText.textContent = hintText;
            } else {
                this.elements.logicHint.classList.add('hidden');
            }
        }
    },

    /**
     * 顯示智能偵測
     */
    showSmartDetection(result) {
        const panel = this.elements.smartDetectPanel;
        if (!panel) return;

        if (!result || result.confidence < 30) {
            panel.classList.add('hidden');
            return;
        }

        this.recommendedMode = result.mode;
        const modeInfo = SmartDetect.getModeInfo(result.mode);

        if (this.elements.smartDetectResult) {
            this.elements.smartDetectResult.textContent = `${modeInfo.name} (${result.confidence}% 信心度) - ${result.reasons.join('、')}`;
        }

        panel.classList.remove('hidden');

        if (result.confidence >= 70) {
            this.selectMode(result.mode);
            this.elements.calcMode.value = result.mode;
            this.showToast('success', `智能推薦已套用：${modeInfo.name}`);
            panel.classList.add('hidden');
        }
    },

    /**
     * 更新錯誤面板
     */
    updateErrorPanel(results) {
        const panel = this.elements.errorPanel;
        if (!panel) return;

        if (!results.hasErrors) {
            panel.classList.add('hidden');
            return;
        }

        panel.classList.remove('hidden');
        if (this.elements.errorCount) this.elements.errorCount.textContent = results.errorCount;
        if (this.elements.errorDiff) {
            const diff = results.totalDiff;
            this.elements.errorDiff.textContent = (diff >= 0 ? '+' : '') + diff.toLocaleString();
        }
    },

    /**
     * 導航錯誤
     */
    navigateError(direction) {
        const errors = Array.from(PasteValidator.errors.values());
        if (errors.length === 0) return;

        if (direction === 'next') {
            this.currentErrorIndex = (this.currentErrorIndex + 1) % errors.length;
        } else {
            this.currentErrorIndex = (this.currentErrorIndex - 1 + errors.length) % errors.length;
        }

        const error = errors[this.currentErrorIndex];
        const cell = document.querySelector(`#paste-page [data-row="${error.row}"][data-col="${error.col}"]`);
        if (cell) {
            cell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
            cell.classList.add('highlight');
            setTimeout(() => cell.classList.remove('highlight'), 2000);
        }
    },

    /**
     * 切換選取
     */
    toggleSelection(index) {
        const pos = this.selectedIndices.indexOf(index);
        if (pos > -1) {
            this.selectedIndices.splice(pos, 1);
            this.selectedSigns.delete(index);
        } else {
            this.selectedIndices.push(index);
            this.selectedSigns.set(index, 1);
        }
        PasteApp.renderGrid();
    },

    /**
     * 切換正負號
     */
    toggleSign(index, event) {
        if (event) event.stopPropagation();
        const current = this.selectedSigns.get(index) || 1;
        this.selectedSigns.set(index, current * -1);
        PasteApp.renderGrid();
    },

    /**
     * Loading
     */
    showLoading(show) {
        this.elements.loadingOverlay?.classList.toggle('hidden', !show);
    },

    /**
     * Toast (共用主頁容器)
     */
    showToast(type, message, duration = 3000) {
        const container = this.elements.toastContainer;
        if (!container) return;

        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.innerHTML = `
            <span class="toast-icon">${type === 'success' ? '✅' : '❌'}</span>
            <span class="toast-message">${message}</span>
        `;
        container.appendChild(toast);
        setTimeout(() => {
            toast.style.animation = 'slideOut 0.3s ease forwards';
            setTimeout(() => toast.remove(), 300);
        }, duration);
    },

    /**
     * 重置
     */
    reset() {
        this.selectedIndices = [];
        this.selectedSigns.clear();
        this.currentErrorIndex = 0;
        this.elements.errorPanel?.classList.add('hidden');
    },

    /**
     * 取得設定值
     */
    getSettings() {
        return {
            mode: this.selectedMode,
            headerRow: parseInt(this.elements.headerRow?.value) || 1,
            endRow: parseInt(this.elements.endRow?.value) || null,
            startCol: parseInt(this.elements.startCol?.value) || 1,
            endCol: parseInt(this.elements.endCol?.value) || null,
            keywords: {
                trigger: (this.elements.keyword1?.value || '').split(/[,，]/).map(k => k.trim()).filter(k => k),
                exclude: (this.elements.keyword2?.value || '').split(/[,，]/).map(k => k.trim()).filter(k => k),
            },
            sumDirection: this.elements.sumDirection?.value || 'top',
            selectedIndices: this.selectedIndices,
            selectedSigns: this.selectedSigns,
        };
    },
};
