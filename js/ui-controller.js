/**
 * Excel 驗算大師 - UI 控制模組
 * 管理使用者介面互動與狀態
 */

const UIController = {
    // 狀態
    currentStep: 1,
    selectedMode: null,
    selectedIndices: [],
    selectedSigns: new Map(),
    currentErrorIndex: 0,

    // DOM 快取
    elements: {},

    /**
     * 初始化 UI
     */
    init() {
        this._cacheElements();
        this._bindEvents();
        this._initTheme();
        this.showStep(1);
    },

    /**
     * 快取 DOM 元素
     */
    _cacheElements() {
        this.elements = {
            // 步驟精靈
            wizardSteps: document.querySelectorAll('.wizard-step'),
            wizardConnectors: document.querySelectorAll('.wizard-connector'),
            stepContents: document.querySelectorAll('.step-content'),

            // 上傳區
            dropzone: document.getElementById('dropzone'),
            fileInput: document.getElementById('fileInput'),
            dropOverlay: document.getElementById('dropOverlay'),

            // 設定區
            sheetSelector: document.getElementById('sheetSelector'),
            modeCards: document.querySelectorAll('.mode-card'),
            smartDetectPanel: document.getElementById('smartDetectPanel'),

            // 範圍設定
            headerRow: document.getElementById('headerRow'),
            endRow: document.getElementById('endRow'),
            startCol: document.getElementById('startCol'),
            endCol: document.getElementById('endCol'),

            // 關鍵字設定
            keyword1: document.getElementById('keyword1'),
            keyword2: document.getElementById('keyword2'),
            sumDirection: document.getElementById('sumDirection'),

            // 表格
            gridContainer: document.getElementById('gridContainer'),

            // 錯誤面板
            errorPanel: document.getElementById('errorPanel'),
            errorCount: document.getElementById('errorCount'),
            errorDiff: document.getElementById('errorDiff'),

            // 按鈕
            btnNext: document.getElementById('btnNext'),
            btnPrev: document.getElementById('btnPrev'),
            btnValidate: document.getElementById('btnValidate'),
            btnDownload: document.getElementById('btnDownload'),
            btnReset: document.getElementById('btnReset'),
            themeToggle: document.getElementById('themeToggle'),

            // 載入中
            loadingOverlay: document.getElementById('loadingOverlay'),

            // Toast
            toastContainer: document.getElementById('toastContainer'),
        };
    },

    /**
     * 綁定事件
     */
    _bindEvents() {
        // 主題切換
        this.elements.themeToggle?.addEventListener('click', () => this.toggleTheme());

        // 拖曳上傳
        this._setupDropzone();

        // 檔案選擇
        this.elements.fileInput?.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                this._handleFileUpload(e.target.files[0]);
            }
        });

        // 模式卡片選擇
        this.elements.modeCards.forEach(card => {
            card.addEventListener('click', () => {
                this.selectMode(card.dataset.mode);
            });
        });

        // 設定變更時重新渲染
        ['headerRow', 'endRow', 'startCol', 'endCol'].forEach(id => {
            document.getElementById(id)?.addEventListener('change', () => {
                if (typeof App !== 'undefined') App.renderGrid();
            });
        });

        // 步驟導航
        this.elements.btnNext?.addEventListener('click', () => this.nextStep());
        this.elements.btnPrev?.addEventListener('click', () => this.prevStep());

        // 手風琴
        document.querySelectorAll('.settings-accordion-header').forEach(header => {
            header.addEventListener('click', () => {
                header.closest('.settings-accordion')?.classList.toggle('open');
            });
        });
    },

    /**
     * 設定拖曳上傳區
     */
    _setupDropzone() {
        const dropzone = this.elements.dropzone;
        const overlay = this.elements.dropOverlay;

        if (!dropzone) return;

        // 點擊觸發檔案選擇
        dropzone.addEventListener('click', () => {
            this.elements.fileInput?.click();
        });

        // 拖曳事件
        document.addEventListener('dragover', (e) => {
            e.preventDefault();
            overlay?.classList.add('active');
        });

        overlay?.addEventListener('dragleave', (e) => {
            e.preventDefault();
            overlay?.classList.remove('active');
        });

        overlay?.addEventListener('drop', (e) => {
            e.preventDefault();
            overlay?.classList.remove('active');
            if (e.dataTransfer.files[0]) {
                this._handleFileUpload(e.dataTransfer.files[0]);
            }
        });

        // Dropzone 自身的拖曳樣式
        dropzone.addEventListener('dragover', () => dropzone.classList.add('dragover'));
        dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
        dropzone.addEventListener('drop', () => dropzone.classList.remove('dragover'));
    },

    /**
     * 處理檔案上傳
     */
    async _handleFileUpload(file) {
        this.showLoading(true);

        try {
            if (typeof App !== 'undefined') {
                await App.loadFile(file);
                this.showToast('success', `成功載入：${file.name}`);
                this.nextStep();
            }
        } catch (err) {
            this.showToast('error', '檔案載入失敗：' + err.message);
        } finally {
            this.showLoading(false);
        }
    },

    /**
     * 初始化主題
     */
    _initTheme() {
        const savedTheme = localStorage.getItem('theme') || 'light';
        document.documentElement.setAttribute('data-theme', savedTheme);
    },

    /**
     * 切換主題
     */
    toggleTheme() {
        const current = document.documentElement.getAttribute('data-theme');
        const next = current === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        localStorage.setItem('theme', next);
    },

    /**
     * 顯示步驟
     */
    showStep(step) {
        this.currentStep = step;

        // 更新步驟指示器
        this.elements.wizardSteps.forEach((el, idx) => {
            el.classList.remove('active', 'completed');
            if (idx + 1 < step) el.classList.add('completed');
            if (idx + 1 === step) el.classList.add('active');
        });

        // 更新連接線
        this.elements.wizardConnectors.forEach((el, idx) => {
            el.classList.toggle('active', idx + 1 < step);
        });

        // 顯示對應內容
        this.elements.stepContents.forEach((el, idx) => {
            el.classList.toggle('hidden', idx + 1 !== step);
        });

        // 更新導航按鈕
        if (this.elements.btnPrev) {
            this.elements.btnPrev.classList.toggle('hidden', step === 1);
        }
        if (this.elements.btnNext) {
            this.elements.btnNext.textContent = step === 3 ? '執行驗算' : '下一步';
        }
    },

    /**
     * 下一步
     */
    nextStep() {
        if (this.currentStep < 3) {
            this.showStep(this.currentStep + 1);
        } else if (this.currentStep === 3) {
            // 執行驗算
            if (typeof App !== 'undefined') App.runValidation();
        }
    },

    /**
     * 上一步
     */
    prevStep() {
        if (this.currentStep > 1) {
            this.showStep(this.currentStep - 1);
        }
    },

    /**
     * 選擇模式
     */
    selectMode(mode) {
        this.selectedMode = mode;

        // 更新卡片樣式
        this.elements.modeCards.forEach(card => {
            card.classList.toggle('selected', card.dataset.mode === mode);
        });

        // 更新設定面板
        this._updateSettingsForMode(mode);

        // 通知 App
        if (typeof App !== 'undefined') {
            App.setMode(mode);
        }
    },

    /**
     * 根據模式更新設定面板
     */
    _updateSettingsForMode(mode) {
        const keywordSection = document.getElementById('keywordSettings');
        const sumDirectionGroup = document.getElementById('sumDirectionGroup');

        if (mode === 'vertical_group' || mode === 'horizontal_group') {
            keywordSection?.classList.remove('hidden');
            sumDirectionGroup?.classList.toggle('hidden', mode !== 'vertical_group');
        } else {
            keywordSection?.classList.add('hidden');
        }
    },

    /**
     * 更新工作表選擇器
     */
    updateSheetSelector(sheetNames, currentSheet) {
        const selector = this.elements.sheetSelector;
        if (!selector) return;

        selector.innerHTML = '';
        sheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            option.selected = name === currentSheet;
            selector.appendChild(option);
        });

        selector.parentElement?.classList.remove('hidden');
    },

    /**
     * 顯示智能偵測結果
     */
    showSmartDetection(result) {
        const panel = this.elements.smartDetectPanel;
        if (!panel) return;

        if (!result || result.confidence < 30) {
            panel.classList.add('hidden');
            return;
        }

        const modeInfo = SmartDetect.getModeInfo(result.mode);

        panel.innerHTML = `
      <span class="smart-detect-icon">🤖</span>
      <div class="smart-detect-content">
        <div class="smart-detect-title">智能推薦：${modeInfo.name}</div>
        <div class="smart-detect-desc">${result.reasons.join('、')}</div>
      </div>
      <div class="smart-detect-confidence">
        <div class="confidence-bar">
          <div class="confidence-fill" style="width: ${result.confidence}%"></div>
        </div>
        <span>${result.confidence}%</span>
      </div>
    `;

        panel.classList.remove('hidden');

        // 自動選取推薦的模式
        this.selectMode(result.mode);
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

        if (this.elements.errorCount) {
            this.elements.errorCount.textContent = results.errorCount;
        }

        if (this.elements.errorDiff) {
            const diff = results.totalDiff;
            const prefix = diff >= 0 ? '+' : '';
            this.elements.errorDiff.textContent = prefix + diff.toLocaleString();
        }
    },

    /**
     * 導航到錯誤
     */
    navigateError(direction) {
        const errors = Array.from(Validator.errors.values());
        if (errors.length === 0) return;

        if (direction === 'next') {
            this.currentErrorIndex = (this.currentErrorIndex + 1) % errors.length;
        } else {
            this.currentErrorIndex = (this.currentErrorIndex - 1 + errors.length) % errors.length;
        }

        const error = errors[this.currentErrorIndex];
        this._scrollToCell(error.row, error.col);
    },

    /**
     * 滾動到指定儲存格
     */
    _scrollToCell(row, col) {
        const cell = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            cell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
            cell.classList.add('highlight');
            setTimeout(() => cell.classList.remove('highlight'), 2000);
        }
    },

    /**
     * 切換欄/列選取
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

        if (typeof App !== 'undefined') App.renderGrid();
    },

    /**
     * 切換正負號
     */
    toggleSign(index, event) {
        if (event) event.stopPropagation();

        const current = this.selectedSigns.get(index) || 1;
        this.selectedSigns.set(index, current * -1);

        if (typeof App !== 'undefined') App.renderGrid();
    },

    /**
     * 顯示/隱藏載入畫面
     */
    showLoading(show) {
        this.elements.loadingOverlay?.classList.toggle('hidden', !show);
    },

    /**
     * 顯示 Toast 通知
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
     * 重置 UI 狀態
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

// 匯出模組
if (typeof module !== 'undefined' && module.exports) {
    module.exports = UIController;
}
