/**
 * Excel 驗算大師 - UI 控制模組 (Flat Toolbar Version)
 * 管理使用者介面互動與狀態
 */

const UIController = {
  // 狀態
  selectedMode: "vertical_group",
  selectedIndices: [],
  selectedSigns: new Map(),
  currentErrorIndex: 0,
  isFileLoaded: false,
  recommendedMode: null,

  // DOM 快取
  elements: {},

  /**
   * 初始化 UI
   */
  init() {
    this._cacheElements();
    this._bindEvents();
    this._initTheme();
    this._updateUIForMode(this.selectedMode);
  },

  /**
   * 快取 DOM 元素
   */
  _cacheElements() {
    this.elements = {
      // 工具列
      btnLoadFile: document.getElementById("btnLoadFile"),
      fileInput: document.getElementById("fileInput"),
      sheetGroup: document.getElementById("sheetGroup"),
      sheetSelector: document.getElementById("sheetSelector"),
      calcMode: document.getElementById("calcMode"),
      btnRangeConfig: document.getElementById("btnRangeConfig"),
      btnLogicConfig: document.getElementById("btnLogicConfig"),
      btnValidate: document.getElementById("btnValidate"),
      btnDownload: document.getElementById("btnDownload"),
      btnReset: document.getElementById("btnReset"),
      themeToggle: document.getElementById("themeToggle"),

      // 配置面板
      rangeConfigPanel: document.getElementById("rangeConfigPanel"),
      logicConfigPanel: document.getElementById("logicConfigPanel"),

      // 模式說明
      btnModeHelp: document.getElementById("btnModeHelp"),
      modeHelpPanel: document.getElementById("modeHelpPanel"),
      btnCloseModeHelp: document.getElementById("btnCloseModeHelp"),
      modeCards: document.querySelectorAll(".mode-card"),

      // 範圍設定
      headerRow: document.getElementById("headerRow"),
      endRow: document.getElementById("endRow"),
      startCol: document.getElementById("startCol"),
      endCol: document.getElementById("endCol"),

      // 關鍵字設定
      keyword1: document.getElementById("keyword1"),
      keyword2: document.getElementById("keyword2"),
      sumDirection: document.getElementById("sumDirection"),
      sumDirectionGroup: document.getElementById("sumDirectionGroup"),

      // 智能偵測
      smartDetectPanel: document.getElementById("smartDetectPanel"),
      smartDetectResult: document.getElementById("smartDetectResult"),
      btnApplyRecommend: document.getElementById("btnApplyRecommend"),

      // 上傳區
      dropzone: document.getElementById("dropzone"),
      dropOverlay: document.getElementById("dropOverlay"),

      // 表格
      gridContainer: document.getElementById("gridContainer"),

      // 操作提示
      logicHint: document.getElementById("logicHint"),
      logicHintText: document.getElementById("logicHintText"),

      // 錯誤面板
      errorPanel: document.getElementById("errorPanel"),
      errorCount: document.getElementById("errorCount"),
      errorDiff: document.getElementById("errorDiff"),
      btnPrevError: document.getElementById("btnPrevError"),
      btnNextError: document.getElementById("btnNextError"),

      // 載入中
      loadingOverlay: document.getElementById("loadingOverlay"),

      // Toast
      toastContainer: document.getElementById("toastContainer"),
    };
  },

  /**
   * 綁定事件
   */
  _bindEvents() {
    // 載入檔案按鈕
    this.elements.btnLoadFile?.addEventListener("click", () => {
      this.elements.fileInput?.click();
    });

    // 檔案選擇
    this.elements.fileInput?.addEventListener("change", (e) => {
      if (e.target.files[0]) {
        this._handleFileUpload(e.target.files[0]);
      }
    });

    // 主題切換
    this.elements.themeToggle?.addEventListener("click", () =>
      this.toggleTheme(),
    );

    // 拖曳上傳
    this._setupDropzone();

    // 模式選擇
    this.elements.calcMode?.addEventListener("change", (e) => {
      this.selectMode(e.target.value);
    });

    // Toggle 配置面板
    this.elements.btnRangeConfig?.addEventListener("click", () => {
      this.togglePanel("rangeConfigPanel");
      this.elements.btnRangeConfig.classList.toggle("active");
    });

    this.elements.btnLogicConfig?.addEventListener("click", () => {
      this.togglePanel("logicConfigPanel");
      this.elements.btnLogicConfig.classList.toggle("active");
    });

    // 模式說明面板
    this.elements.btnModeHelp?.addEventListener("click", () => {
      this.toggleModeHelp();
    });

    this.elements.btnCloseModeHelp?.addEventListener("click", () => {
      this.closeModeHelp();
    });

    // 模式卡片點擊選擇
    this.elements.modeCards?.forEach((card) => {
      card.addEventListener("click", () => {
        const mode = card.dataset.mode;
        if (mode) {
          this.selectMode(mode);
          this.elements.calcMode.value = mode;
          this.closeModeHelp();
          this.showToast("success", `已切換至：${card.querySelector(".mode-card-title").textContent}`);
        }
      });
    });

    // 驗算按鈕
    this.elements.btnValidate?.addEventListener("click", () => {
      if (typeof App !== "undefined") App.runValidation();
    });

    // 下載報告
    this.elements.btnDownload?.addEventListener("click", () => {
      if (typeof App !== "undefined") App.downloadReport();
    });

    // 重置
    this.elements.btnReset?.addEventListener("click", () => {
      if (typeof App !== "undefined") App.reset();
    });

    // 錯誤導航
    this.elements.btnPrevError?.addEventListener("click", () => {
      this.navigateError("prev");
    });
    this.elements.btnNextError?.addEventListener("click", () => {
      this.navigateError("next");
    });

    // 工作表切換
    this.elements.sheetSelector?.addEventListener("change", (e) => {
      if (typeof App !== "undefined") App.switchSheet(e.target.value);
    });

    // 設定變更時重新渲染
    ["headerRow", "endRow", "startCol", "endCol"].forEach((id) => {
      document.getElementById(id)?.addEventListener("change", () => {
        if (typeof App !== "undefined") App.renderGrid();
      });
    });

    // 套用推薦
    this.elements.btnApplyRecommend?.addEventListener("click", () => {
      if (this.recommendedMode) {
        this.selectMode(this.recommendedMode);
        this.elements.calcMode.value = this.recommendedMode;
      }
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
    dropzone.addEventListener("click", () => {
      this.elements.fileInput?.click();
    });

    // 拖曳事件
    document.addEventListener("dragover", (e) => {
      e.preventDefault();
      overlay?.classList.add("active");
    });

    overlay?.addEventListener("dragleave", (e) => {
      e.preventDefault();
      overlay?.classList.remove("active");
    });

    overlay?.addEventListener("drop", (e) => {
      e.preventDefault();
      overlay?.classList.remove("active");
      if (e.dataTransfer.files[0]) {
        this._handleFileUpload(e.dataTransfer.files[0]);
      }
    });

    // Dropzone 自身的拖曳樣式
    dropzone.addEventListener("dragover", () =>
      dropzone.classList.add("dragover"),
    );
    dropzone.addEventListener("dragleave", () =>
      dropzone.classList.remove("dragover"),
    );
    dropzone.addEventListener("drop", () =>
      dropzone.classList.remove("dragover"),
    );
  },

  /**
   * 處理檔案上傳
   */
  async _handleFileUpload(file) {
    this.showLoading(true);

    try {
      if (typeof App !== "undefined") {
        await App.loadFile(file);
        this.isFileLoaded = true;
        this.showToast("success", `成功載入：${file.name}`);

        // 顯示資料表格，隱藏上傳區
        this.elements.dropzone?.classList.add("hidden");
        this.elements.gridContainer?.classList.remove("hidden");
      }
    } catch (err) {
      this.showToast("error", "檔案載入失敗：" + err.message);
    } finally {
      this.showLoading(false);
    }
  },

  /**
   * Toggle 面板顯示/隱藏
   */
  togglePanel(panelId) {
    const panel = document.getElementById(panelId);
    if (panel) {
      panel.classList.toggle("hidden");
    }
  },

  /**
   * 顯示/隱藏模式說明面板
   */
  toggleModeHelp() {
    const panel = this.elements.modeHelpPanel;
    if (panel) {
      panel.classList.toggle("hidden");
      // 高亮當前選擇的模式
      this.elements.modeCards?.forEach((card) => {
        card.classList.toggle("selected", card.dataset.mode === this.selectedMode);
      });
    }
  },

  /**
   * 關閉模式說明面板
   */
  closeModeHelp() {
    this.elements.modeHelpPanel?.classList.add("hidden");
  },

  /**
   * 初始化主題
   */
  _initTheme() {
    const savedTheme = localStorage.getItem("theme") || "light";
    document.documentElement.setAttribute("data-theme", savedTheme);
  },

  /**
   * 切換主題
   */
  toggleTheme() {
    const current = document.documentElement.getAttribute("data-theme");
    const next = current === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("theme", next);
  },

  /**
   * 選擇模式
   */
  selectMode(mode) {
    this.selectedMode = mode;
    this._updateUIForMode(mode);

    // 單選模式需要清除選取
    if (mode === "vertical_group" || mode === "vertical_indent") {
      this.selectedIndices = [];
    }

    // 通知 App
    if (typeof App !== "undefined") {
      App.setMode(mode);
    }
  },

  /**
   * 根據模式更新 UI
   */
  _updateUIForMode(mode) {
    // 更新邏輯設定的可見性
    const isKeywordMode = ["vertical_group", "horizontal_group"].includes(mode);
    const isManualMode = ["horizontal", "vertical_row"].includes(mode);

    // 顯示/隱藏邏輯配置按鈕
    if (this.elements.btnLogicConfig) {
      this.elements.btnLogicConfig.style.display = isKeywordMode ? "" : "none";
    }

    // 加總位置只在縱向關鍵字模式顯示
    if (this.elements.sumDirectionGroup) {
      this.elements.sumDirectionGroup.style.display =
        mode === "vertical_group" ? "" : "none";
    }

    // 操作提示
    if (this.elements.logicHint) {
      if (isManualMode && this.isFileLoaded) {
        this.elements.logicHint.classList.remove("hidden");
        const hintText =
          mode === "horizontal"
            ? "請點選欄位標題以設定驗算邏輯 (第一個點選 = 結果欄)"
            : "請點選列號以設定驗算邏輯 (第一個點選 = 結果列)";
        if (this.elements.logicHintText) {
          this.elements.logicHintText.textContent = hintText;
        }
      } else {
        this.elements.logicHint.classList.add("hidden");
      }
    }
  },

  /**
   * 更新工作表選擇器
   */
  updateSheetSelector(sheetNames, currentSheet) {
    const selector = this.elements.sheetSelector;
    if (!selector) return;

    selector.innerHTML = "";
    sheetNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      option.selected = name === currentSheet;
      selector.appendChild(option);
    });

    if (this.elements.sheetGroup) {
      this.elements.sheetGroup.style.display =
        sheetNames.length > 1 ? "" : "none";
    }
  },

  /**
   * 顯示智能偵測結果
   */
  showSmartDetection(result) {
    const panel = this.elements.smartDetectPanel;
    if (!panel) return;

    if (!result || result.confidence < 30) {
      panel.classList.add("hidden");
      return;
    }

    this.recommendedMode = result.mode;
    const modeInfo = SmartDetect.getModeInfo(result.mode);

    if (this.elements.smartDetectResult) {
      this.elements.smartDetectResult.textContent = `${modeInfo.name} (${result.confidence}% 信心度) - ${result.reasons.join("、")}`;
    }

    panel.classList.remove("hidden");

    // 如果信心度很高，自動套用
    if (result.confidence >= 70) {
      this.selectMode(result.mode);
      this.elements.calcMode.value = result.mode;
      this.showToast("success", `智能推薦已套用：${modeInfo.name}`);
      panel.classList.add("hidden");
    }
  },

  /**
   * 更新錯誤面板
   */
  updateErrorPanel(results) {
    const panel = this.elements.errorPanel;
    if (!panel) return;

    if (!results.hasErrors) {
      panel.classList.add("hidden");
      return;
    }

    panel.classList.remove("hidden");

    if (this.elements.errorCount) {
      this.elements.errorCount.textContent = results.errorCount;
    }

    if (this.elements.errorDiff) {
      const diff = results.totalDiff;
      const prefix = diff >= 0 ? "+" : "";
      this.elements.errorDiff.textContent = prefix + diff.toLocaleString();
    }
  },

  /**
   * 導航到錯誤
   */
  navigateError(direction) {
    const errors = Array.from(Validator.errors.values());
    if (errors.length === 0) return;

    if (direction === "next") {
      this.currentErrorIndex = (this.currentErrorIndex + 1) % errors.length;
    } else {
      this.currentErrorIndex =
        (this.currentErrorIndex - 1 + errors.length) % errors.length;
    }

    const error = errors[this.currentErrorIndex];
    this._scrollToCell(error.row, error.col);
  },

  /**
   * 滾動到指定儲存格
   */
  _scrollToCell(row, col) {
    const cell = document.querySelector(
      `[data-row="${row}"][data-col="${col}"]`,
    );
    if (cell) {
      cell.scrollIntoView({
        behavior: "smooth",
        block: "center",
        inline: "center",
      });
      cell.classList.add("highlight");
      setTimeout(() => cell.classList.remove("highlight"), 2000);
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

    if (typeof App !== "undefined") App.renderGrid();
  },

  /**
   * 切換正負號
   */
  toggleSign(index, event) {
    if (event) event.stopPropagation();

    const current = this.selectedSigns.get(index) || 1;
    this.selectedSigns.set(index, current * -1);

    if (typeof App !== "undefined") App.renderGrid();
  },

  /**
   * 顯示/隱藏載入畫面
   */
  showLoading(show) {
    this.elements.loadingOverlay?.classList.toggle("hidden", !show);
  },

  /**
   * 顯示 Toast 通知
   */
  showToast(type, message, duration = 3000) {
    const container = this.elements.toastContainer;
    if (!container) return;

    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
            <span class="toast-icon">${type === "success" ? "✅" : "❌"}</span>
            <span class="toast-message">${message}</span>
        `;

    container.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = "slideOut 0.3s ease forwards";
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
    this.elements.errorPanel?.classList.add("hidden");
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
        trigger: (this.elements.keyword1?.value || "")
          .split(/[,，]/)
          .map((k) => k.trim())
          .filter((k) => k),
        exclude: (this.elements.keyword2?.value || "")
          .split(/[,，]/)
          .map((k) => k.trim())
          .filter((k) => k),
      },
      sumDirection: this.elements.sumDirection?.value || "top",
      selectedIndices: this.selectedIndices,
      selectedSigns: this.selectedSigns,
    };
  },
};

// 匯出模組
if (typeof module !== "undefined" && module.exports) {
  module.exports = UIController;
}
