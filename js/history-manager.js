/**
 * 歷史記錄組件
 * 保存和顯示驗算歷史
 */

const HistoryManager = {
  HISTORY_KEY: "excel-validator-history",
  MAX_HISTORY: 20,

  elements: {},
  history: [],

  init() {
    this._cacheElements();
    this._loadHistory();
    this._bindEvents();
  },

  _cacheElements() {
    this.elements = {
      historyPanel: document.getElementById("historyPanel"),
      historyList: document.getElementById("historyList"),
      historyCount: document.getElementById("historyCount"),
      btnHistory: document.getElementById("btnHistory"),
      btnClearHistory: document.getElementById("btnClearHistory"),
      btnImportConfig: document.getElementById("btnImportConfig"),
      btnExportConfig: document.getElementById("btnExportConfig"),
    };
  },

  _bindEvents() {
    // 歷史按鈕
    this.elements.btnHistory?.addEventListener("click", () => {
      this._toggleHistoryPanel();
    });

    // 清除歷史
    this.elements.btnClearHistory?.addEventListener("click", () => {
      this._clearHistory();
    });

    // 導入/導出配置
    this.elements.btnImportConfig?.addEventListener("click", () => {
      this._importConfig();
    });

    this.elements.btnExportConfig?.addEventListener("click", () => {
      ConfigManager.exportConfig();
    });

    // 監聽驗算完成事件
    document.addEventListener("validationComplete", (e) => {
      this._addHistoryItem(e.detail);
    });
  },

  _loadHistory() {
    try {
      const historyData = localStorage.getItem(this.HISTORY_KEY);
      if (historyData) {
        this.history = JSON.parse(historyData);
        this._renderHistory();
      }
    } catch (err) {
      console.error("載入歷史失敗:", err);
      this.history = [];
    }
  },

  _renderHistory() {
    const list = this.elements.historyList;
    if (!list) return;

    // 更新計數
    if (this.elements.historyCount) {
      this.elements.historyCount.textContent = this.history.length;
    }

    // 清空列表
    list.innerHTML = "";

    if (this.history.length === 0) {
      list.innerHTML = `
                <div class="history-empty">
                    <div class="history-empty-icon">📋</div>
                    <div class="history-empty-text">尚無歷史記錄</div>
                </div>
            `;
      return;
    }

    // 渲染歷史項目
    this.history.forEach((item, index) => {
      const historyItem = document.createElement("div");
      historyItem.className = "history-item";
      historyItem.dataset.index = index;

      const date = new Date(item.timestamp);
      const dateStr = this._formatDate(date);

      historyItem.innerHTML = `
                <div class="history-item-header">
                    <div class="history-item-mode">${this._getModeName(item.mode)}</div>
                    <div class="history-item-date">${dateStr}</div>
                    <div class="history-item-actions">
                        <button class="history-btn-load" title="重新運行">🔄</button>
                        <button class="history-btn-delete" title="刪除">🗑️</button>
                    </div>
                </div>
                <div class="history-item-details">
                    <div class="history-item-file">📁 ${item.fileName}</div>
                    <div class="history-item-stats">
                        <span class="history-item-errors">錯誤：${item.errorCount}</span>
                        <span class="history-item-duration">耗時：${item.duration}s</span>
                    </div>
                </div>
            `;

      // 綁定事件
      historyItem
        .querySelector(".history-btn-load")
        ?.addEventListener("click", () => {
          this._loadHistoryItem(item);
        });

      historyItem
        .querySelector(".history-btn-delete")
        ?.addEventListener("click", () => {
          this._deleteHistoryItem(index);
        });

      list.appendChild(historyItem);
    });
  },

  _formatDate(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMins < 1) {
      return "剛剛";
    } else if (diffMins < 60) {
      return `${diffMins} 分鐘前`;
    } else if (diffHours < 24) {
      return `${diffHours} 小時前`;
    } else if (diffDays < 7) {
      return `${diffDays} 天前`;
    } else {
      return date.toLocaleDateString("zh-TW", {
        month: "short",
        day: "numeric",
        year: "numeric",
      });
    }
  },

  _getModeName(mode) {
    const modeNames = {
      vertical_group: "縱向 (關鍵字)",
      vertical_indent: "縱向 (縮排)",
      horizontal_group: "橫向 (關鍵字)",
      vertical_manual: "縱向 (手動)",
      horizontal_manual: "橫向 (手動)",
    };
    return modeNames[mode] || mode;
  },

  _addHistoryItem(detail) {
    const item = {
      id: Date.now(),
      timestamp: Date.now(),
      fileName: detail.fileName,
      sheetName: detail.sheetName,
      mode: detail.mode,
      errorCount: detail.errorCount,
      duration: detail.duration,
      config: detail.config,
      results: detail.results,
    };

    // 添加到歷史
    this.history.unshift(item);

    // 限制歷史數量
    if (this.history.length > this.MAX_HISTORY) {
      this.history = this.history.slice(0, this.MAX_HISTORY);
    }

    // 保存到 localStorage
    this._saveHistory();
    this._renderHistory();
  },

  _loadHistoryItem(item) {
    // 加載配置
    if (item.config) {
      this._applyConfig(item.config);
    }

    // 通知其他組件
    const event = new CustomEvent("historyLoaded", { detail: { item } });
    document.dispatchEvent(event);

    UIController.showToast("success", "已載入歷史配置");
  },

  _applyConfig(config) {
    // 恢復模式
    if (config.ranges) {
      ConfigManager._applyRanges(config.ranges);
    }

    if (config.keywords) {
      ConfigManager._applyKeywords(config.keywords);
    }

    if (config.selectedMode) {
      ConfigManager._selectMode(config.selectedMode);
    }
  },

  _deleteHistoryItem(index) {
    if (confirm("確定要刪除這條歷史記錄嗎？")) {
      this.history.splice(index, 1);
      this._saveHistory();
      this._renderHistory();
      UIController.showToast("success", "已刪除歷史記錄");
    }
  },

  _clearHistory() {
    if (confirm("確定要清除所有歷史記錄嗎？")) {
      this.history = [];
      this._saveHistory();
      this._renderHistory();
      UIController.showToast("success", "已清除所有歷史記錄");
    }
  },

  _saveHistory() {
    try {
      localStorage.setItem(this.HISTORY_KEY, JSON.stringify(this.history));
    } catch (err) {
      console.error("保存歷史失敗:", err);
      // 如果存儲空間不足，刪除最舊的記錄
      if (err.name === "QuotaExceededError") {
        this.history = this.history.slice(0, Math.floor(this.MAX_HISTORY / 2));
        this._saveHistory();
      }
    }
  },

  _toggleHistoryPanel() {
    const panel = this.elements.historyPanel;
    if (!panel) {
      this._createHistoryPanel();
      return;
    }

    panel.classList.toggle("visible");
  },

  _createHistoryPanel() {
    const container = document.querySelector(".main") || document.body;

    const panel = document.createElement("div");
    panel.id = "historyPanel";
    panel.className = "history-panel";
    panel.innerHTML = `
            <div class="history-overlay"></div>
            <div class="history-content">
                <div class="history-header">
                    <h3>📋 歷史記錄</h3>
                    <div class="history-controls">
                        <button id="btnImportConfig" class="btn btn-outline" title="導入配置">
                            📥 導入
                        </button>
                        <button id="btnExportConfig" class="btn btn-outline" title="導出配置">
                            📤 導出
                        </button>
                        <button id="btnClearHistory" class="btn btn-danger" title="清除歷史">
                            🗑️ 清除
                        </button>
                        <button class="history-close">&times;</button>
                    </div>
                </div>
                <div class="history-stats">
                    <span id="historyCount">0</span> 條記錄
                </div>
                <div id="historyList" class="history-list"></div>
            </div>
        `;

    container.appendChild(panel);

    // 綁定事件
    panel.querySelector(".history-close")?.addEventListener("click", () => {
      panel.classList.remove("visible");
    });

    panel.querySelector(".history-overlay")?.addEventListener("click", () => {
      panel.classList.remove("visible");
    });

    // 更新元素快取
    this._cacheElements();

    // 渲染歷史
    this._renderHistory();
  },

  _importConfig() {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";
    input.style.display = "none";

    input.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (file) {
        ConfigManager.importConfig(file)
          .then(() => {
            UIController.showToast("success", "配置導入成功");
          })
          .catch((err) => {
            UIController.showToast("error", "配置導入失敗：" + err.message);
          });
      }
    });

    document.body.appendChild(input);
    input.click();
    setTimeout(() => document.body.removeChild(input), 100);
  },

  getCurrentConfig() {
    return {
      ranges: ConfigManager._getRanges(),
      keywords: ConfigManager._getKeywords(),
      selectedMode: ConfigManager._getSelectedMode(),
    };
  },

  saveCurrentValidation(detail) {
    // 保存配置
    ConfigManager.saveConfig();

    // 添加到歷史
    this._addHistoryItem(detail);
  },

  reset() {
    this.history = [];
    this._saveHistory();
    this._renderHistory();
  },
};
