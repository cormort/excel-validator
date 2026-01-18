/**
 * 配置管理組件
 * 自動保存和載入配置到 localStorage
 */

const ConfigManager = {
  CONFIG_KEY: "excel-validator-config",

  elements: {},

  init() {
    this._cacheElements();
    this._loadConfig();
  },

  _cacheElements() {
    this.elements = {
      // 設定相關
      headerRow: document.getElementById("headerRow"),
      endRow: document.getElementById("endRow"),
      startCol: document.getElementById("startCol"),
      endCol: document.getElementById("endCol"),

      // 關鍵字設定
      keyword1: document.getElementById("keyword1"),
      keyword2: document.getElementById("keyword2"),
      sumDirection: document.getElementById("sumDirection"),

      // 模式和範圍
      selectedMode: null,
      modeCards: document.querySelectorAll(".mode-card"),
    };
  },

  _loadConfig() {
    try {
      const config = localStorage.getItem(this.CONFIG_KEY);
      if (config) {
        const parsed = JSON.parse(config);
        this._applyConfig(parsed);
      }
    } catch (err) {
      console.error("載入配置失敗:", err);
    }
  },

  _applyConfig(config) {
    // 恢復上傳的檔案名稱（如果需要）
    if (config.fileName) {
      const fileDisplay = document.getElementById("uploadedFileName");
      if (fileDisplay) {
        fileDisplay.textContent = `上次檔案：${config.fileName}`;
      }
    }

    // 恢復工作表選擇
    if (config.sheetName) {
      const sheetSelector = document.getElementById("sheetSelector");
      if (sheetSelector) {
        const option = Array.from(sheetSelector.options).find(
          (opt) => opt.value === config.sheetName,
        );
        if (option) {
          sheetSelector.value = config.sheetName;
        }
      }
    }

    // 恢復範圍設定
    if (config.ranges) {
      this._applyRanges(config.ranges);
    }

    // 恢復關鍵字設定
    if (config.keywords) {
      this._applyKeywords(config.keywords);
    }

    // 恢復選擇的模式
    if (config.selectedMode) {
      this._selectMode(config.selectedMode);
    }
  },

  _applyRanges(ranges) {
    if (ranges.headerRow && this.elements.headerRow) {
      this.elements.headerRow.value = ranges.headerRow;
    }
    if (ranges.endRow && this.elements.endRow) {
      this.elements.endRow.value = ranges.endRow;
    }
    if (ranges.startCol && this.elements.startCol) {
      this.elements.startCol.value = ranges.startCol;
    }
    if (ranges.endCol && this.elements.endCol) {
      this.elements.endCol.value = ranges.endCol;
    }
  },

  _applyKeywords(keywords) {
    if (keywords.trigger && this.elements.keyword1) {
      this.elements.keyword1.value = keywords.trigger;
    }
    if (keywords.exclude && this.elements.keyword2) {
      this.elements.keyword2.value = keywords.exclude;
    }
    if (keywords.direction && this.elements.sumDirection) {
      this.elements.sumDirection.value = keywords.direction;
    }
  },

  _selectMode(mode) {
    this.elements.modeCards.forEach((card) => {
      card.classList.remove("selected");
      if (card.dataset.mode === mode) {
        card.classList.add("selected");
        card.click();
      }
    });
  },

  saveConfig() {
    const config = {
      fileName: Store.getState("file")?.name,
      sheetName: Store.getState("currentSheet"),
      ranges: this._getRanges(),
      keywords: this._getKeywords(),
      selectedMode: this._getSelectedMode(),
      timestamp: Date.now(),
    };

    try {
      localStorage.setItem(this.CONFIG_KEY, JSON.stringify(config));
      console.log("配置已保存:", config);
    } catch (err) {
      console.error("保存配置失敗:", err);
    }

    return config;
  },

  _getRanges() {
    return {
      headerRow: this.elements.headerRow?.value,
      endRow: this.elements.endRow?.value,
      startCol: this.elements.startCol?.value,
      endCol: this.elements.endCol?.value,
    };
  },

  _getKeywords() {
    return {
      trigger: this.elements.keyword1?.value,
      exclude: this.elements.keyword2?.value,
      direction: this.elements.sumDirection?.value,
    };
  },

  _getSelectedMode() {
    const selectedCard = document.querySelector(".mode-card.selected");
    return selectedCard?.dataset.mode || null;
  },

  exportConfig() {
    const config = this.saveConfig();
    const blob = new Blob([JSON.stringify(config, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "excel-validator-config.json";
    a.click();
    URL.revokeObjectURL(url);
  },

  importConfig(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const config = JSON.parse(e.target.result);
          this._applyConfig(config);
          resolve(config);
        } catch (err) {
          reject(new Error("無效的配置檔"));
        }
      };
      reader.onerror = () => reject(new Error("讀取配置檔失敗"));
      reader.readAsText(file);
    });
  },

  clearConfig() {
    localStorage.removeItem(this.CONFIG_KEY);
    console.log("配置已清除");

    // 清除表單
    if (this.elements.headerRow) this.elements.headerRow.value = "";
    if (this.elements.endRow) this.elements.endRow.value = "";
    if (this.elements.startCol) this.elements.startCol.value = "";
    if (this.elements.endCol) this.elements.endCol.value = "";
    if (this.elements.keyword1) this.elements.keyword1.value = "";
    if (this.elements.keyword2) this.elements.keyword2.value = "";

    // 清除模式選擇
    this.elements.modeCards.forEach((card) => {
      card.classList.remove("selected", "recommended");
    });
  },

  reset() {
    this.clearConfig();
  },
};
