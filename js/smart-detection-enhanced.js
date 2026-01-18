/**
 * 智能偵測和流程簡化組件
 * 自動分析表格結構並推薦最適合的驗算模式
 */

const SmartDetection = {
  elements: {},
  currentRecommendation: null,

  init() {
    this._cacheElements();
  },

  _cacheElements() {
    this.elements = {
      panel: document.getElementById("smartDetectPanel"),
      title: document.querySelector(".smart-detect-title"),
      desc: document.querySelector(".smart-detect-desc"),
      confidenceBar: document.querySelector(".confidence-bar"),
      confidenceFill: document.querySelector(".confidence-fill"),
      quickValidateBtn: document.getElementById("btnQuickValidate"),
    };
  },

  async analyze(sheetData, headers) {
    this._showAnalyzing();

    // 模擬分析過程（實際應該調用 SmartDetect 模組）
    const analysis = await this._analyzeStructure(sheetData, headers);

    this._showRecommendation(analysis);

    return analysis;
  },

  _analyzeStructure(sheetData, headers) {
    return new Promise((resolve) => {
      setTimeout(() => {
        const analysis = this._detectPattern(sheetData, headers);
        resolve(analysis);
      }, 1500); // 模擬分析時間
    });
  },

  _detectPattern(sheetData, headers) {
    // 簡單的啟發式算法來檢測表格模式

    // 1. 檢查是否有明顯的分組關鍵字（如：總計、小計）
    const hasSumKeywords = this._hasSumKeywords(sheetData);

    // 2. 檢查是否有縮排結構
    const hasIndentation = this._hasIndentation(sheetData);

    // 3. 檢查數據分佈
    const isHorizontal = this._isHorizontalLayout(sheetData);

    // 4. 計算信心度
    let recommendation = null;
    let confidence = 0;

    if (hasSumKeywords && isHorizontal) {
      recommendation = "horizontal_group";
      confidence = 0.85;
    } else if (hasSumKeywords) {
      recommendation = hasIndentation ? "vertical_indent" : "vertical_group";
      confidence = 0.8;
    } else if (isHorizontal) {
      recommendation = "horizontal_group";
      confidence = 0.7;
    }

    return {
      recommendedMode: recommendation,
      confidence: confidence,
      reasoning: this._getReasoning(
        hasSumKeywords,
        hasIndentation,
        isHorizontal,
      ),
      features: {
        hasSumKeywords,
        hasIndentation,
        isHorizontal,
      },
    };
  },

  _hasSumKeywords(sheetData) {
    const keywords = ["總計", "小計", "合計", "Total", "Subtotal"];
    return sheetData.some((row) => {
      return row.some((cell) => {
        const value = String(cell).trim();
        return keywords.some((keyword) => value.includes(keyword));
      });
    });
  },

  _hasIndentation(sheetData) {
    // 檢查是否第一列有明顯的縮排結構（如空格）
    const firstColValues = sheetData.map((row) => row[0] || "").slice(0, 20);
    const indentedCount = firstColValues.filter((value) => {
      return (value && value.startsWith("　")) || value.startsWith("  ");
    }).length;

    return indentedCount > firstColValues.length * 0.3; // 30%以上有縮排
  },

  _isHorizontalLayout(sheetData) {
    // 如果列數遠大於行數，可能是橫向佈局
    const rows = sheetData.length;
    const cols = Math.max(...sheetData.map((row) => row.length));
    return cols > rows * 1.5;
  },

  _getReasoning(hasSum, hasIndent, isHorizontal) {
    const reasons = [];

    if (hasSum) {
      reasons.push("檢測到合計關鍵字");
    }
    if (hasIndent) {
      reasons.push("發現階層縮排結構");
    }
    if (isHorizontal) {
      reasons.push("數據傾向於橫向排列");
    }

    return reasons.join("、");
  },

  _showAnalyzing() {
    if (!this.elements.panel) return;

    this.elements.panel.classList.remove("hidden");
    this.elements.title.textContent = "分析中...";
    this.elements.desc.textContent = "正在分析表格結構";
    this.elements.confidenceFill.style.width = "0%";
  },

  _showRecommendation(analysis) {
    if (!this.elements.panel) return;

    const { recommendedMode, confidence, reasoning } = analysis;

    if (recommendedMode) {
      const modeNames = {
        vertical_group: "縱向 (關鍵字分組)",
        vertical_indent: "縱向 (縮排分層)",
        horizontal_group: "橫向 (關鍵字分組)",
        vertical_manual: "縱向 (手動指定)",
        horizontal_manual: "橫向 (手動指定)",
      };

      this.elements.title.textContent = "推薦模式";
      this.elements.desc.textContent =
        modeNames[recommendedMode] || "建議自定義";
      this.elements.confidenceFill.style.width = `${confidence * 100}%`;

      // 高亮推薦的模式卡片
      this._highlightRecommendedMode(recommendedMode);

      // 更新 Store
      Store.setState({
        smartRecommendation: analysis,
      });

      // 顯示一鍵驗算按鈕
      this._showQuickValidateButton();
    } else {
      this.elements.title.textContent = "無法自動判斷";
      this.elements.desc.textContent = "請手動選擇驗算模式";
      this.elements.panel.classList.remove("success");
    }
  },

  _highlightRecommendedMode(mode) {
    const cards = document.querySelectorAll(".mode-card");
    cards.forEach((card) => {
      card.classList.remove("recommended");
      if (card.dataset.mode === mode) {
        card.classList.add("recommended");
      }
    });
  },

  _showQuickValidateButton() {
    // 檢查是否已有一鍵驗算按鈕，如果沒有則創建
    if (!this.elements.quickValidateBtn) {
      const btn = document.createElement("button");
      btn.id = "btnQuickValidate";
      btn.className = "btn btn-success btn-quick-validate";
      btn.innerHTML = "✨ 一鍵驗算";
      btn.addEventListener("click", () => {
        if (typeof App !== "undefined") {
          App.runValidation();
        }
      });

      // 插入到適當位置
      const wizardActions = document.querySelector(".wizard-actions");
      if (wizardActions) {
        wizardActions.insertBefore(btn, wizardActions.firstChild);
      }

      this.elements.quickValidateBtn = btn;
    }

    this.elements.quickValidateBtn?.classList.remove("hidden");
  },

  reset() {
    this.currentRecommendation = null;
    const cards = document.querySelectorAll(".mode-card");
    cards.forEach((card) => card.classList.remove("recommended"));

    if (this.elements.quickValidateBtn) {
      this.elements.quickValidateBtn.classList.add("hidden");
    }
  },
};
