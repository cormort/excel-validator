/**
 * 錯誤可視化組件
 * 渲染原始Excel表格，高亮顯示錯誤單元格
 */

const ErrorVisualization = {
  elements: {},
  highlightedErrors: new Set(),

  init() {
    this._cacheElements();
    this._bindEvents();
  },

  _cacheElements() {
    this.elements = {
      gridContainer: document.getElementById("gridContainer"),
      errorPanel: document.getElementById("errorPanel"),
      errorList: document.getElementById("errorList"),
      errorCount: document.getElementById("errorCount"),
      btnNextError: document.getElementById("btnNextError"),
      btnPrevError: document.getElementById("btnPrevError"),
    };
  },

  _bindEvents() {
    // 錯誤導航按鈕
    this.elements.btnNextError?.addEventListener("click", () => {
      this.navigateError("next");
    });

    this.elements.btnPrevError?.addEventListener("click", () => {
      this.navigateError("prev");
    });

    // 監聽錯誤狀態變化
    Store.subscribe("errors", (errors) => {
      this._renderErrorList(errors);
    });
  },

  /**
   * 渲染帶高亮錯誤的表格
   */
  renderTable(sheetData, headers, errors) {
    const container = this.elements.gridContainer;
    if (!container) return;

    // 清空容器
    container.innerHTML = "";

    // 創建表格
    const table = document.createElement("table");
    table.className = "data-table data-table-with-errors";

    // 渲染表頭
    if (headers && headers.length > 0) {
      const thead = this._createTableHead(headers, errors);
      table.appendChild(thead);
    }

    // 渲染數據行
    const tbody = document._createTableBody(sheetData, errors);
    table.appendChild(tbody);

    container.appendChild(table);

    // 高亮錯誤單元格
    this._highlightErrors();
  },

  _createTableHead(headers, errors) {
    const thead = document.createElement("thead");
    const tr = document.createElement("tr");

    headers.forEach((header, colIndex) => {
      const th = document.createElement("th");
      th.textContent = header;
      th.className = "table-header";

      // 檢查此列是否有錯誤
      const hasError = errors.some((err) => err.col === colIndex + 1);
      if (hasError) {
        th.classList.add("has-error");
      }

      tr.appendChild(th);
    });

    thead.appendChild(tr);
    return thead;
  },

  _createTableBody(sheetData, errors) {
    const tbody = document.createElement("tbody");

    sheetData.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      tr.className = "table-row";

      row.forEach((cell, colIndex) => {
        const td = document.createElement("td");
        td.textContent = cell !== null && cell !== undefined ? cell : "";
        td.className = "table-cell";

        // 計算實際行列索引（考慮標題行）
        const actualRowIndex = rowIndex + 1;
        const actualColIndex = colIndex + 1;

        // 檢查此單元格是否有錯誤
        const cellError = errors.find(
          (err) => err.row === actualRowIndex && err.col === actualColIndex,
        );

        if (cellError) {
          td.classList.add("error-cell");
          td.dataset.errorId = cellError.id;
          td.dataset.errorType = cellError.type;
        }

        // 檢查此單元格是否相關於錯誤
        const isRelated = errors.some((err) =>
          (err.relatedCells || []).some(
            (related) =>
              related.row === actualRowIndex && related.col === actualColIndex,
          ),
        );

        if (isRelated) {
          td.classList.add("related-cell");
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    return tbody;
  },

  _highlightErrors() {
    const errors = Store.getState("errors");
    const errorCells = document.querySelectorAll(".error-cell");

    // 滾動到第一個錯誤
    if (errorCells.length > 0) {
      const firstError = errorCells[0];
      firstError.scrollIntoView({ behavior: "smooth", block: "center" });

      // 添加動畫效果
      setTimeout(() => {
        firstError.classList.add("error-highlighted");
      }, 300);
    }

    // 為所有錯誤單元格添加點擊事件
    errorCells.forEach((cell) => {
      cell.addEventListener("click", () => {
        this._showErrorDetails(cell);
      });

      cell.addEventListener("mouseenter", () => {
        cell.classList.add("error-hover");
      });

      cell.addEventListener("mouseleave", () => {
        cell.classList.remove("error-hover");
      });
    });
  },

  _renderErrorList(errors) {
    const list = this.elements.errorList;
    if (!list) return;

    // 更新錯誤計數
    if (this.elements.errorCount) {
      this.elements.errorCount.textContent = errors.length;
    }

    // 清空列表
    list.innerHTML = "";

    if (errors.length === 0) {
      list.innerHTML = `
                <div class="error-empty">
                    <div class="error-empty-icon">✅</div>
                    <div class="error-empty-text">未發現錯誤！</div>
                </div>
            `;
      return;
    }

    // 渲染錯誤項目
    errors.forEach((error, index) => {
      const errorItem = document.createElement("div");
      errorItem.className = "error-item";
      errorItem.dataset.errorId = error.id;
      errorItem.dataset.index = index;

      const errorIcons = {
        sum_mismatch: "🔢",
        missing_value: "❓",
        formula_error: "⚠️",
        data_inconsistency: "📊",
      };

      errorItem.innerHTML = `
                <div class="error-item-header">
                    <span class="error-item-icon">${errorIcons[error.type] || "❌"}</span>
                    <span class="error-item-title">錯誤 ${index + 1}</span>
                    <span class="error-item-type">${this._getErrorTypeName(error.type)}</span>
                </div>
                <div class="error-item-location">行 ${error.row}, 列 ${error.col}</div>
                <div class="error-item-diff">
                    <span class="error-diff-label">預期：</span>
                    <span class="error-diff-expected">${error.expected}</span>
                    <span class="error-diff-label">實際：</span>
                    <span class="error-diff-actual">${error.actual}</span>
                </div>
            `;

      errorItem.addEventListener("click", () => {
        this._scrollToError(error);
      });

      list.appendChild(errorItem);
    });
  },

  _getErrorTypeName(type) {
    const typeNames = {
      sum_mismatch: "加總不符",
      missing_value: "遺漏值",
      formula_error: "公式錯誤",
      data_inconsistency: "數據不一致",
    };
    return typeNames[type] || type;
  },

  _scrollToError(error) {
    const errorCell = document.querySelector(`[data-error-id="${error.id}"]`);
    if (errorCell) {
      errorCell.scrollIntoView({ behavior: "smooth", block: "center" });

      // 高亮效果
      setTimeout(() => {
        errorCell.classList.add("error-highlighted");
        setTimeout(() => {
          errorCell.classList.remove("error-highlighted");
        }, 2000);
      }, 300);
    }
  },

  navigateError(direction) {
    const errors = Store.getState("errors");
    const currentItem = document.querySelector(".error-item.active");

    let currentIndex = -1;
    if (currentItem) {
      currentIndex = parseInt(currentItem.dataset.index);
    }

    let nextIndex;
    if (direction === "next") {
      nextIndex = currentIndex < errors.length - 1 ? currentIndex + 1 : 0;
    } else {
      nextIndex = currentIndex > 0 ? currentIndex - 1 : errors.length - 1;
    }

    // 更新選中狀態
    document.querySelectorAll(".error-item").forEach((item) => {
      item.classList.remove("active");
    });

    const nextItem = document.querySelector(
      `.error-item[data-index="${nextIndex}"]`,
    );
    if (nextItem) {
      nextItem.classList.add("active");
      nextItem.scrollIntoView({ behavior: "smooth", block: "nearest" });

      // 滾動到對應的錯誤單元格
      this._scrollToError(errors[nextIndex]);
    }
  },

  _showErrorDetails(cell) {
    const errorId = cell.dataset.errorId;
    const errorType = cell.dataset.errorType;
    const errors = Store.getState("errors");
    const error = errors.find((err) => err.id === errorId);

    if (!error) return;

    // 創建詳情彈窗
    this._showErrorModal(error);
  },

  _showErrorModal(error) {
    // 移除現有彈窗
    const existingModal = document.querySelector(".error-modal");
    if (existingModal) {
      existingModal.remove();
    }

    const modal = document.createElement("div");
    modal.className = "error-modal";
    modal.innerHTML = `
            <div class="error-modal-overlay"></div>
            <div class="error-modal-content">
                <div class="error-modal-header">
                    <h3>錯誤詳情</h3>
                    <button class="error-modal-close">&times;</button>
                </div>
                <div class="error-modal-body">
                    <div class="error-modal-section">
                        <div class="error-modal-label">錯誤類型</div>
                        <div class="error-modal-value error-modal-type">${this._getErrorTypeName(error.type)}</div>
                    </div>
                    <div class="error-modal-section">
                        <div class="error-modal-label">位置</div>
                        <div class="error-modal-value">行 ${error.row}, 列 ${this._getColumnName(error.col)}</div>
                    </div>
                    <div class="error-modal-section">
                        <div class="error-modal-label">預期值</div>
                        <div class="error-modal-value error-modal-expected">${error.expected}</div>
                    </div>
                    <div class="error-modal-section">
                        <div class="error-modal-label">實際值</div>
                        <div class="error-modal-value error-modal-actual">${error.actual}</div>
                    </div>
                    <div class="error-modal-section">
                        <div class="error-modal-label">差異</div>
                        <div class="error-modal-value error-modal-diff">${error.diff}</div>
                    </div>
                    <div class="error-modal-actions">
                        <button class="btn btn-primary error-modal-fix">使用系統值</button>
                        <button class="btn btn-outline error-modal-accept">保持原值</button>
                    </div>
                </div>
            </div>
        `;

    document.body.appendChild(modal);

    // 綁定事件
    modal.querySelector(".error-modal-close").addEventListener("click", () => {
      modal.remove();
    });

    modal
      .querySelector(".error-modal-overlay")
      .addEventListener("click", () => {
        modal.remove();
      });

    modal.querySelector(".error-modal-fix")?.addEventListener("click", () => {
      this._fixError(error);
      modal.remove();
    });

    modal
      .querySelector(".error-modal-accept")
      ?.addEventListener("click", () => {
        modal.remove();
      });

    // 添加動畫
    setTimeout(() => {
      modal.classList.add("visible");
    }, 10);
  },

  _getColumnName(colIndex) {
    // 將列索引轉換為Excel列名（A, B, C...）
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    if (colIndex <= 26) {
      return letters[colIndex - 1] || colIndex;
    } else {
      const first = Math.floor((colIndex - 1) / 26);
      const second = ((colIndex - 1) % 26) + 1;
      return letters[first - 1] + (letters[second - 1] || "");
    }
  },

  _fixError(error) {
    // 觸發修復錯誤的邏輯
    if (typeof UIController !== "undefined") {
      UIController.showToast(
        "success",
        `已修正錯誤：${this._getErrorTypeName(error.type)}`,
      );
    }

    // 通知其他組件
    const event = new CustomEvent("errorFixed", { detail: { error } });
    document.dispatchEvent(event);
  },

  clear() {
    this.highlightedErrors.clear();
    const container = this.elements.gridContainer;
    if (container) {
      container.innerHTML = "";
    }
  },
};
