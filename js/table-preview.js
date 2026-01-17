/**
 * 表格預覽組件
 * 可視化配置，支持縮放、凍結標題行、篩選列
 */

const TablePreview = {
  elements: {},
  state: {
    zoomLevel: 1,
    frozenRows: 1,
    frozenCols: 0,
    selectedCells: new Set(),
    highlightedRange: null,
    filters: new Map(),
  },

  init() {
    this._cacheElements();
    this._bindEvents();
  },

  _cacheElements() {
    this.elements = {
      gridContainer: document.getElementById("gridContainer"),
      toolbar:
        document.getElementById("previewToolbar") || this._createToolbar(),
      zoomIn: document.getElementById("btnZoomIn"),
      zoomOut: document.getElementById("btnZoomOut"),
      zoomReset: document.getElementById("btnZoomReset"),
      freezeToggle: document.getElementById("btnFreeze"),
      filterToggle: document.getElementById("btnFilter"),
    };
  },

  _createToolbar() {
    const container = document.getElementById("gridContainer")?.parentElement;
    if (!container) return;

    const toolbar = document.createElement("div");
    toolbar.id = "previewToolbar";
    toolbar.className = "preview-toolbar";
    toolbar.innerHTML = `
            <div class="toolbar-group">
                <button id="btnZoomOut" class="toolbar-btn" title="縮小">
                    <span>-</span>
                </button>
                <span class="toolbar-separator">|</span>
                <span class="toolbar-label" id="zoomLabel">100%</span>
                <button id="btnZoomIn" class="toolbar-btn" title="放大">
                    <span>+</span>
                </button>
                <button id="btnZoomReset" class="toolbar-btn" title="重置縮放">
                    <span>↺</span>
                </button>
            </div>
            <div class="toolbar-group">
                <button id="btnFreeze" class="toolbar-btn" title="凍結標題">
                    <span>🔒</span>
                </button>
                <button id="btnFilter" class="toolbar-btn" title="篩選">
                    <span>🔍</span>
                </button>
            </div>
            <div class="toolbar-info">
                <span id="selectionInfo">未選取</span>
            </div>
        `;

    container.insertBefore(toolbar, container.firstChild);
    return toolbar;
  },

  _bindEvents() {
    // 縮放按鈕
    this.elements.zoomIn?.addEventListener("click", () =>
      this.setZoom(this.state.zoomLevel + 0.25),
    );
    this.elements.zoomOut?.addEventListener("click", () =>
      this.setZoom(this.state.zoomLevel - 0.25),
    );
    this.elements.zoomReset?.addEventListener("click", () => this.setZoom(1));

    // 凍結和篩選切換
    this.elements.freezeToggle?.addEventListener("click", () =>
      this.toggleFreeze(),
    );
    this.elements.filterToggle?.addEventListener("click", () =>
      this.toggleFilter(),
    );

    // 監聽工作表數據變化
    Store.subscribe("sheetData", (sheetData) => {
      if (sheetData) {
        this.renderTable(sheetData);
      }
    });

    // 監聽錯誤變化以高亮範圍
    Store.subscribe("errors", (errors) => {
      this._highlightValidationRange(errors);
    });
  },

  /**
   * 渲染可預覽的表格
   */
  renderTable(sheetData, options = {}) {
    const container = this.elements.gridContainer;
    if (!container) return;

    // 清空容器
    container.innerHTML = "";
    container.style.overflow = "auto";

    // 創建表格
    const table = document.createElement("table");
    table.className = "preview-table";
    table.style.transform = `scale(${this.state.zoomLevel})`;
    table.style.transformOrigin = "top left";

    // 應用篩選
    const filteredData = this._applyFilters(sheetData);

    // 渲染表頭（凍結）
    if (this.state.frozenRows > 0 && filteredData.length > 0) {
      const headers = filteredData[0];
      const thead = this._createFrozenHeaders(headers, filteredData);
      table.appendChild(thead);
    }

    // 渲染數據體
    const tbody = this._createTableBody(filteredData);
    table.appendChild(tbody);

    container.appendChild(table);

    // 設置滾動位置
    this._setupScrollSync();
  },

  _createFrozenHeaders(headers, data) {
    const thead = document.createElement("thead");
    const theadRow = document.createElement("tr");

    headers.forEach((header, index) => {
      const th = document.createElement("th");
      th.textContent = header;
      th.className = "preview-header preview-header-frozen";
      th.style.width = this._calculateColumnWidth(data, index);

      // 可選擇功能
      th.addEventListener("click", () => this._selectColumn(index));
      th.addEventListener("dblclick", () => this._autoFitColumn(index));

      // 范圍選擇
      th.addEventListener("mousedown", (e) => {
        this._startRangeSelection(e, "header", index);
      });

      theadRow.appendChild(th);
    });

    thead.appendChild(theadRow);
    return thead;
  },

  _createTableBody(data) {
    const tbody = document.createElement("tbody");
    const startIndex = this.state.frozenRows > 0 ? 1 : 0;

    data.slice(startIndex).forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      tr.className = "preview-row";

      row.forEach((cell, colIndex) => {
        const td = document.createElement("td");
        td.textContent = cell !== null && cell !== undefined ? cell : "";
        td.className = "preview-cell";

        // 篩選高亮
        if (this._isFilteredCell(rowIndex + startIndex, colIndex)) {
          td.classList.add("cell-filtered");
        }

        // 選中狀態
        if (
          this.state.selectedCells.has(`${rowIndex + startIndex},${colIndex}`)
        ) {
          td.classList.add("cell-selected");
        }

        // 範圍選擇
        td.addEventListener("click", () =>
          this._toggleCellSelection(rowIndex + startIndex, colIndex),
        );
        td.addEventListener("mousedown", (e) => {
          this._startRangeSelection(e, "body", rowIndex + startIndex, colIndex);
        });

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    return tbody;
  },

  _calculateColumnWidth(data, colIndex) {
    // 自動計算列寬
    let maxWidth = 0;
    data.forEach((row) => {
      if (row[colIndex]) {
        const cell = document.createElement("div");
        cell.style.visibility = "hidden";
        cell.style.whiteSpace = "nowrap";
        cell.textContent = row[colIndex];
        document.body.appendChild(cell);
        const width = cell.offsetWidth;
        document.body.removeChild(cell);
        maxWidth = Math.max(maxWidth, width);
      }
    });

    return `${Math.min(maxWidth + 20, 300)}px`;
  },

  _autoFitColumn(colIndex) {
    const cells = document.querySelectorAll(
      `td:nth-child(${colIndex + 1}), th:nth-child(${colIndex + 1})`,
    );
    let maxWidth = 0;

    cells.forEach((cell) => {
      maxWidth = Math.max(maxWidth, cell.offsetWidth);
    });

    cells.forEach((cell) => {
      cell.style.minWidth = `${maxWidth + 10}px`;
      cell.style.maxWidth = `${maxWidth + 10}px`;
    });
  },

  /**
   * 縮放控制
   */
  setZoom(level) {
    // 限制縮放範圍
    this.state.zoomLevel = Math.max(0.5, Math.min(3, level));

    // 更新顯示
    const table = document.querySelector(".preview-table");
    if (table) {
      table.style.transform = `scale(${this.state.zoomLevel})`;
    }

    // 更新標籤
    if (this.elements.zoomLabel) {
      this.elements.zoomLabel.textContent = `${Math.round(this.state.zoomLevel * 100)}%`;
    }

    // 更新容器寬度
    const container = this.elements.gridContainer;
    if (container) {
      container.style.width = `${100 * this.state.zoomLevel}%`;
    }
  },

  toggleFreeze() {
    this.state.frozenRows = this.state.frozenRows > 0 ? 0 : 1;

    const btn = this.elements.freezeToggle;
    if (btn) {
      btn.classList.toggle("active", this.state.frozenRows > 0);
    }

    // 重新渲染表格
    const sheetData = Store.getState("sheetData");
    if (sheetData) {
      this.renderTable(sheetData);
    }
  },

  toggleFilter() {
    const container = document.getElementById("filterPanel");
    if (container) {
      container.classList.toggle("visible");
    } else {
      this._createFilterPanel();
    }
  },

  _createFilterPanel() {
    const existingPanel = document.getElementById("filterPanel");
    if (existingPanel) {
      existingPanel.remove();
    }

    const panel = document.createElement("div");
    panel.id = "filterPanel";
    panel.className = "filter-panel";
    panel.innerHTML = `
            <div class="filter-header">
                <h3>篩選列</h3>
                <button class="filter-close">&times;</button>
            </div>
            <div class="filter-body">
                <input type="text" id="filterInput" placeholder="輸入篩選條件..." />
                <div class="filter-options">
                    <label>
                        <input type="radio" name="filterMode" value="contains" checked />
                        包含
                    </label>
                    <label>
                        <input type="radio" name="filterMode" value="equals" />
                        等於
                    </label>
                    <label>
                        <input type="radio" name="filterMode" value="starts" />
                        開頭是
                    </label>
                </div>
                <div class="filter-actions">
                    <button class="btn btn-primary" id="btnApplyFilter">套用</button>
                    <button class="btn btn-outline" id="btnClearFilter">清除</button>
                </div>
            </div>
        `;

    document.body.appendChild(panel);

    // 綁定事件
    panel.querySelector(".filter-close").addEventListener("click", () => {
      panel.classList.remove("visible");
    });

    panel.querySelector("#btnApplyFilter").addEventListener("click", () => {
      this._applyFilter();
      panel.classList.remove("visible");
    });

    panel.querySelector("#btnClearFilter").addEventListener("click", () => {
      this._clearFilters();
      panel.classList.remove("visible");
    });
  },

  _applyFilters(data) {
    if (this.state.filters.size === 0) return data;

    // 根據篩選條件過濾數據
    return data.filter((row) => {
      return Array.from(this.state.filters.entries()).every(
        ([colIndex, filter]) => {
          const cellValue = String(row[colIndex] || "");

          switch (filter.mode) {
            case "contains":
              return cellValue.includes(filter.value);
            case "equals":
              return cellValue === filter.value;
            case "starts":
              return cellValue.startsWith(filter.value);
            default:
              return true;
          }
        },
      );
    });
  },

  _isFilteredCell(rowIndex, colIndex) {
    const filter = this.state.filters.get(colIndex);
    if (!filter) return false;

    const sheetData = Store.getState("sheetData");
    if (!sheetData) return false;

    const cellValue = String(sheetData[rowIndex]?.[colIndex] || "");

    switch (filter.mode) {
      case "contains":
        return !cellValue.includes(filter.value);
      case "equals":
        return cellValue !== filter.value;
      case "starts":
        return !cellValue.startsWith(filter.value);
      default:
        return false;
    }
  },

  _applyFilter() {
    const input = document.getElementById("filterInput");
    const mode = document.querySelector('input[name="filterMode"]:checked');
    const colSelect = document.getElementById("filterColumnSelect");

    if (input && mode) {
      const colIndex = colSelect ? parseInt(colSelect.value) : 0;
      this.state.filters.set(colIndex, {
        value: input.value,
        mode: mode.value,
      });

      const sheetData = Store.getState("sheetData");
      if (sheetData) {
        this.renderTable(sheetData);
      }
    }
  },

  _clearFilters() {
    this.state.filters.clear();
    const sheetData = Store.getState("sheetData");
    if (sheetData) {
      this.renderTable(sheetData);
    }
  },

  /**
   * 單元格選擇和範圍選擇
   */
  _toggleCellSelection(rowIndex, colIndex) {
    const key = `${rowIndex},${colIndex}`;

    if (this.state.selectedCells.has(key)) {
      this.state.selectedCells.delete(key);
    } else {
      this.state.selectedCells.clear();
      this.state.selectedCells.add(key);
    }

    this._updateCellStyles();
    this._updateSelectionInfo();
  },

  _startRangeSelection(e, type, row, col) {
    // 簡化版本，實際應該支持拖動選擇
    this._toggleCellSelection(row, col);

    // 更新 UI 以顯示選擇開始
    Store.setState({
      selectionStart: { type, row, col },
    });
  },

  _updateCellStyles() {
    const cells = document.querySelectorAll(".preview-cell");
    cells.forEach((cell) => {
      cell.classList.remove(
        "cell-selected",
        "cell-range-start",
        "cell-range-end",
      );
    });

    this.state.selectedCells.forEach((key) => {
      const [row, col] = key.split(",").map(Number);
      const cell = document.querySelector(
        `.preview-row:nth-child(${row + 1}) .preview-cell:nth-child(${col + 1})`,
      );
      if (cell) {
        cell.classList.add("cell-selected");
      }
    });
  },

  _updateSelectionInfo() {
    const info = document.getElementById("selectionInfo");
    if (info) {
      const count = this.state.selectedCells.size;
      if (count === 0) {
        info.textContent = "未選取";
      } else if (count === 1) {
        const [row, col] = Array.from(this.state.selectedCells)[0].split(",");
        info.textContent = `已選取: ${this._getColumnName(parseInt(col) + 1)}${row + 1}`;
      } else {
        info.textContent = `已選取 ${count} 個單元格`;
      }
    }
  },

  _selectColumn(colIndex) {
    // 選擇整列
    const sheetData = Store.getState("sheetData");
    if (!sheetData) return;

    this.state.selectedCells.clear();
    sheetData.forEach((row, rowIndex) => {
      this.state.selectedCells.add(`${rowIndex},${colIndex}`);
    });

    this._updateCellStyles();
    this._updateSelectionInfo();
  },

  _getColumnName(colIndex) {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    if (colIndex <= 26) {
      return letters[colIndex - 1] || colIndex;
    } else {
      const first = Math.floor((colIndex - 1) / 26);
      const second = ((colIndex - 1) % 26) + 1;
      return letters[first - 1] + (letters[second - 1] || "");
    }
  },

  /**
   * 滾動同步（凍結標題時）
   */
  _setupScrollSync() {
    const container = this.elements.gridContainer;
    if (!container || this.state.frozenRows === 0) return;

    const thead = document.querySelector(".preview-header-frozen");
    if (!thead) return;

    container.addEventListener("scroll", () => {
      thead.style.transform = `translateX(${-container.scrollLeft}px)`;
    });
  },

  _highlightValidationRange(errors) {
    // 高亮顯示錯誤的單元格
    document.querySelectorAll(".preview-cell").forEach((cell) => {
      cell.classList.remove("cell-error", "cell-warning");
    });

    errors.forEach((error) => {
      const cell = document.querySelector(
        `.preview-row:nth-child(${error.row}) .preview-cell:nth-child(${error.col})`,
      );
      if (cell) {
        cell.classList.add("cell-error");
      }
    });
  },

  reset() {
    this.state = {
      zoomLevel: 1,
      frozenRows: 1,
      frozenCols: 0,
      selectedCells: new Set(),
      highlightedRange: null,
      filters: new Map(),
    };

    const container = this.elements.gridContainer;
    if (container) {
      container.innerHTML = "";
    }

    if (this.elements.zoomLabel) {
      this.elements.zoomLabel.textContent = "100%";
    }
  },
};
