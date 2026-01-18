/**
 * 文件上傳組件 - 增強版
 * 支持進度指示、文件信息、多文件拖放
 */

const UploadComponent = {
  elements: {},

  init() {
    this._cacheElements();
    this._bindEvents();
  },

  _cacheElements() {
    this.elements = {
      dropzone: document.getElementById("dropzone"),
      fileInput: document.getElementById("fileInput"),
      dropOverlay: document.getElementById("dropOverlay"),
      progressContainer: document.getElementById("uploadProgressContainer"),
    };
  },

  _bindEvents() {
    const dropzone = this.elements.dropzone;
    const overlay = this.elements.dropOverlay;

    if (!dropzone) return;

    // 點擊觸發文件選擇
    dropzone.addEventListener("click", () => {
      this.elements.fileInput?.click();
    });

    // 文件選擇
    this.elements.fileInput?.addEventListener("change", (e) => {
      this._handleFiles(e.target.files);
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
      if (e.dataTransfer.files.length > 0) {
        this._handleFiles(e.dataTransfer.files);
      }
    });

    // Dropzone 自身的拖曳樣式
    dropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });

    dropzone.addEventListener("dragleave", (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });

    dropzone.addEventListener("drop", (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  },

  _handleFiles(files) {
    if (!files || files.length === 0) return;

    // 處理多個文件（如果需要）
    const file = files[0];

    // 顯示文件信息和進度
    this._showFileProgress(file);

    // 更新 Store
    Store.setState({
      file: file,
      isProcessing: true,
      progress: 0,
    });

    // 處理文件
    if (typeof App !== "undefined") {
      App.loadFile(file);
    }
  },

  _showFileProgress(file) {
    const container = this.elements.progressContainer;
    if (!container) {
      this._createProgressContainer();
      return;
    }

    // 計算文件大小
    const fileSize = this._formatFileSize(file.size);

    container.innerHTML = `
            <div class="upload-progress">
                <div class="upload-progress-info">
                    <div class="upload-file-name">${file.name}</div>
                    <div class="upload-file-size">${fileSize}</div>
                </div>
                <div class="upload-progress-bar">
                    <div class="upload-progress-fill" id="uploadProgressFill"></div>
                </div>
                <div class="upload-progress-text" id="uploadProgressText">準備中...</div>
            </div>
        `;

    // 監聽進度變化
    Store.subscribe("progress", (progress) => {
      this._updateProgress(progress);
    });
  },

  _updateProgress(progress) {
    const fill = document.getElementById("uploadProgressFill");
    const text = document.getElementById("uploadProgressText");

    if (fill) {
      fill.style.width = `${progress}%`;
    }

    if (text) {
      if (progress < 100) {
        text.textContent = `載入中... ${progress}%`;
      } else {
        text.textContent = "完成！";
      }
    }
  },

  _formatFileSize(bytes) {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
  },

  _createProgressContainer() {
    const dropzone = this.elements.dropzone;
    if (!dropzone) return;

    const container = document.createElement("div");
    container.id = "uploadProgressContainer";
    container.className = "upload-progress-container";
    dropzone.appendChild(container);
  },

  hideProgress() {
    const container = this.elements.progressContainer;
    if (container) {
      setTimeout(() => {
        container.innerHTML = "";
      }, 1000);
    }
  },
};
