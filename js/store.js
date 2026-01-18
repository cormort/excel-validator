/**
 * 簡易狀態管理器
 * 使用響應式模式自動更新 UI
 */

const Store = {
  state: {
    // 檔案相關
    file: null,
    workbook: null,
    currentSheet: null,
    sheetData: [],
    headers: [],

    // 設定相關
    selectedMode: null,
    step: 1,
    smartRecommendation: null,

    // 驗算結果
    errors: [],
    corrections: [],
    isProcessing: false,
    progress: 0,

    // 主題
    theme: "light",
  },

  listeners: new Map(),

  subscribe(key, callback) {
    if (!this.listeners.has(key)) {
      this.listeners.set(key, new Set());
    }
    this.listeners.get(key).add(callback);

    return () => this.unsubscribe(key, callback);
  },

  unsubscribe(key, callback) {
    if (this.listeners.has(key)) {
      this.listeners.get(key).delete(callback);
    }
  },

  setState(updates) {
    const oldState = { ...this.state };
    this.state = { ...this.state, ...updates };

    // 通知監聽器
    Object.keys(updates).forEach((key) => {
      if (this.listeners.has(key)) {
        this.listeners.get(key).forEach((callback) => {
          callback(this.state[key], oldState[key], this.state);
        });
      }
    });
  },

  getState(key) {
    return key ? this.state[key] : { ...this.state };
  },

  reset() {
    this.state = {
      file: null,
      workbook: null,
      currentSheet: null,
      sheetData: [],
      headers: [],
      selectedMode: null,
      step: 1,
      smartRecommendation: null,
      errors: [],
      corrections: [],
      isProcessing: false,
      progress: 0,
      theme: "light",
    };

    this.listeners.forEach((callbacks, key) => {
      callbacks.forEach((callback) => {
        callback(this.state[key], this.state[key], this.state);
      });
    });
  },
};

// CommonJS module export
if (typeof module !== "undefined" && module.exports) {
  module.exports = Store;
}
