/**
 * Excel 驗算大師 - Excel 解析模組
 * 處理 Excel 檔案載入與資料轉換
 */

const ExcelParser = {
    workbook: null,
    currentSheet: null,
    sheetData: [],
    headers: [],

    /**
     * 從 File 物件讀取 Excel
     * @param {File} file - Excel 檔案
     * @returns {Promise<Object>} 解析結果
     */
    async parseFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    this.workbook = XLSX.read(data, { type: 'array' });

                    const sheetNames = this.workbook.SheetNames;

                    // 自動載入第一個工作表
                    if (sheetNames.length > 0) {
                        this.loadSheet(sheetNames[0]);
                    }

                    resolve({
                        success: true,
                        sheetNames,
                        currentSheet: sheetNames[0],
                        rowCount: this.sheetData.length,
                        colCount: this.sheetData[0]?.length || 0,
                    });
                } catch (err) {
                    reject(new Error('Excel 解析失敗：' + err.message));
                }
            };

            reader.onerror = () => reject(new Error('檔案讀取失敗'));
            reader.readAsArrayBuffer(file);
        });
    },

    /**
     * 切換工作表
     * @param {string} sheetName - 工作表名稱
     */
    loadSheet(sheetName) {
        if (!this.workbook || !this.workbook.Sheets[sheetName]) {
            throw new Error('工作表不存在：' + sheetName);
        }

        this.currentSheet = sheetName;
        const ws = this.workbook.Sheets[sheetName];
        this.sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        return {
            sheetName,
            rowCount: this.sheetData.length,
            colCount: this.sheetData.reduce((max, row) => Math.max(max, row?.length || 0), 0),
        };
    },

    /**
     * 取得目前工作表資料
     */
    getData() {
        return this.sheetData;
    },

    /**
     * 取得所有工作表名稱
     */
    getSheetNames() {
        return this.workbook?.SheetNames || [];
    },

    /**
     * 取得指定範圍的資料
     * @param {Object} range - { headerRow, endRow, startCol, endCol }
     */
    getDataRange(range) {
        const { headerRow = 1, endRow, startCol = 1, endCol } = range;

        const hIdx = headerRow - 1;
        const sCol = startCol - 1;
        const eRow = endRow || this.sheetData.length;
        const eCol = endCol || (this.sheetData[hIdx]?.length || 0);

        this.headers = this.sheetData[hIdx]?.slice(sCol, eCol) || [];

        return {
            headers: this.headers,
            data: this.sheetData.slice(hIdx + 1, eRow).map(row =>
                (row || []).slice(sCol, eCol)
            ),
            headerRowIndex: hIdx,
            startColIndex: sCol,
            endRowIndex: eRow - 1,
            endColIndex: eCol - 1,
        };
    },

    /**
     * 自動偵測資料範圍
     */
    detectDataRange() {
        if (!this.sheetData.length) return null;

        let headerRow = 1;
        let startCol = 1;
        let endRow = this.sheetData.length;
        let endCol = 0;

        // 找最大欄數
        this.sheetData.forEach(row => {
            if (row?.length > endCol) endCol = row.length;
        });

        // 嘗試找標題列（假設前幾列可能是說明文字）
        for (let r = 0; r < Math.min(10, this.sheetData.length); r++) {
            const row = this.sheetData[r];
            if (!row) continue;

            // 如果這列有多個非空儲存格，可能是標題列
            const nonEmptyCells = row.filter(cell => cell !== null && cell !== undefined && cell !== '').length;
            if (nonEmptyCells >= 3) {
                headerRow = r + 1;
                break;
            }
        }

        return { headerRow, endRow, startCol, endCol };
    },

    /**
     * 數值清理與轉換
     * @param {*} val - 原始值
     * @returns {number} 數值或 NaN
     */
    parseNumber(val) {
        if (val === null || val === undefined || val === '') return NaN;

        // 已是數字
        if (typeof val === 'number') return isNaN(val) ? NaN : val;

        const str = val.toString().trim();
        if (str === '') return NaN;

        // 排除含中英文字母的儲存格
        if (/[a-zA-Z\u4e00-\u9fa5]/.test(str)) return NaN;

        // 白名單檢查
        const validPattern = /^[-\d\.,，\(\)\$%\uFF05\s]+$/;
        if (!validPattern.test(str)) return NaN;

        // 清理並轉換
        let cleanStr = str
            .replace(/[,，\$\s]/g, '')
            .replace(/[%％]/g, '');

        // 處理會計負數格式 (123) -> -123
        if (cleanStr.startsWith('(') && cleanStr.endsWith(')')) {
            cleanStr = '-' + cleanStr.replace(/[\(\)]/g, '');
        }

        const num = parseFloat(cleanStr);
        return isNaN(num) ? NaN : num;
    },

    /**
     * 格式化數字顯示
     */
    formatNumber(val) {
        if (typeof val === 'number') {
            return val.toLocaleString();
        }
        return val ?? '';
    },

    /**
     * 清除資料
     */
    reset() {
        this.workbook = null;
        this.currentSheet = null;
        this.sheetData = [];
        this.headers = [];
    },
};

// 匯出模組
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelParser;
}
