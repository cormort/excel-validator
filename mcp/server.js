#!/usr/bin/env node
/**
 * SumCheck MCP Server
 * 將 js/validator.js 與 js/smart-detect.js 包成 MCP tools，
 * agent 只需傳檔案路徑（或 TSV 字串），拿回壓縮過的錯誤摘要。
 */

import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import XLSX from 'xlsx';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const Validator = require(path.join(__dirname, '../js/validator.js'));
const SmartDetect = require(path.join(__dirname, '../js/smart-detect.js'));

// 與網頁版 index.html 相同的預設關鍵字
const DEFAULT_TRIGGER = ['主管', '小計', '核定'];
const DEFAULT_EXCLUDE = ['合計', '總計'];
const MODES = ['vertical_group', 'vertical_indent', 'horizontal_group', 'horizontal', 'vertical_row'];

/** 讀取資料來源為 2D 陣列 */
function loadSheet({ file, tsv, sheet }) {
    if (tsv) {
        // 與 paste-app.js 相同的解析：tab 分隔，數字（含千分位）轉 number
        return tsv.replace(/\r/g, '').split('\n').filter((l) => l.length).map((line) =>
            line.split('\t').map((cell) => {
                const trimmed = cell.trim();
                const num = parseFloat(trimmed.replace(/,/g, ''));
                return !isNaN(num) && /^[-\d.,()$%\s]+$/.test(trimmed) ? num : trimmed;
            })
        );
    }
    const wb = XLSX.readFile(file);
    const name = sheet || wb.SheetNames[0];
    const ws = wb.Sheets[name];
    if (!ws) throw new Error(`找不到工作表「${name}」，可用：${wb.SheetNames.join(', ')}`);
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

function colLetter(idx) {
    let s = '';
    for (let n = idx; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(65 + (n % 26)) + s;
    return s;
}

const sourceFields = {
    file: z.string().optional().describe('Excel 檔案絕對路徑（.xlsx/.xls/.ods）'),
    tsv: z.string().optional().describe('TSV 字串（tab 分隔，取代 file，適合驗算 agent 自己產的表格）'),
    sheet: z.string().optional().describe('工作表名稱，預設第一個'),
};

const server = new McpServer({ name: 'sumcheck', version: '0.1.0' });

server.tool(
    'list_sheets',
    '列出 Excel 檔案中的工作表名稱',
    { file: z.string().describe('Excel 檔案絕對路徑') },
    async ({ file }) => ({
        content: [{ type: 'text', text: JSON.stringify(XLSX.readFile(file).SheetNames) }],
    })
);

server.tool(
    'detect_mode',
    '智能偵測表格適合的驗算模式（回傳模式、信心度與理由）',
    { ...sourceFields, header_row: z.number().optional().describe('標題列（1-based，預設 1）') },
    async (args) => {
        const data = loadSheet(args);
        const result = SmartDetect.analyze(data, (args.header_row || 1) - 1);
        const info = result.mode ? SmartDetect.getModeInfo(result.mode) : null;
        return {
            content: [{ type: 'text', text: JSON.stringify({ ...result, modeName: info?.name }) }],
        };
    }
);

server.tool(
    'validate',
    '驗算試算表的合計/小計。不給 mode 時自動偵測。回傳錯誤清單（列/欄、期望值、實際值、差額）與修正建議。',
    {
        ...sourceFields,
        mode: z.enum(MODES).optional()
            .describe('驗算模式；省略則用智能偵測。vertical_group=縱向關鍵字分組(預算表)、vertical_indent=縱向縮排分層、horizontal_group=橫向關鍵字分組、horizontal/vertical_row=手動指定 A±B=C'),
        header_row: z.number().optional().describe('標題列（1-based，預設 1）'),
        end_row: z.number().optional().describe('結束列（1-based，預設最後一列）'),
        start_col: z.number().optional().describe('起始欄（1-based，預設 1）'),
        end_col: z.number().optional().describe('結束欄（1-based，預設最後一欄）'),
        name_col: z.number().optional().describe('名稱欄（1-based，關鍵字/縮排模式用，預設 1）'),
        indices: z.array(z.number()).optional()
            .describe('手動模式（horizontal/vertical_row）的欄或列索引（1-based），最後一個是結果欄/列'),
        signs: z.array(z.number()).optional()
            .describe('與 indices 對應的正負號（1 或 -1），省略全為 1；結果欄的值忽略'),
        trigger_keywords: z.array(z.string()).optional().describe(`小計觸發關鍵字，預設 ${JSON.stringify(DEFAULT_TRIGGER)}`),
        exclude_keywords: z.array(z.string()).optional().describe(`排除關鍵字，預設 ${JSON.stringify(DEFAULT_EXCLUDE)}`),
        sum_direction: z.enum(['top', 'bottom']).optional().describe('縱向關鍵字分組的加總位置：top=關鍵字列在明細上方（預設）、bottom=在下方'),
        max_errors: z.number().optional().describe('回傳錯誤上限，預設 50'),
    },
    async (args) => {
        const data = loadSheet(args);
        const headerRow = args.header_row || 1;

        let mode = args.mode;
        let detected = null;
        if (!mode) {
            detected = SmartDetect.analyze(data, headerRow - 1);
            mode = detected.mode;
            if (!mode) throw new Error(`無法自動判斷驗算模式（${detected.reasons.join('、')}），請指定 mode`);
        }

        const indices = (args.indices || [args.name_col || 1]).map((i) => i - 1);
        if (['horizontal', 'vertical_row'].includes(mode) && indices.length < 2) {
            throw new Error('手動模式需要 indices（至少 2 個，最後一個為結果欄/列）');
        }
        const signs = new Map(indices.map((idx, i) => [idx, (args.signs || [])[i] === -1 ? -1 : 1]));

        const results = Validator.validate({
            mode,
            sheetData: data,
            headerRow,
            endRow: args.end_row || null,
            startCol: args.start_col || 1,
            endCol: args.end_col || null,
            selectedIndices: indices,
            selectedSigns: signs,
            keywords: {
                trigger: args.trigger_keywords || DEFAULT_TRIGGER,
                exclude: args.exclude_keywords || DEFAULT_EXCLUDE,
            },
            sumDirection: args.sum_direction || 'top',
        });

        const max = args.max_errors || 50;
        const errors = Array.from(results.errors.values()).slice(0, max).map((e) => ({
            cell: `${colLetter(e.col)}${e.row + 1}`,
            expected: e.expected,
            actual: e.actual,
            diff: e.diff,
        }));

        return {
            content: [{
                type: 'text',
                text: JSON.stringify({
                    mode,
                    ...(detected ? { autoDetected: true, confidence: detected.confidence } : {}),
                    errorCount: results.errorCount,
                    totalDiff: results.totalDiff,
                    truncated: results.errorCount > max,
                    errors,
                }),
            }],
        };
    }
);

const transport = new StdioServerTransport();
await server.connect(transport);
