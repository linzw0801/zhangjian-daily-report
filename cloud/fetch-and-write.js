/**
 * 杖剑 港台/欧美 日报 云端 OpenAPI 版（不依赖 lark-cli / 不依赖本机 keychain）
 *
 * 认证：app_id + app_secret 换 tenant_access_token（应用身份，需应用已加为表协作者）。
 *       token 每 2 小时自动换新，永不过期，无需用户重复授权。
 * 飞书读写：range 用 `sheetId!A1:B2` 格式（单格也写成 sheetId!J106:J106）。
 * 凭据走环境变量：FEISHU_APP_ID / FEISHU_APP_SECRET / QQ_MAIL_PASSWORD / REGION(gt|mw)
 *
 * 数据流：IMAP 取看板 ZIP → 解析 base/ios/google/留存/LTV CSV → 写飞书三 sheet。
 *   先写日报表（新增/流水/活跃，最重要），再补留存/LTV 缺失列（读现状→对比→只补缺失）。
 */
'use strict';

const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

const REGION = (process.env.REGION || 'mw').trim();
const FEISHU_APP_ID = process.env.FEISHU_APP_ID || '';
const FEISHU_APP_SECRET = process.env.FEISHU_APP_SECRET || '';
const QQ_MAIL_PASSWORD = process.env.QQ_MAIL_PASSWORD || '';

// ==================== 地区配置 ====================
const REGIONS = {
  gt: {
    feishuSpreadsheetToken: 'WIZxszmsCh9CKJtpcYKczPAinKc',
    sheetIds: { ribao: 'ROl670', ios: 'ObmQcB', google: 'Wdmzbw' },
    dashboardSubjectPrefix: '发行数据表自用-港台', dashboardSubjectAltPrefix: null,
    adMailSubject: '杖剑港台日报',
    baseCsvTitle: '发行数据表-基础数据-发行数据表(1)',
    iosCsvTitle: 'iOS基础数据-发行数据表(1)', googleCsvTitle: 'Google基础数据-发行数据表(1)',
    retentionCsvName: '账号留存', ltvCsvName: '账号LTV',
    rlSheets: {
      Total:  { sheetId: 'FMtVhz', channel: '总体', retentionStart: 'K' },
      iOS:    { sheetId: 'ObmQcB', channel: '410001', retentionStart: 'J' },
      Google: { sheetId: 'Wdmzbw', channel: '310001', retentionStart: 'J' },
    },
    rlTotalRow: 3,
    dataTableUrl: 'https://leiting.feishu.cn/wiki/TmLYw7DCViOOZZkUR7uckJJKnsc',
    // 日报表广告列布局（与本地脚本一致）：
    // 港台：AD=gp广告新增, AE-AH=公式列(不写), AI=iOS消耗, AJ=GP消耗
    adCol: { colGpAcct: 'AD', colIosCost: 'AI', colGpCost: 'AJ' },
    ribaoCols: 15, // 港台 C-Q
    label: '港台',
  },
  mw: {
    feishuSpreadsheetToken: 'ThjWsmJmIhS0xntmrnBcWcyLnog',
    sheetIds: { ribao: 'CoTDkN', ios: 'vppnTP', google: 'iMhQNi' },
    dashboardSubjectPrefix: '发行数据表自用-欧美', dashboardSubjectAltPrefix: '发行数据表自用-副本',
    adMailSubject: '杖剑欧美日报',
    baseCsvTitle: '发行数据表-基础数据-发行数据表(2)',
    iosCsvTitle: 'iOS基础数据-发行数据表(2)', googleCsvTitle: 'Google基础数据-发行数据表(2)',
    retentionCsvName: '账号留存', ltvCsvName: '账号LTV',
    rlSheets: {
      Total:  { sheetId: 'ka5WTk', channel: '总体', retentionStart: 'J' },
      iOS:    { sheetId: 'vppnTP', channel: '410001', retentionStart: 'J' },
      Google: { sheetId: 'iMhQNi', channel: '310001', retentionStart: 'J' },
    },
    rlTotalRow: 3,
    dataTableUrl: 'https://leiting.feishu.cn/wiki/FO2mwaVvdiurTsk3b2ackqPFnTd',
    adCol: { colGpAcct: 'V', colIosCost: 'AA', colGpCost: 'AB' },
    ribaoCols: 11, // 欧美 C-M
    label: '欧美',
  },
};

const C = REGIONS[REGION];
if (!C) { console.error('未知 REGION:', REGION); process.exit(1); }
if (!FEISHU_APP_ID || !FEISHU_APP_SECRET) { console.error('缺少 FEISHU_APP_ID / FEISHU_APP_SECRET'); process.exit(1); }
if (!QQ_MAIL_PASSWORD) { console.error('缺少 QQ_MAIL_PASSWORD'); process.exit(1); }

const CONFIG = {
  feishuSpreadsheetToken: C.feishuSpreadsheetToken,
  sheetIds: C.sheetIds,
  imap: { user: '3059402@qq.com', password: QQ_MAIL_PASSWORD, host: 'imap.qq.com', port: 993 },
  dashboardMailFrom: 'notice@email.thinkingdata.cn',
  dashboardSubjectPrefix: C.dashboardSubjectPrefix, dashboardSubjectAltPrefix: C.dashboardSubjectAltPrefix,
  adMailFrom: 'ads@g-bits.com', adMailSubject: C.adMailSubject,
  baseCsvTitle: C.baseCsvTitle, iosCsvTitle: C.iosCsvTitle, googleCsvTitle: C.googleCsvTitle,
  retentionCsvName: C.retentionCsvName, ltvCsvName: C.ltvCsvName,
  rlSheets: C.rlSheets, rlTotalRow: C.rlTotalRow, dataTableUrl: C.dataTableUrl,
  adCol: C.adCol, ribaoCols: C.ribaoCols, label: C.label,
};

function log(...a) { console.log(`[${new Date().toISOString()}]`, ...a); }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ==================== 飞书 OpenAPI（应用身份） ====================
let _tenantToken = null; let _tenantExp = 0;
async function getTenantAccessToken() {
  if (_tenantToken && Date.now() < _tenantExp - 60000) return _tenantToken;
  const r = await httpsReq('POST', 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
    { 'Content-Type': 'application/json' }, JSON.stringify({ app_id: FEISHU_APP_ID, app_secret: FEISHU_APP_SECRET }));
  if (!r || r.code !== 0) throw new Error('获取 tenant_access_token 失败: ' + JSON.stringify(r));
  _tenantToken = r.tenant_access_token; _tenantExp = Date.now() + (r.expire || 7200) * 1000;
  return _tenantToken;
}
function httpsReq(method, url, headers, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, { method, headers }, res => {
      let d = ''; res.on('data', c => d += c); res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { resolve({ raw: d, status: res.statusCode }); } });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}
async function feishu(method, apiPath, bodyObj) {
  const token = await getTenantAccessToken();
  const headers = { 'Authorization': 'Bearer ' + token };
  if (bodyObj !== undefined) headers['Content-Type'] = 'application/json';
  for (let i = 1; i <= 3; i++) {
    try {
      const r = await httpsReq(method, 'https://open.feishu.cn' + apiPath, headers, bodyObj !== undefined ? JSON.stringify(bodyObj) : undefined);
      if (r && /900015205|cs recommited|cannot.*revision/i.test(JSON.stringify(r))) { if (i < 3) { log('revision 冲突重试', i, apiPath); await sleep(i * 2000); continue; } }
      return r;
    } catch (e) { if (i < 3) { log('网络重试', i, e.message); await sleep(i * 2000); continue; } throw e; }
  }
}
// 读 range: sheetId!A1:B2 → { rows:[[{value}]], col_indices, row_indices }
async function readRange(sheetId, a1Range) {
  const full = `${sheetId}!${a1Range}`;
  const r = await feishu('GET', `/open-apis/sheets/v2/spreadsheets/${CONFIG.feishuSpreadsheetToken}/values/${encodeURIComponent(full)}`);
  if (!r || r.code !== 0) throw new Error('读失败 ' + full + ': ' + JSON.stringify(r));
  const vals = ((r.data || {}).valueRange || {}).values || [];
  const colIdx = colSeqBetween(splitRange(a1Range));
  const rowIdx = rowNums(splitRange(a1Range));
  const rows = vals.map(row => row.map(v => ({ value: v === undefined || v === null ? null : v })));
  return { rows, col_indices: colIdx, row_indices: rowIdx };
}
// 写 range: sheetId!A1:B2, cells=[[{value}...],...]（值静态覆盖）
async function writeRange(sheetId, a1Range, cells2D) {
  const full = `${sheetId}!${a1Range}`;
  const values = cells2D.map(row => row.map(c => {
    if (c === null || c === undefined) return '';
    if (typeof c === 'object' && 'value' in c) return (c.value !== undefined && c.value !== null) ? c.value : '';
    return c; // 纯数值
  }));
  const r = await feishu('PUT', `/open-apis/sheets/v2/spreadsheets/${CONFIG.feishuSpreadsheetToken}/values`,
    { valueRange: { range: full, values } });
  if (!r || r.code !== 0) throw new Error('写失败 ' + full + ': ' + JSON.stringify(r));
  return r;
}
function splitRange(a1Range) { const m = a1Range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)?$/); return m ? { c1: m[1], r1: +m[2], c2: m[3], r2: m[4] ? +m[4] : +m[2] } : null; }
function colIndex(letter) { let n = 0; for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function colLetter(idx) { let s = ''; idx++; while (idx > 0) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); } return s; }
function colSeqBetween(sp) { const a = colIndex(sp.c1), b = colIndex(sp.c2); const o = []; for (let i = a; i <= b; i++) o.push(colLetter(i)); return o; }
function rowNums(sp) { const o = []; for (let i = sp.r1; i <= sp.r2; i++) o.push(i); return o; }

// ==================== 邮件 / ZIP ====================
function fetchMailBySender(sender, pred) {
  return new Promise((resolve, reject) => {
    const imap = new Imap({ ...CONFIG.imap, tls: true, tlsOptions: { rejectUnauthorized: false }, connTimeout: 20000, authTimeout: 20000, timeout: 20000 });
    let settled = false;
    const done = (fn, val) => { if (!settled) { settled = true; fn(val); } };
    imap.once('ready', () => { imap.openBox('INBOX', true, err => {
      if (err) return done(reject, err);
      imap.search([['FROM', sender]], (err, uids) => {
        if (err) return done(reject, err);
        if (!uids.length) { imap.end(); return done(resolve, null); }
        const recent = uids.slice(-7).reverse();
        let done2 = 0, chosen = null;
        recent.forEach(uid => {
          const f = imap.fetch(uid, { bodies: '' });
          let raw = '';
          f.on('message', m => m.on('body', s => s.on('data', c => raw += c.toString('utf8'))));
          f.once('error', e => done(reject, e));
          f.once('end', () => { simpleParser(raw, (e, mail) => { if (e) return done(reject, e); done2++;
            if (!chosen && pred(mail.subject || '')) { chosen = mail; log('选中邮件', mail.subject, mail.date); }
            if (done2 === recent.length) { imap.end(); done(resolve, chosen); } }); });
        });
      });
    }); });
    imap.once('error', e => done(reject, e));
    imap.once('end', () => { if (!settled) done(resolve, null); });
    imap.connect();
  });
}
function httpsDownload(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, res => {
      if (res.statusCode !== 200) { res.resume(); return reject(new Error('HTTP ' + res.statusCode)); }
      const ch = []; res.on('data', c => ch.push(c)); res.on('end', () => resolve(Buffer.concat(ch)));
    }).on('error', reject);
  });
}
async function fetchZips() {
  const matchSubj = s => s.startsWith(CONFIG.dashboardSubjectPrefix) ||
    (CONFIG.dashboardSubjectAltPrefix && s.startsWith(CONFIG.dashboardSubjectAltPrefix));
  const mail = await withRetry(() => fetchMailBySender(CONFIG.dashboardMailFrom, matchSubj), '取看板邮件', 4, 3000);
  if (!mail) throw new Error('未找到看板邮件');
  const task = mail.text.match(/taskId=(\d+)/); const code = mail.text.match(/提取码[：:]\s*(\d+)/);
  if (!task || !code) throw new Error('未找到 taskId/提取码');
  const url = `https://shushuos.boltray.com/v1/ta/dashboard/daily/reportFileDownload?taskId=${task[1]}&downloadCode=${code[1]}&lang=zh-CN`;
  return new AdmZip(await httpsDownload(url));
}

// 通用重试（IMAP/HTTP 偶发失败时重试）
async function withRetry(fn, label, retries = 4, delayMs = 3000) {
  let lastErr;
  for (let i = 1; i <= retries; i++) {
    try { return await fn(); } catch (e) {
      lastErr = e; log(`${label} 第 ${i}/${retries} 次失败: ${e.message}`);
      if (i < retries) await sleep(delayMs * i);
    }
  }
  throw lastErr;
}

// ==================== CSV ====================
function parseCsv(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = []; let cur = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) { const c = text[i];
    if (inQ) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; } else field += c; }
    else { if (c === '"') inQ = true; else if (c === ',') { cur.push(field); field = ''; } else if (c === '\n') { cur.push(field); rows.push(cur); cur = []; field = ''; } else if (c === '\r') {} else field += c; }
  }
  if (field.length || cur.length) { cur.push(field); rows.push(cur); }
  return { rows };
}
function extractDashboard(zip) {
  const out = { base: {}, ios: {}, google: {} };
  for (const e of zip.getEntries()) {
    if (e.isDirectory) continue;
    const name = Buffer.from(e.rawEntryName).toString('utf-8');
    const text = e.getData().toString('utf-8');
    let key = null;
    if (name.includes(CONFIG.baseCsvTitle)) key = 'base';
    else if (name.includes(CONFIG.iosCsvTitle)) key = 'ios';
    else if (name.includes(CONFIG.googleCsvTitle)) key = 'google';
    if (!key) continue;
    const parsed = parseCsv(text);
    for (const row of parsed.rows) {
      const dateStr = row[0] || '';
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;
      out[key][dateStr] = row.slice(1).map(v => parseFloat(String(v).replace(/,/g, '')));
    }
  }
  return out;
}
function extractRetentionLtv(zip, csvNamePart) {
  const entry = zip.getEntries().find(e => { if (e.isDirectory) return false; const n = Buffer.from(e.rawEntryName).toString('utf-8'); return n.includes(csvNamePart) && /\.csv$/i.test(n) && !n.includes('付费'); });
  if (!entry) return { byDate: {}, total: {} };
  const rows = parseCsv(entry.getData().toString('utf-8')).rows;
  const hdr = rows[0]; const byDate = {}, total = {};
  for (const r of rows.slice(1)) { const date = r[0], ch = r[1]; let target;
    if (date === '阶段值') { target = total[ch] || (total[ch] = {}); }
    else if (/^\d{4}-\d{2}-\d{2}/.test(date)) { if (!byDate[date]) byDate[date] = {}; target = byDate[date][ch] || (byDate[date][ch] = {}); }
    else continue;
    for (let i = 4; i < hdr.length; i++) target[hdr[i]] = r[i];
  }
  return { byDate, total };
}
function parseAdXls(file) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const out = {};
  for (let i = 4; i < rows.length; i++) { const r = rows[i];
    const dateStr = (r[0] || '').trim(); const os_ = (r[2] || '').trim();
    const cost = parseFloat(String(r[3]).replace(/[$,]/g, '')) || 0;
    const acct = parseInt(String(r[4]).replace(/[,]/g, '')) || 0;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;
    if (!out[dateStr]) out[dateStr] = {};
    if (os_ === 'iOS') out[dateStr].iosCost = cost;
    else if (os_ === '安卓') { out[dateStr].gpCost = cost; out[dateStr].gpAdAcct = acct; }
  }
  return out;
}
async function fetchAdData() {
  const mail = await withRetry(() => fetchMailBySender(CONFIG.adMailFrom, s => s.trim() === CONFIG.adMailSubject), '取广告邮件', 4, 3000);
  if (!mail) return {};
  const att = (mail.attachments || []).find(a => /\.xls$/i.test(a.filename || ''));
  if (!att) return {};
  const tmpFile = path.join(os.tmpdir(), `daily-report-cloud-${Date.now()}.xls`);
  fs.writeFileSync(tmpFile, att.content);
  return parseAdXls(tmpFile);
}

// ==================== 留存/LTV 列映射 ====================
const RETENTION_KEYS = ['1日','2日','3日','4日','5日','6日','13日','29日','59日','89日','119日','149日','179日','209日','239日','269日','299日','329日','359日'];
const LTV_KEYS = ['当日','1日','2日','3日','4日','5日','6日','13日','29日','59日','89日','119日','149日','179日','209日','239日','269日','299日','329日','359日'];
function buildRetentionLtvMaps(startLetter) { const sIdx = colIndex(startLetter);
  const retCols = RETENTION_KEYS.map((k, i) => [colLetter(sIdx + i), k]);
  const ltvCols = LTV_KEYS.map((k, i) => [colLetter(sIdx + RETENTION_KEYS.length + i), k]);
  return { retCols, ltvCols };
}
function toNum(v) { if (v === undefined || v === null) return null; const s = String(v).trim(); if (s === '' || s === '-' || s === '--') return null; const n = parseFloat(s.replace(/%/g, '').replace(/,/g, '')); return isNaN(n) ? null : n; }
function buildRetentionLtvCells(retRow, ltvRow, maps) { const cells = [];
  for (const [col, key] of maps.retCols) cells.push({ col, value: toNum(retRow[key]) });
  for (const [col, key] of maps.ltvCols) cells.push({ col, value: toNum(ltvRow[key]) });
  return cells;
}

// ==================== 飞书读工具 ====================
async function readDateRows(sheetId, startRow, endRow) {
  const r = await readRange(sheetId, `B${startRow}:B${endRow}`);
  const map = {};
  r.rows.forEach((rowArr, i) => {
    const v = (rowArr[0] || {}).value;
    const dateKey = excelOrStringDate(v);
    if (dateKey) map[dateKey] = (r.row_indices[i] || (startRow + i));
  });
  return map;
}
// 兼容两种日期形式：字符串 "2026/8/29" 或 Excel 日期序列号 46263 → "YYYY-MM-DD"
function excelOrStringDate(v) {
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{4})\/(\d+)\/(\d+)$/);
    if (m) return `${m[1]}-${String(+m[2]).padStart(2, '0')}-${String(+m[3]).padStart(2, '0')}`;
    return null;
  }
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  return null;
}
async function readRangeGrid(sheetId, startCol, endCol, startRow, endRow) {
  const r = await readRange(sheetId, `${startCol}${startRow}:${endCol}${endRow}`);
  const grid = {};
  r.rows.forEach((rowArr, ri) => {
    const rowNum = r.row_indices[ri] || (startRow + ri);
    grid[rowNum] = {};
    r.col_indices.forEach((col, ci) => { grid[rowNum][col] = (rowArr[ci] || {}).value; });
  });
  return grid;
}
async function isCellEmpty(sheetId, cellAddr) {
  const m = cellAddr.match(/^([A-Z]+)(\d+)$/);
  const col = m[1], row = m[2];
  const r = await readRange(sheetId, `${col}${row}:${col}${row}`);
  const v = ((r.rows[0] || [])[0] || {}).value;
  return v === undefined || v === null || v === '';
}
async function findDateRow(sheetId, searchStr) {
  const [y, m, d] = searchStr.split('/');
  const key = `${y}-${String(+m).padStart(2, '0')}-${String(+d).padStart(2, '0')}`;
  const rows = await readDateRows(sheetId, 4, 700);
  return rows[key] || null;
}

// ==================== 行构造 ====================
function buildRibaoRow(values, n) {
  return values.slice(0, n).map((v, i) => ({ value: i >= 6 ? Math.round(v) : v }));
}
function buildIosCtoFRow(values) {
  return [ { value: Number(values[0].toFixed(2)) }, { value: values[1] }, { value: values[2] }, { value: values[3] } ];
}

// ==================== 主流程 ====================
async function main() {
  log('=== 启动 ' + CONFIG.label + ' 日报（云端 OpenAPI 版）===');
  const zip = await fetchZips(); log('ZIP 下载完成');
  const dashboard = extractDashboard(zip);
  const retention = extractRetentionLtv(zip, CONFIG.retentionCsvName);
  const ltv = extractRetentionLtv(zip, CONFIG.ltvCsvName);
  const adData = await fetchAdData();
  log(`看板: base=${Object.keys(dashboard.base).length} ios=${Object.keys(dashboard.ios).length} google=${Object.keys(dashboard.google).length} 留存=${Object.keys(retention.byDate).length} LTV=${Object.keys(ltv.byDate).length}`);

  const allDates = new Set([...Object.keys(dashboard.base), ...Object.keys(dashboard.ios), ...Object.keys(dashboard.google), ...Object.keys(adData)]);

  // ---- 日报表（先写，最重要） ----
  for (const date of [...allDates].sort()) {
    const [y, m, d] = date.split('-');
    const searchStr = `${y}/${parseInt(m)}/${parseInt(d)}`;
    const ribaoRow = await findDateRow(CONFIG.sheetIds.ribao, searchStr);
    if (ribaoRow) {
      const baseRow = dashboard.base[date];
      if (baseRow && baseRow.length >= CONFIG.ribaoCols) {
        if (await isCellEmpty(CONFIG.sheetIds.ribao, `C${ribaoRow}`)) {
          await writeRange(CONFIG.sheetIds.ribao, `C${ribaoRow}:${colLetter(colIndex('C') + CONFIG.ribaoCols - 1)}${ribaoRow}`, [buildRibaoRow(baseRow, CONFIG.ribaoCols)]);
          log('  日报表 C' + ribaoRow + ' 已写入');
        }
      }
      const ad = adData[date];
      if (ad) {
        const { colGpAcct, colIosCost, colGpCost } = CONFIG.adCol;
        // 只写 3 个值格（gp广告新增 / iOS消耗 / GP消耗），绝不触碰中间的公式列（AE-AH / W-Z）
        // 分开逐格写，避免连续 range 覆盖公式列
        if (await isCellEmpty(CONFIG.sheetIds.ribao, `${colGpAcct}${ribaoRow}`)) {
          await writeRange(CONFIG.sheetIds.ribao, `${colGpAcct}${ribaoRow}:${colGpAcct}${ribaoRow}`, [[ad.gpAdAcct || 0]]);
        }
        if (await isCellEmpty(CONFIG.sheetIds.ribao, `${colIosCost}${ribaoRow}`)) {
          await writeRange(CONFIG.sheetIds.ribao, `${colIosCost}${ribaoRow}:${colIosCost}${ribaoRow}`, [[Math.round(ad.iosCost || 0)]]);
        }
        if (await isCellEmpty(CONFIG.sheetIds.ribao, `${colGpCost}${ribaoRow}`)) {
          await writeRange(CONFIG.sheetIds.ribao, `${colGpCost}${ribaoRow}:${colGpCost}${ribaoRow}`, [[Math.round(ad.gpCost || 0)]]);
        }
        log('  日报表 广告列 ' + ribaoRow + ' 已写入');
      }
    }
    const iosRow = await findDateRow(CONFIG.sheetIds.ios, searchStr);
    if (iosRow) {
      const d = dashboard.ios[date];
      if (d && d.length >= 4 && await isCellEmpty(CONFIG.sheetIds.ios, `C${iosRow}`)) {
        await writeRange(CONFIG.sheetIds.ios, `C${iosRow}:F${iosRow}`, [buildIosCtoFRow(d)]); log('  iOS C' + iosRow + ' 已写入');
      }
    }
    const gRow = await findDateRow(CONFIG.sheetIds.google, searchStr);
    if (gRow) {
      const d = dashboard.google[date];
      if (d && d.length >= 4 && await isCellEmpty(CONFIG.sheetIds.google, `C${gRow}`)) {
        await writeRange(CONFIG.sheetIds.google, `C${gRow}:F${gRow}`, [buildIosCtoFRow(d)]); log('  Google C' + gRow + ' 已写入');
      }
    }
  }

  // ---- 留存/LTV（读现状→对比→只补缺失） ----
  log('补充留存/LTV 数据...');
  for (const sheetName of Object.keys(CONFIG.rlSheets)) {
    const { sheetId, channel, retentionStart } = CONFIG.rlSheets[sheetName];
    const maps = buildRetentionLtvMaps(retentionStart);
    const allCols = [...maps.retCols, ...maps.ltvCols];
    const firstCol = allCols[0][0], lastCol = allCols[allCols.length - 1][0];
    const dateRows = await readDateRows(sheetId, 4, 700);
    const grid = await readRangeGrid(sheetId, firstCol, lastCol, 3, 700);
    // 合计行（阶段值）：只写缺失列所在的最小区间，段内已有列保留原值
    const tSrc = buildRetentionLtvCells(retention.total[channel] || {}, ltv.total[channel] || {}, maps);
    const tCur = grid[CONFIG.rlTotalRow] || {};
    const tMissing = tSrc.filter(c => c.value !== null && (tCur[c.col] === undefined || tCur[c.col] === null || tCur[c.col] === ''));
    if (tMissing.length) await writeMissingRange(sheetId, CONFIG.rlTotalRow, tSrc, tMissing, tCur, allCols);
    // 每日行：只写该行数据源有值且飞书为空的列（迟到留存在此补上）
    let cnt = 0;
    for (const dateKey of Object.keys(retention.byDate).sort()) {
      const dateOnly = dateKey.slice(0, 10); const row = dateRows[dateOnly]; if (!row) continue;
      const src = buildRetentionLtvCells(retention.byDate[dateKey][channel] || {}, ltv.byDate[dateKey][channel] || {}, maps);
      if (!src.some(c => c.value !== null)) continue;
      const cur = grid[row] || {};
      const missing = src.filter(c => c.value !== null && (cur[c.col] === undefined || cur[c.col] === null || cur[c.col] === ''));
      if (!missing.length) continue;
      // 只写缺失列所在的最小区间，绝不写区间外（保护更早的已有列），区间内已有列保留原值
      await writeMissingRange(sheetId, row, src, missing, cur, allCols);
      cnt++;
    }
    log(`  [留存/LTV] ${sheetName}(${sheetId}) start=${retentionStart}: 补缺失 ${cnt} 行`);
  }
  log('=== ' + CONFIG.label + ' 日报完成，推送通知 ===');
  await sendWebhook(`✅ ${CONFIG.label}日报执行完成\n${CONFIG.label}数据表：${CONFIG.dataTableUrl}`);
}

// ==================== 飞书 webhook 推送 ====================
const WEBHOOK_URL = process.env.FEISHU_WEBHOOK_URL || '';
function sendWebhook(text) {
  return new Promise((resolve) => {
    if (!WEBHOOK_URL) { log('未配置 FEISHU_WEBHOOK_URL，跳过推送'); return resolve(); }
    try {
      const body = JSON.stringify({ msg_type: 'text', content: { text } });
      const req = https.request(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
      }, res => { res.resume(); res.on('end', resolve); });
      req.on('error', () => resolve());
      req.setTimeout(10000, () => { req.destroy(); resolve(); });
      req.write(body);
      req.end();
    } catch (e) { resolve(); }
  });
}

// 只写该行缺失列所在的最小区间；区间内每列值 = 缺失→源值，非缺失→保留当前（绝不用 '' 覆盖已有值）
async function writeMissingRange(sheetId, row, srcCells, missingCells, cur, allCols) {
  const missingCols = missingCells.map(c => c.col);
  const minI = Math.min(...missingCols.map(c => colIndex(c)));
  const maxI = Math.max(...missingCols.map(c => colIndex(c)));
  const startCol = colLetter(minI), endCol = colLetter(maxI);
  // 该范围内每列：缺失→源值；否则保留 cur（有值给值，空才给 ''）
  const segVals = [];
  for (let i = minI; i <= maxI; i++) {
    const col = colLetter(i);
    const missCell = missingCells.find(c => c.col === col);
    if (missCell) { segVals.push(missCell.value); continue; }
    const curVal = cur[col];
    segVals.push((curVal !== undefined && curVal !== null && curVal !== '') ? curVal : '');
  }
  // 用 invoke_write 写"值+0.00格式"，避免预置行百分比格式把 39.67 渲染成 3967.00%
  const range = `${sheetId}!${startCol}${row}:${endCol}${row}`;
  const input = {
    cells: [segVals.map(v => ({ value: v, cell_styles: { number_format: '0.00' } }))],
    excel_id: CONFIG.feishuSpreadsheetToken,
    range,
    sheet_id: sheetId,
  };
  const r = await feishu('POST', `/open-apis/sheet_ai/v2/spreadsheets/${CONFIG.feishuSpreadsheetToken}/tools/invoke_write`,
    { input: JSON.stringify(input), tool_name: 'set_cell_range' });
  if (!r || r.code !== 0) throw new Error('invoke_write 失败 ' + range + ': ' + JSON.stringify(r));
}

main().then(() => process.exit(0)).catch(e => {
  console.error('FATAL:', e);
  sendWebhook(`❌ ${CONFIG.label}日报执行失败：${e.message}\n${CONFIG.label}数据表：${CONFIG.dataTableUrl}`).then(() => process.exit(1));
});
