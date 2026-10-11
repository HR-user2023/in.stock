/**
 * 髮廊點貨本 — Google 試算表後端
 * 用法：在試算表「擴充功能 → Apps Script」把這段全部貼上 → 部署成「網頁應用程式」
 *       執行身分：我　誰可以存取：所有人　→ 把拿到的網址貼進網頁的「設定」
 * 第一次被呼叫時會自動建立「品項」「叫貨紀錄」「異動紀錄」「廠商」「設定」分頁。
 */
const T = {
  items: { sheet: '品項', cols: ['id', 'code', 'supplier', 'name', 'category', 'unit', 'price', 'safety', 'count', 'updatedAt', 'note'],
    head: ['ID', '編號', '廠商', '品名', '分類', '單位', '單價', '安全庫存', '目前庫存', '最後更新', '備註'],
    widths: [90, 90, 90, 260, 70, 50, 70, 70, 70, 130, 160] },
  orders: { sheet: '叫貨紀錄', cols: ['id', 'orderNo', 'orderDate', 'createdAt', 'supplier', 'itemId', 'name', 'qty', 'price',
    'arrivedAt', 'recvQty', 'recvPrice', 'check', 'by'],
    head: ['ID', '訂單編號', '叫貨日期', '建立時間', '廠商', '品項ID', '品名', '叫貨數量', '叫貨單價',
      '到貨時間', '實到數量', '實際單價', '核對結果', '經手人'],
    widths: [90, 120, 100, 130, 90, 90, 220, 70, 70, 130, 70, 70, 140, 80] },
  moves: { sheet: '異動紀錄', cols: ['id', 'at', 'type', 'itemId', 'name', 'supplier', 'qty', 'after', 'ref', 'by', 'note'],
    head: ['ID', '時間', '類型', '品項ID', '品名', '廠商', '數量', '異動後庫存', '相關訂單', '經手人', '備註'],
    widths: [90, 130, 70, 90, 220, 90, 60, 80, 120, 80, 160] },
  suppliers: { sheet: '廠商', cols: ['name', 'min', 'contact', 'note'],
    head: ['廠商', '最低出貨金額', '聯絡方式', '備註'], widths: [140, 110, 200, 200] }
};
const DATE_COLS = { updatedAt: 1, createdAt: 1, arrivedAt: 1, at: 1 };
const DEFAULT_SUPPLIERS = [['台灣歌薇', 3000], ['威娜', 3000], ['哥德式', 3000], ['萊肯', 3000], ['覺亞', 4000],
  ['京喚羽', 3000], ['金韓今', 3000], ['好日梓', 3000], ['資生堂', 3000]];

function doGet(e) { return out_(handle_((e && e.parameter) || {})); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  return out_(handle_(body));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(req) {
  try {
    setup_();
    if (req.ops && req.ops.length) {
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      try { req.ops.forEach(applyOp_); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
    }
    return { ok: true, data: readAll_() };
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}

/* ---------- 讀 ---------- */
function readAll_() {
  const out = {};
  Object.keys(T).forEach(k => out[k] = read_(k));
  const st = sheet_('設定');
  out.shop = st ? String(st.getRange('B1').getValue() || '') : '';
  return out;
}
function read_(k) {
  const t = T[k], sh = sheet_(t.sheet), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, t.cols.length).getValues().filter(r => r[0] !== '').map(r => {
    const o = {}; t.cols.forEach((c, i) => { const v = r[i]; o[c] = v instanceof Date ? (c === 'orderDate' ? fmtDate_(v) : v.getTime()) : v; }); return o;
  });
}
function fmtDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd'); }

/* ---------- 寫（網頁把動作排隊送來） ---------- */
function applyOp_(op) {
  const d = op.data || {}, k = op.table;
  switch (op.type) {
    case 'upsert': return upsert_(k, d);
    case 'delete': return del_(k, d[T[k].cols[0]]);
    case 'adjust': { // 庫存加減（多台手機同時操作也不會蓋掉）
      const t = T.items, sh = sheet_(t.sheet), row = findRow_(sh, d.id); if (!row) return;
      const ci = t.cols.indexOf('count') + 1, ui = t.cols.indexOf('updatedAt') + 1;
      const now = Math.max(0, (Number(sh.getRange(row, ci).getValue()) || 0) + (Number(d.delta) || 0));
      sh.getRange(row, ci).setValue(now); sh.getRange(row, ui).setValue(new Date(Number(d.at) || Date.now()));
      if (d.move) { d.move.after = now; upsert_('moves', d.move); }
      return;
    }
    case 'shop': return sheet_('設定').getRange('B1').setValue(String(d.shop || ''));
  }
}
function upsert_(k, d) {
  const t = T[k], sh = sheet_(t.sheet), key = d[t.cols[0]], row = findRow_(sh, key);
  if (row) {
    const cur = sh.getRange(row, 1, 1, t.cols.length).getValues()[0];
    sh.getRange(row, 1, 1, t.cols.length).setValues([t.cols.map((c, i) => (c in d) ? val_(c, d[c]) : cur[i])]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, t.cols.length).setValues([t.cols.map(c => (c in d) ? val_(c, d[c]) : '')]);
  }
}
function val_(c, v) { return DATE_COLS[c] ? (v ? new Date(Number(v)) : '') : (v === null || v === undefined ? '' : v); }
function del_(k, key) { const sh = sheet_(T[k].sheet), row = findRow_(sh, key); if (row) sh.deleteRow(row); }
function findRow_(sh, key) {
  const last = sh.getLastRow(); if (last < 2) return 0;
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(key)) return i + 2;
  return 0;
}
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }

/* ---------- 第一次：建立分頁 ---------- */
function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = Object.keys(T).map(k => T[k].sheet).concat(['設定']);
  if (all.every(n => ss.getSheetByName(n))) return;
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    Object.keys(T).forEach(k => {
      const t = T[k]; if (ss.getSheetByName(t.sheet)) return;
      const sh = ss.insertSheet(t.sheet);
      sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold').setBackground('#0F9488').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
      t.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
      sh.getRange('A:B').setNumberFormat('@');
      t.cols.forEach((c, i) => { if (DATE_COLS[c]) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy/mm/dd hh:mm'); });
      if (k === 'orders') sh.getRange('C:C').setNumberFormat('@');
      if (k === 'suppliers') sh.getRange(2, 1, DEFAULT_SUPPLIERS.length, 2).setValues(DEFAULT_SUPPLIERS);
    });
    if (!ss.getSheetByName('設定')) {
      const st = ss.insertSheet('設定');
      st.getRange(1, 1, 2, 2).setValues([['店名', ''], ['（店名會顯示在叫貨清單上）', '']]);
      st.getRange('A1').setFontWeight('bold'); st.setColumnWidth(1, 220); st.setColumnWidth(2, 200);
    }
    ss.getSheets().forEach(s => { if (all.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && ss.getSheets().length > all.length) ss.deleteSheet(s); });
  } finally { lock.releaseLock(); }
}
