/**
 * 秋津手作り革工房 メニュー：注文受付用 Apps Script
 * 注文ページから送られた注文を、このスプレッドシートの「注文」シートに
 * カートの1項目＝1行で追加します。お客さまの整理券番号で管理します。
 * 「集計」シートには、板ぬり絵の大・小が何枚かを状態ごとに数えて表示します。
 * 管理ダッシュボード（admin.html）からの操作は、スクリプト プロパティ ADMIN_PIN と一致したときだけ受け付けます。
 */
const SHEET_NAME = '注文';
const SUMMARY_NAME = '集計';
const STATUSES = ['受付', '加工中', '完了', 'キャンセル'];
const SIZES = ['大', '小'];
const HEADERS = ['受付日時', '整理券番号', '状態', 'お名前', '受け取り',
  '商品', 'サイズ', '絵柄', '枠', '色', '表の文字', '裏の文字', '書体', '彫り方',
  '枚数', '単価', '小計', '注文合計', '完了日時'];
const col_ = name => HEADERS.indexOf(name) + 1;
const COL_TICKET = col_('整理券番号');
const COL_STATUS = col_('状態');
const COL_DONE = col_('完了日時');

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  try {
    return json_(req.action && req.action !== 'order' ? admin_(req) : order_(req));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 管理ダッシュボードの読み込みは GET でも受け付ける（?action=dashboard&pin=…） */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'dashboard') return json_({ ok: false, error: 'unknown action' });
  try {
    return json_(admin_({ action: 'dashboard', pin: p.pin }));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 整理券番号：全角数字も受け付け、1〜9999 の数にそろえる。だめなら null */
function normTicket_(v) {
  const s = String(v == null ? '' : v).trim()
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  if (!/^\d{1,4}$/.test(s) || Number(s) < 1) return null;
  return Number(s);
}

function order_(order) {
  const ticket = normTicket_(order.整理券番号);
  const name = String(order.お名前 || '').trim();
  if (!ticket || !name || !Array.isArray(order.明細) || !order.明細.length) {
    return { ok: false, error: 'invalid order' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    if (readRows_(sheet).some(r => r.ticket === ticket && r.status !== 'キャンセル')) {
      return { ok: false, error: 'duplicate', ticket: ticket };
    }
    const now = new Date();
    const rows = order.明細.map(it => [
      now, ticket, '受付', name, order.受け取り || '当日ブースで受け取り',
      it.商品, it.サイズ, it.絵柄, it.枠, it.色, it.表の文字, it.裏の文字, it.書体, it.彫り方,
      it.枚数, it.単価, it.小計, order.合計, ''
    ]);
    const start = lastOrderRow_(sheet) + 1;
    sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    sheet.getRange(start, COL_STATUS, rows.length, 1).setDataValidation(statusRule_());
    return { ok: true, ticket: ticket };
  } finally {
    lock.releaseLock();
  }
}

/* ===== 管理ダッシュボード用 ===== */

function admin_(req) {
  const props = PropertiesService.getScriptProperties();
  const pin = props.getProperty('ADMIN_PIN');
  if (!pin) return { ok: false, error: 'no pin' };
  if (String(req.pin || '') !== pin) return { ok: false, error: 'unauthorized' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    if (req.action === 'setStock') {
      const count = Number(req.count);
      if (SIZES.indexOf(req.size) < 0 || !Number.isInteger(count) || count < 0) return { ok: false, error: 'invalid stock' };
      props.setProperty('stock_' + req.size, JSON.stringify({ count: count, setAt: Date.now() }));
    } else if (req.action === 'setCall') {
      const n = Number(req.number);
      if (!Number.isInteger(n) || n < 0 || n > 9999) return { ok: false, error: 'invalid number' };
      props.setProperty('currentCall', String(n));
    } else if (req.action === 'setStatus') {
      const ticket = normTicket_(req.ticket);
      if (!ticket || STATUSES.indexOf(req.status) < 0) return { ok: false, error: 'invalid status' };
      const hits = readRows_(sheet).filter(r => r.ticket === ticket);
      if (!hits.length) return { ok: false, error: 'not found' };
      const now = new Date();
      hits.forEach(r => {
        sheet.getRange(r.row, COL_STATUS).setValue(req.status);
        if (req.status === '完了') {
          if (!r.doneAt) sheet.getRange(r.row, COL_DONE).setValue(now);
        } else {
          sheet.getRange(r.row, COL_DONE).setValue('');
        }
      });
    } else if (req.action !== 'dashboard') {
      return { ok: false, error: 'unknown action' };
    }
    return dashboard_(sheet, props);
  } finally {
    lock.releaseLock();
  }
}

function dashboard_(sheet, props) {
  const stock = {};
  SIZES.forEach(s => {
    const v = props.getProperty('stock_' + s);
    stock[s] = v ? JSON.parse(v) : null;
  });
  const call = props.getProperty('currentCall');
  const rows = readRows_(sheet).filter(r => r.product === '板ぬり絵').map(r => ({
    ticket: r.ticket, name: r.name, size: r.size, qty: r.qty, status: r.status,
    design: r.design, frame: r.frame, orderedAt: r.orderedAt, doneAt: r.doneAt
  }));
  return { ok: true, stock: stock, currentCall: call == null ? null : Number(call), rows: rows, serverTime: Date.now() };
}

/** 「注文」シートの行を読みやすい形で返す（row はシート上の行番号） */
function readRows_(sheet) {
  const last = lastOrderRow_(sheet);
  if (last < 2) return [];
  const ms = v => (v instanceof Date ? v.getTime() : null);
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues().map((v, i) => ({
    row: i + 2,
    orderedAt: ms(v[col_('受付日時') - 1]),
    ticket: Number(v[COL_TICKET - 1]) || null,
    status: String(v[COL_STATUS - 1]),
    name: String(v[col_('お名前') - 1]),
    product: String(v[col_('商品') - 1]),
    size: String(v[col_('サイズ') - 1]),
    design: String(v[col_('絵柄') - 1]),
    frame: String(v[col_('枠') - 1]),
    qty: Number(v[col_('枚数') - 1]) || 0,
    doneAt: ms(v[COL_DONE - 1])
  })).filter(r => r.ticket);
}

/** シートで「状態」を手で変えたときも、完了日時を記録・消去する */
function onEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEET_NAME) return;
  const c0 = range.getColumn(), c1 = c0 + range.getNumColumns() - 1;
  if (COL_STATUS < c0 || COL_STATUS > c1) return;
  const sheet = range.getSheet();
  const r0 = Math.max(range.getRow(), 2), r1 = range.getRow() + range.getNumRows() - 1;
  if (r1 < r0) return;
  const status = sheet.getRange(r0, COL_STATUS, r1 - r0 + 1, 1).getValues();
  const done = sheet.getRange(r0, COL_DONE, r1 - r0 + 1, 1).getValues();
  const now = new Date();
  const out = status.map((s, i) => [s[0] === '完了' ? (done[i][0] || now) : '']);
  sheet.getRange(r0, COL_DONE, out.length, 1).setValues(out);
}

/* ===== シートの準備 ===== */

/** 整理券番号の列で最後に埋まっている行。ほかの列に何か書かれていても影響されない */
function lastOrderRow_(sheet) {
  const n = sheet.getLastRow();
  if (n < 2) return 1;
  const vals = sheet.getRange(2, COL_TICKET, n - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '') return i + 2;
  return 1;
}

function statusRule_() {
  return SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.getRange(2, col_('受付日時'), sheet.getMaxRows() - 1, 1).setNumberFormat('M/d H:mm');
    sheet.getRange(2, COL_DONE, sheet.getMaxRows() - 1, 1).setNumberFormat('M/d H:mm');
    setupStatusColors_(sheet);
  }
  return sheet;
}

/** 状態ごとに行の色を変える（加工中＝黄、完了＝緑、キャンセル＝灰） */
function setupStatusColors_(sheet) {
  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, HEADERS.length);
  const col = String.fromCharCode(64 + COL_STATUS);
  const rule = (text, color) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + col + '2="' + text + '"').setBackground(color).setRanges([range]).build();
  sheet.setConditionalFormatRules([
    rule('加工中', '#fff2cc'), rule('完了', '#d9ead3'), rule('キャンセル', '#e0e0e0')
  ]);
}

/** 「集計」シート：板ぬり絵の大・小の枚数を状態ごとに数える（数式なので自動で更新される） */
function setupSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SUMMARY_NAME);
  if (!sh) sh = ss.insertSheet(SUMMARY_NAME);
  sh.clear();
  const c = name => {
    const L = String.fromCharCode(64 + col_(name));
    return "'" + SHEET_NAME + "'!" + L + '2:' + L;
  };
  const sum = (size, statusCond) => '=SUMIFS(' + c('枚数') + ',' + c('商品') + ',"板ぬり絵",' +
    c('サイズ') + ',"' + size + '",' + c('状態') + ',' + statusCond + ')';
  const rows = [['板ぬり絵', '受付', '加工中', '完了', '合計（キャンセル除く）', 'キャンセル']];
  SIZES.forEach((size, i) => rows.push([size,
    sum(size, '"受付"'), sum(size, '"加工中"'), sum(size, '"完了"'),
    '=B' + (i + 2) + '+C' + (i + 2) + '+D' + (i + 2), sum(size, '"キャンセル"')]));
  rows.push(['計', '=B2+B3', '=C2+C3', '=D2+D3', '=E2+E3', '=F2+F3']);
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#f3f3f3');
  sh.getRange(rows.length, 1, 1, rows[0].length).setFontWeight('bold');
  sh.getRange(1, 4, rows.length, 1).setBackground('#d9ead3');
  sh.setColumnWidth(5, 170);
  sh.setFrozenRows(1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 準備のときに実行：今の「注文」シートを「注文_旧_日付」に名前を変えて残し、
 * 新しい「注文」シート（状態のプルダウンと色分けつき）と「集計」シートを作ります。
 * 在庫と呼び出し番号の登録も消えます。注文データは消しません。
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_NAME);
  if (old) {
    const base = SHEET_NAME + '_旧_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    let name = base;
    for (let i = 2; ss.getSheetByName(name); i++) name = base + '_' + i;
    old.setName(name);
  }
  const sheet = getSheet_();
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  setupSummary_();
  const props = PropertiesService.getScriptProperties();
  ['seq', 'currentCall'].concat(SIZES.map(s => 'stock_' + s)).forEach(k => props.deleteProperty(k));
  if (!props.getProperty('ADMIN_PIN')) console.log('管理用の暗証番号がまだありません。setAdminPin を実行してください。');
}

/**
 * 管理ダッシュボードの暗証番号を決める：下の '' の中に好きな番号（4桁以上）を書いて、
 * この関数を一度だけ実行し、実行が終わったら '' に戻して保存してください。
 * （暗証番号をコードに残さないため。プロジェクトの設定 → スクリプト プロパティ の ADMIN_PIN でも変えられます）
 */
function setAdminPin() {
  const pin = '';
  if (!/^\w{4,}$/.test(pin)) throw new Error("setAdminPin の '' の中に、4文字以上の暗証番号を書いてから実行してください");
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  console.log('暗証番号を登録しました。コードの暗証番号は消して保存しておいてください。');
}
