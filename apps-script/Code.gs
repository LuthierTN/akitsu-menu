/**
 * 秋津手作り革工房 メニュー：注文受付用 Apps Script
 * 注文ページから送られた注文を、このスプレッドシートの「注文」シートに
 * カートの1項目＝1行で追加し、受付番号（A-001, A-002…）を返します。
 * 「集計」シートには、板ぬり絵の大・小が何枚かを状態ごとに数えて表示します。
 */
const SHEET_NAME = '注文';
const SUMMARY_NAME = '集計';
const STATUSES = ['受付', '加工中', '完了', 'キャンセル'];
const HEADERS = ['受付日時', '受付番号', '状態', 'お名前', '電話番号', '受け取り',
  '商品', 'サイズ', '絵柄', '枠', '色', '表の文字', '裏の文字', '書体', '彫り方',
  '枚数', '単価', '小計', '注文合計'];
const COL_STATUS = HEADERS.indexOf('状態') + 1;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const order = JSON.parse(e.postData.contents);
    if (!order.お名前 || !order.電話番号 || !Array.isArray(order.明細) || !order.明細.length) {
      return json_({ ok: false, error: 'invalid order' });
    }
    const sheet = getSheet_();
    const props = PropertiesService.getScriptProperties();
    const seq = Number(props.getProperty('seq') || 0) + 1;
    props.setProperty('seq', String(seq));
    const orderNo = 'A-' + String(seq).padStart(3, '0');
    const now = new Date();
    const rows = order.明細.map(it => [
      now, orderNo, '受付', order.お名前, "'" + order.電話番号, order.受け取り,
      it.商品, it.サイズ, it.絵柄, it.枠, it.色, it.表の文字, it.裏の文字, it.書体, it.彫り方,
      it.枚数, it.単価, it.小計, order.合計
    ]);
    const start = lastOrderRow_(sheet) + 1;
    sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    sheet.getRange(start, COL_STATUS, rows.length, 1).setDataValidation(statusRule_());
    return json_({ ok: true, orderNo: orderNo });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 受付番号の列（B列）で最後に埋まっている行。ほかの列に何か書かれていても影響されない */
function lastOrderRow_(sheet) {
  const n = sheet.getLastRow();
  if (n < 2) return 1;
  const vals = sheet.getRange(2, 2, n - 1, 1).getValues();
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
    const L = String.fromCharCode(64 + HEADERS.indexOf(name) + 1);
    return "'" + SHEET_NAME + "'!" + L + '2:' + L;
  };
  const sum = (size, statusCond) => '=SUMIFS(' + c('枚数') + ',' + c('商品') + ',"板ぬり絵",' +
    c('サイズ') + ',"' + size + '",' + c('状態') + ',' + statusCond + ')';
  const rows = [['板ぬり絵', '受付', '加工中', '完了', '合計（キャンセル除く）', 'キャンセル']];
  ['大', '小'].forEach((size, i) => rows.push([size,
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
 * 最初に一度だけ実行：「注文」シートを作り直し（今ある「注文」シートは消えます）、
 * 状態のプルダウンと色分け、「集計」シートを用意して、受付番号を A-001 から振り直す
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_NAME);
  const temp = old && ss.getSheets().length === 1 ? ss.insertSheet() : null;
  if (old) ss.deleteSheet(old);
  const sheet = getSheet_();
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  setupSummary_();
  if (temp) ss.deleteSheet(temp);
  resetOrderNumber();
}

/** 受付番号を A-001 から振り直したいときに一度だけ実行 */
function resetOrderNumber() {
  PropertiesService.getScriptProperties().deleteProperty('seq');
}
