/**
 * 秋津手作り革工房 メニュー：注文受付用 Apps Script
 * 注文ページから送られた注文を、このスプレッドシートの「注文」シートに
 * カートの1項目＝1行で追加し、受付番号（A-001, A-002…）を返します。
 */
const SHEET_NAME = '注文';
const HEADERS = ['受付日時', '受付番号', 'お名前', '電話番号', '受け取り',
  '商品', 'サイズ', '絵柄', '枠', '色', '表の文字', '裏の文字', '書体', '彫り方',
  '枚数', '単価', '小計', '注文合計', '状態'];

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
      now, orderNo, order.お名前, "'" + order.電話番号, order.受け取り,
      it.商品, it.サイズ, it.絵柄, it.枠, it.色, it.表の文字, it.裏の文字, it.書体, it.彫り方,
      it.枚数, it.単価, it.小計, order.合計, '受付'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json_({ ok: true, orderNo: orderNo });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 受付番号を A-001 から振り直したいときに一度だけ実行 */
function resetOrderNumber() {
  PropertiesService.getScriptProperties().deleteProperty('seq');
}
