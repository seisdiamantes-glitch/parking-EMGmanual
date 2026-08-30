// =============================================
// 緊急通達（emergency_notice.html）連携用 GAS
//
// 【初期設定】
// 1. 記録先にしたい Google スプレッドシートを開く
// 2. 拡張機能 → Apps Script を開く
// 3. このファイルの内容を貼り付けて保存
// 4. Apps Script エディタで setup() を一度だけ実行し、
//    「緊急通達」シートとヘッダー行を作成する
//    （実行時に権限の承認を求められたら許可する）
// 5. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
//      - 実行ユーザー：自分
//      - アクセスできるユーザー：全員
//    でデプロイし、発行された URL を
//    emergency_notice.html の GAS_WEB_APP_URL に貼り付ける
// =============================================

const SHEET_NAME = '緊急通達';
const HEADERS = ['タイムスタンプ', '駐車場', '種別', '氏名', '補足'];
const URGENT_LABEL = '緊急要請';

const TYPE_COL = 3; // C列 = 種別

function setup() {
  getOrCreateSheet();
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getRange(1, 1).getValue() !== HEADERS[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sheet;
}

// emergency_notice.html からの送信を受け取り、1行追記する
function doPost(e) {
  const sheet = getOrCreateSheet();
  const data = JSON.parse(e.postData.contents);

  const row = [
    new Date(),
    data.park || '',
    data.type || '',
    data.name || '',
    data.note || '',
  ];
  sheet.appendRow(row);
  applyTypeFormatting(sheet, sheet.getLastRow(), data.type);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// スプレッドシート上で種別（C列）を直接入力・修正した場合にも
// 同じ書式ルールを適用する
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getRow() === 1) return; // ヘッダー行は対象外
  if (e.range.getColumn() !== TYPE_COL) return;

  const type = sheet.getRange(e.range.getRow(), TYPE_COL).getValue();
  applyTypeFormatting(sheet, e.range.getRow(), type);
}

// 「緊急要請」の行だけ太字・赤字にする。それ以外は通常表示に戻す。
function applyTypeFormatting(sheet, row, type) {
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const range = sheet.getRange(row, 1, 1, lastCol);
  if (type === URGENT_LABEL) {
    range.setFontWeight('bold').setFontColor('#C0392B');
  } else {
    range.setFontWeight('normal').setFontColor('#000000');
  }
}
