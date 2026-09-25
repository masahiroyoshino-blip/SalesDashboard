// 16桁のランダムな英数字を生成する共通関数
function generateRandomId() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var randomId = '';
  for (var i = 0; i < 16; i++) {
    randomId += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return randomId;
}

// 【今後の自動用】行が編集された際に自動でIDを付与する関数
function generateUniqueIdOnEdit(e) {
  if (!e || !e.source) return;
  
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== '案件マスタ') return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  if (row <= 1 || col === 1) return;

  var idCell = sheet.getRange(row, 1);

  if (!idCell.getValue()) {
    idCell.setValue(generateRandomId());
  }
}

// 【初回・手動用】現在空欄になっているA列のIDを一括で埋める関数
function fillMissingIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('案件マスタ');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', '「案件マスタ」シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データ（2行目以降）がない場合は終了

  // A2セルから最終行までのA列のデータを取得
  var range = sheet.getRange(2, 1, lastRow - 1, 1);
  var values = range.getValues();
  var count = 0;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === '') { // 空欄の場合のみIDを生成
      values[i][0] = generateRandomId();
      count++;
    }
  }

  // 変更があった場合のみシートに一括書き戻し
  if (count > 0) {
    range.setValues(values);
    SpreadsheetApp.getUi().alert('完了', count + '件の空欄にIDを付番しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('確認', '空欄のIDはありませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// スプレッドシートを開いたときに、手動実行用のメニューを追加する
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタム機能')
    .addItem('空欄のIDを一括付番する', 'fillMissingIds')
    .addToUi();
}
