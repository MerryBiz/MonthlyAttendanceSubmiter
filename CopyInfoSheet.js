/**
 * 2023.9.21 作成
 * @n.yasuda
 * 各種情報シートコピースクリプト */

function copyInfoSheet(staffSheet) {

  // テンプレファイルと勤務実績表のIDを設定
  var sourceSpreadsheetId = '1AYQmjaeMcjYsYJyTg4Mmn89ZxYRIfaHwtRssKUKzdvo';
  var targetSpreadsheetId = staffSheet;
  var newSheetName  = '各種情報';
  
  // テンプレファイルから各種情報シートを取得
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName('各種情報'); // シート名を適切に変更
  
  // 勤務実績表に各種情報シートをコピー
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  if (targetSpreadsheet.getSheetByName(newSheetName)) {
    Browser.msgBox('そのシートは既に存在しています');
  } else {
  
  var newSheet = sourceSheet.copyTo(targetSpreadsheet);
  
  
  // 新しいシートの名前を設定（任意）
  newSheet.setName(newSheetName); // シート名を適切に変更

  // 勤務実績表に新しいシートを右から2番目に挿入
  var sheetNum = targetSpreadsheet.getNumSheets();
  targetSpreadsheet.setActiveSheet(newSheet);
  targetSpreadsheet.moveActiveSheet(sheetNum - 1); // 後ろから2番目に移動
  
  Logger.log('シートをコピーしました。');

  // 作成した各種情報シートをロックする
  protectSheetWithEditors(targetSpreadsheet);
  // 基本情報シートを非表示にする
  hideSheet(targetSpreadsheet);
  }
}

// シートをロックする
function protectSheetWithEditors(targetSpreadsheet) {
  var sheetName = "各種情報"; // ロックするシートの名前を宣言
  var editors = ["natsuki.yasuda@merrybiz.jp", "yuko.kamikawa@merrybiz.jp", "sayaka.kubota@merrybiz.jp"]; // ロックを解除できる編集者のメールアドレス

  var sheet = targetSpreadsheet.getSheetByName(sheetName);

  if (sheet) {
    var protection = sheet.protect().setDescription('ここから情報を更新することはできません');

    // ロックを解除できる編集者を設定
    protection.addEditors(editors);

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } else {
    Logger.log("指定されたシートが見つかりませんでした。");
  }
}

// 基本情報シートを非表示にする
function hideSheet(targetSpreadsheet) {
  var sheetName = "基本情報"; // 基本情報シートの名前を宣言
  var sheet = targetSpreadsheet.getSheetByName(sheetName);

  if (sheet) {
    sheet.hideSheet(); // シートを非表示にする
    Logger.log(sheetName + " を非表示にしました。");
  } else {
    Logger.log(sheetName + " が見つかりませんでした。");
  }
}

