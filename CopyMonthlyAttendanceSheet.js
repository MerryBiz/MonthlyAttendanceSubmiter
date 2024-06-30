/**
 * 2023.9.21 作成
 * @n.yasuda
 * 勤務実績表シートコピースクリプト */

function copyMonthlyAttendanceSheet(staffSheet) {
  // テンプレファイルと勤務実績表のIDを設定
  var sourceSpreadsheetId = '1AYQmjaeMcjYsYJyTg4Mmn89ZxYRIfaHwtRssKUKzdvo';
  var targetSpreadsheetId = staffSheet;
  var newSheetName  = getNextMonth();
  
  // テンプレファイルから各種情報シートを取得
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheets()[0]; // 一番左側のシート
  
  // 勤務実績表に各種情報シートをコピー
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  if (targetSpreadsheet.getSheetByName(newSheetName)) {
    Browser.msgBox('そのシートは既に存在しています');
  } else {
  
  var newSheet = sourceSheet.copyTo(targetSpreadsheet);
  
  
  // 新しいシートの名前を設定（任意）
  newSheet.setName(newSheetName); // シート名を適切に変更

  // 勤務実績表に新しいシートを挿入
  targetSpreadsheet.setActiveSheet(newSheet);
  targetSpreadsheet.moveActiveSheet(0); // 先頭に移動
  targetSpreadsheet.getSheetByName(newSheetName).getRange("K2").setValue(newSheetName);
  Logger.log('シートをコピーしました。');
  }
}


function getNextMonth() {
  var today = new Date(); // 現在の日付を取得
  var nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1); // 来月の1日を取得

  var nextMonthYear = nextMonth.getFullYear();
  var nextMonthMonth = nextMonth.getMonth() + 1; // 月は0から始まるため、1を加える
  var newSheetName = nextMonthYear + "年" + nextMonthMonth + "月";
  Logger.log(nextMonthYear + "年" + nextMonthMonth + "月");
  return newSheetName;
}

