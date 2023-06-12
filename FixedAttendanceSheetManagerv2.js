function submitWithoutEnquete() {
  console.time("TOTAL EXECUTION TIME");
  var canExecute = askExecutable();
  if (canExecute) {
    var result = submitAttendancev2();
    showResultMessage(result);
  }
  console.timeEnd("TOTAL EXECUTION TIME");

}


// 収集処理
function submitAttendancev2() {

  var targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  console.log("■■ 処理対象ファイル名：" + targetSpreadSheet.getName());

  var currentAttendanceSheet = targetSpreadSheet.getActiveSheet();
  if (!currentAttendanceSheet) {
    console.warn("対象月の勤務実績表を取得できませんでした。ファイル名:" + targetSpreadSheet.getName())
    return false;
  }
  console.log("処理対象シート：" + currentAttendanceSheet.getName())
  var staffId = currentAttendanceSheet.getRange(STAFF_ID_RANGE_POSITION).getValue();

  var regex = new RegExp(/^S[0-9]{4}$/);
  if (typeof (staffId) !== "string" || !regex.test(staffId)) {
    console.log("スタッフIDが検知できないか、命名規則に沿っていません。ファイル名：" + targetSpreadSheet.getName() + ", スタッフID：" + staffId);
    return false;
  }

  var sheetName = currentAttendanceSheet.getName();
  var yearStr = sheetName.split("年")[0];
  var monthStr = sheetName.split("年")[1].split("月")[0];
  var startDate = new Date(yearStr,monthStr-1);
  var endDate = new Date(yearStr,monthStr);
  endDate.setDate(0);
  console.log(startDate+" ~ "+endDate);
  var dateRange = currentAttendanceSheet.getRange("K7:K90");
  console.log("dateRange.getValues().length"+dateRange.getValues().length);
  for(var k=0;k<dateRange.getValues().length;k++){
    var value  = dateRange.getValues()[k][0];
    if(value === "合計金額(税込)"){
      break;
    }
    if(value===""){
      continue;
    }
    var tranDate = new Date(value);
    if(startDate<=tranDate && tranDate <= endDate){
      console.log("OK");
    }else{
      console.log("該当月に含まれていない日時が存在します。ファイル名：" + targetSpreadSheet.getName() + ", スタッフID：" + staffId);
      return false;
    }
  }
  console.log(dateRange.getValues());

  


  var total = currentAttendanceSheet.getRange(TOTAL_RANGE_POSITION).getValue();
  if (typeof (total) !== "number" || total <= 0) {
    console.log("金額カラムが不正です。ファイル名：" + targetSpreadSheet.getName() + ", 金額：" + total);
    return false;
  }

// var TOTAL_CHECK_RANGE_POSITION = "I24";
  var totalCheckStatus = currentAttendanceSheet.getRange("I24").getValue();
  if (totalCheckStatus !== CHECK_OK_TEXT) {
    console.log("金額チェックNG。ファイル名：" + targetSpreadSheet.getName() + ", チェック結果：" + totalCheckStatus);
    return false;
  }

// var FIXED_TOTAL_RANGE_POSITION = "I25";
  currentAttendanceSheet.getRange("I25").setValue(total);

  protect(currentAttendanceSheet);
  return true;

}



function askExecutable() {
  var ui = SpreadsheetApp.getUi();
  var title = '請求額の確定';
  var prompt = '請求額を確定しますか？\n確定した場合、シートが保護され編集ができなくなります。'
  var response = ui.alert(title, prompt, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    return true;
  } else {
    var msg = "処理をキャンセルしました。"
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'キャンセル', 5);
    return false
  }

}

function showResultMessage(result) {
  if (result) {
    var msg = "シートを保護しました。修正したい際には管理者までお問い合わせください。";
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '確定処理成功', 7);
  } else {
    var msg = "エラーのため確定処理を中止しました。管理者までお問い合わせお願いします。";
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '確定エラー', 7);
  }

}

function protect(currentAttendanceSheet) {
  var protection = currentAttendanceSheet.protect();
  protection.setDescription(PROTECTION_DESCRIPTION);
  protection.setWarningOnly(true);

// var FIXED_MESSAGE_POSITION = "E26";
  var messageRange = currentAttendanceSheet.getRange("E26");
  messageRange.setValue(FIXED_MESSAGE);
  messageRange.setFontColor("red");
  messageRange.setFontWeight("bold");
  messageRange.setFontSize(14);
  messageRange.setHorizontalAlignment("right");

  console.log("sheetを保護しました。");
}