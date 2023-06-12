//アンケートの質問項目追加対応
function submitIncludeEnqueteV2() {
  console.time("TOTAL EXECUTION TIME");
  var canExecute = askExecutable();
  if (canExecute) {
    var result = submitAttendancev4();
    showResultMessage(result);
    //PDF送付処理
  }
  console.timeEnd("TOTAL EXECUTION TIME");

}


// 収集処理
function submitAttendancev4() {

  var targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  console.log("■■ 処理対象ファイル名：" + targetSpreadSheet.getName());

  var currentAttendanceSheet = targetSpreadSheet.getActiveSheet();
  if (!currentAttendanceSheet) {
    console.warn("対象月の勤務実績表を取得できませんでした。ファイル名:" + targetSpreadSheet.getName())
    return false;
  }
  console.log("処理対象シート：" + currentAttendanceSheet.getName())

  var firstEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIRST_RANGE_POSITION).getValue();
  var fourthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FOURTH_RANGE_POSITION).getValue();
  // if (!ENQUETE_FIRST_ANSWER_LIST.includes(firstEnqueteAnsewer)) {
  if(ENQUETE_FIRST_ANSWER_LIST.indexOf(firstEnqueteAnsewer) == -1){
    console.warn("①アンケート未回答");
    errorMessage = "『稼働アンケート』の必須回答が未回答です。お手数をおかけしますが、先に稼働アンケートにご回答ください。";
    return false;
  }
  // if (!ENQUETE_FOURTH_ANSWER_LIST.includes(fourthEnqueteAnsewer)) {
  if(ENQUETE_FOURTH_ANSWER_LIST.indexOf(fourthEnqueteAnsewer) == -1){
    console.warn("④アンケート未回答");
    errorMessage = "『稼働アンケート』の必須回答が未回答です。お手数をおかけしますが、先に稼働アンケートにご回答ください。";
    return false;
  }
  var secondEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SECOND_RANGE_POSITION).getValue();
  if (firstEnqueteAnsewer == "1.増やしたい ↑") {
    if (secondEnqueteAnsewer == "") {
      console.warn("②アンケート未回答");
      errorMessage = "稼働アンケートの②が未回答です。①で「1.増やしたい ↑」を選択の場合には必ずご記入ください。"
      return false;
    }
  }
  var fifthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIFTH_RANGE_POSITION).getValue();
  if (fourthEnqueteAnsewer == "減る予定・可能性がある" ||fourthEnqueteAnsewer == "増える予定・可能性がある" ) {
    if (fifthEnqueteAnsewer == "") {
      console.warn("⑤アンケート未回答");
      errorMessage = "稼働アンケートの⑤が未回答です。④で「変わる可能性がある」と回答された場合には必ずご記入ください。"
      return false;
    }
  }
  var staffId = currentAttendanceSheet.getRange(STAFF_ID_RANGE_POSITION).getValue();

  var regex = new RegExp(/^S[0-9]{4}$/);
  if (typeof (staffId) !== "string" || !regex.test(staffId)) {
    console.log("スタッフIDが検知できないか、命名規則に沿っていません。ファイル名：" + targetSpreadSheet.getName() + ", スタッフID：" + staffId);
    return false;
  }

  var sheetName = currentAttendanceSheet.getName();
  var yearStr = sheetName.split("年")[0];
  var monthStr = sheetName.split("年")[1].split("月")[0];
  var startDate = new Date(yearStr, monthStr - 1);
  var endDate = new Date(yearStr, monthStr);
  endDate.setDate(0);
  console.log(startDate + " ~ " + endDate);
  var dateRange = currentAttendanceSheet.getRange("K7:K90");
  console.log("dateRange.getValues().length" + dateRange.getValues().length);
  for (var k = 0; k < dateRange.getValues().length; k++) {
    var value = dateRange.getValues()[k][0];
    if (value === "合計金額(税込)") {
      break;
    }
    if (value === "") {
      continue;
    }
    var tranDate = new Date(value);
    if (startDate <= tranDate && tranDate <= endDate) {
      console.log("OK");
    } else {
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

  var totalCheckStatus = currentAttendanceSheet.getRange(TOTAL_CHECK_RANGE_POSITION).getValue();
  // var totalCheckStatus = currentAttendanceSheet.getRange("I37").getValue();
  if (totalCheckStatus !== CHECK_OK_TEXT) {
    console.log("金額チェックNG。ファイル名：" + targetSpreadSheet.getName() + ", チェック結果：" + totalCheckStatus);
    return false;
  }

  currentAttendanceSheet.getRange(FIXED_TOTAL_RANGE_POSITION).setValue(total);
  // currentAttendanceSheet.getRange("I38").setValue(total);

  protectv4(currentAttendanceSheet);

  //集計時に移行
  // var cm_sheet = SpreadsheetApp.openById(CM_SHEET_ID).getSheetByName(CM_ENQUETE_SHEET_NAME);
  // var thirdEnqueteAnswere = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_THIRD_RANGE_POSITION).getValue();
  // var sixthEnqueteAnswere = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SIXTH_RANGE_POSITION).getValue();
  // cm_sheet.appendRow([staffId, targetSpreadSheet.getName(), currentAttendanceSheet.getName(), firstEnqueteAnsewer, secondEnqueteAnsewer, thirdEnqueteAnswere,fourthEnqueteAnsewer,fifthEnqueteAnsewer,sixthEnqueteAnswere]);

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
    var msg;
    if (errorMessage) {
      msg = errorMessage;
    } else {
      msg = "エラーのため確定処理を中止しました。管理者までお問い合わせお願いします。";
    }
    Browser.msgBox(msg)
    //    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '確定エラー', 7);
  }

}

function protectv4(currentAttendanceSheet) {
  var protection = currentAttendanceSheet.protect();
  protection.setDescription(PROTECTION_DESCRIPTION);
  protection.setWarningOnly(true);

  var messageRange = currentAttendanceSheet.getRange(FIXED_MESSAGE_POSITION);
  // var messageRange = currentAttendanceSheet.getRange("E39");
  messageRange.setValue(FIXED_MESSAGE);
  messageRange.setFontColor("red");
  messageRange.setFontWeight("bold");
  messageRange.setFontSize(14);
  messageRange.setHorizontalAlignment("right");

  console.log("sheetを保護しました。");
}