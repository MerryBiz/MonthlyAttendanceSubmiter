//アンケートの質問項目追加対応
function submitIncludeEnquetev5() {
  console.time("TOTAL EXECUTION TIME");
  var canExecute = askExecutable();
  if (canExecute) {
    var result = submitAttendancev5();
    showResultMessage(result);
    //PDF送付処理
  }
  console.timeEnd("TOTAL EXECUTION TIME");

}


// 収集処理
function submitAttendancev5() {

  var targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  console.log("■■ 処理対象ファイル名：" + targetSpreadSheet.getName());

  var currentAttendanceSheet = targetSpreadSheet.getActiveSheet();
  if (!currentAttendanceSheet) {
    console.warn("対象月の勤務実績表を取得できませんでした。ファイル名:" + targetSpreadSheet.getName())
    return false;
  }
  console.log("処理対象シート：" + currentAttendanceSheet.getName())

  var firstEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIRST_RANGE_POSITION).getValue();
  var seventhEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SEVENTH_RANGE_POSITION).getValue();
  // if (!ENQUETE_FIRST_ANSWER_LIST.includes(firstEnqueteAnsewer)) {
  if(ENQUETE_FIRST_ANSWER_LIST_V2.indexOf(firstEnqueteAnsewer) == -1){
    console.warn("①アンケート未回答");
    errorMessage = "『稼働アンケート』の必須回答が未回答です。お手数をおかけしますが、先に稼働アンケートにご回答ください。";
    return false;
  }
  // if (!ENQUETE_FOURTH_ANSWER_LIST.includes(seventhEnqueteAnsewer)) {
  if(ENQUETE_SEVENTH_ANSWER_LIST_V2.indexOf(seventhEnqueteAnsewer) == -1){
    console.warn("④アンケート未回答");
    errorMessage = "『稼働アンケート』の必須回答が未回答です。お手数をおかけしますが、先に稼働アンケートにご回答ください。";
    return false;
  }
  var secondEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SECOND_RANGE_POSITION).getValue();
  if (firstEnqueteAnsewer == "4.減らしたい") {
    if (secondEnqueteAnsewer == "") {
      console.warn("②アンケート未回答");
      errorMessage = "稼働アンケートの②が未回答です。①で「4.減らしたい」を選択の場合には必ずご記入ください。"
      return false;
    }
  }
  var thirdEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_THIRD_RANGE_POSITION).getValue();
  if (firstEnqueteAnsewer == "1.積極的に追加したい" ||firstEnqueteAnsewer == "2.条件によっては追加可能" ) {
    if (thirdEnqueteAnsewer == "") {
      console.warn("③-1アンケート未回答");
      errorMessage = "稼働アンケートの③-1が未回答です。①で「1.積極的に追加したい」か「2.条件によっては追加可能」を選択の場合には必ずご記入ください。"
      return false;
    }
  }
  var forthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FOURTH_RANGE_POSITION).getValue();
  if (firstEnqueteAnsewer == "1.積極的に追加したい" ||firstEnqueteAnsewer == "2.条件によっては追加可能" ) {
    if (forthEnqueteAnsewer == "") {
      console.warn("③-2アンケート未回答");
      errorMessage = "稼働アンケートの③-2が未回答です。①で「1.積極的に追加したい」か「2.条件によっては追加可能」を選択の場合には必ずご記入ください。"
      return false;
    }
  }  
  var fifthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIFTH_RANGE_POSITION).getValue();
  if (firstEnqueteAnsewer == "1.積極的に追加したい" ||firstEnqueteAnsewer == "2.条件によっては追加可能" ) {
    if (fifthEnqueteAnsewer == "") {
      console.warn("③-3アンケート未回答");
      errorMessage = "稼働アンケートの③-3が未回答です。①で「1.積極的に追加したい」か「2.条件によっては追加可能」を選択の場合には必ずご記入ください。"
      return false;
    }
  }

  var eighthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_EIGHTH_RANGE_POSITION).getValue();
  if (seventhEnqueteAnsewer == "2. 減る予定・可能性がある" ||seventhEnqueteAnsewer == "3. 増える予定・可能性がある" ) {
    if (eighthEnqueteAnsewer == "") {
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
        console.log(value);
    if (value === "小計" || value === "合計金額(税込)") {
      break;
    }
    if (value === "") {
      continue;
    }
    var tranDate = new Date(value);
        console.log(tranDate); // 後で消す
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
  // var totalCheckStatus = currentAttendanceSheet.getRange("I41").getValue();
  if (totalCheckStatus !== CHECK_OK_TEXT) {
    console.log("金額チェックNG。ファイル名：" + targetSpreadSheet.getName() + ", チェック結果：" + totalCheckStatus);
    return false;
  }

  currentAttendanceSheet.getRange(FIXED_TOTAL_RANGE_POSITION).setValue(total);
  // currentAttendanceSheet.getRange("I42").setValue(total);

  protect(currentAttendanceSheet,FIXED_MESSAGE_POSITION);
  // protect(currentAttendanceSheet,"E43");

  return true;

}
