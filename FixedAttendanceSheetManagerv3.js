function submitIncludeEnquete() {
  console.time("TOTAL EXECUTION TIME");
  var canExecute = askExecutable();
  if (canExecute) {
    var result = submitAttendancev3();
    showResultMessage(result);
  }
  console.timeEnd("TOTAL EXECUTION TIME");

}


// 収集処理
function submitAttendancev3() {

  var targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  console.log("■■ 処理対象ファイル名：" + targetSpreadSheet.getName());

  var currentAttendanceSheet = targetSpreadSheet.getActiveSheet();
  if (!currentAttendanceSheet) {
    console.warn("対象月の勤務実績表を取得できませんでした。ファイル名:" + targetSpreadSheet.getName())
    return false;
  }
  console.log("処理対象シート：" + currentAttendanceSheet.getName())
  
  var enqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIRST_RANGE_POSITION).getValue();
  console.log(ENQUETE_FIRST_ANSWER_LIST);
  // if(!ENQUETE_FIRST_ANSWER_LIST.includes(enqueteAnsewer)){
  if(ENQUETE_FIRST_ANSWER_LIST.indexOf(enqueteAnsewer) == -1){
    console.warn("①アンケート未回答");
    errorMessage = "『稼働アンケート』が未回答です。お手数をおかけしますが、先に稼働アンケートにご回答ください。";
    return false;
  }
  var secondEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SECOND_RANGE_POSITION).getValue();
  if(enqueteAnsewer=="1.増やしたい ↑"){
    if(secondEnqueteAnsewer==""){
      console.warn("②アンケート未回答");
      errorMessage = "稼働アンケートの②が未回答です。①で「1.増やしたい ↑」を選択の場合には必ずご記入ください。";
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

  // var totalCheckStatus = currentAttendanceSheet.getRange(TOTAL_CHECK_RANGE_POSITION).getValue();
  var totalCheckStatus = currentAttendanceSheet.getRange("I31").getValue();
  if (totalCheckStatus !== CHECK_OK_TEXT) {
    console.log("金額チェックNG。ファイル名：" + targetSpreadSheet.getName() + ", チェック結果：" + totalCheckStatus);
    return false;
  }

  // currentAttendanceSheet.getRange(FIXED_TOTAL_RANGE_POSITION).setValue(total);
  currentAttendanceSheet.getRange("I32").setValue(total);

  protect(currentAttendanceSheet,"E33");

  return true;

}
