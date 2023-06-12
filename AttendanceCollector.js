//すべての勤務実績表をチェックし、集計するツール
//★ 手動で集計を実行するときはこれを実行する！！
function attendanceCollectTrigger() {
  collectMonthlyAttendanceSummary();
}

// 収集処理
function collectMonthlyAttendanceSummary() {
  console.time("TOTAL TIME");
  var start_time = new Date();

  prepareConfig(CURRENT_CNT_PROPERTY_KEY);

  delete_specific_triggers("collectMonthlyAttendanceSummary");

  var staffSpreadSheetList = getSortedStaffAttendanceSheetList();

  //各スタッフのスプシ毎の処理
  for (var cnt = countProperty; cnt < staffSpreadSheetList.length; cnt++) {

    var isSucceeded = false;
    var file = staffSpreadSheetList[cnt];
    //    sheetCnt++;
    var prevMonthTitle = getPrevMonthTitle();

    var currentAttendanceSheet = SpreadsheetApp.open(file).getSheetByName(prevMonthTitle);
    if (!currentAttendanceSheet) {
      // notFindSheetCnt++;
      console.log("先月分の勤務シートが見つかりませんでした。処理をスキップします。:" + file.getName())
      continue;
    }

    var staffId = currentAttendanceSheet.getRange(STAFF_ID_RANGE_POSITION).getValue();

    var regex = new RegExp(/^S[0-9]{4}$/);
    if (typeof (staffId) !== "string" || !regex.test(staffId)) {
      // otherError++;
      console.warn("スタッフIDが検知できないか、命名規則に沿っていません。ファイル名：" + file.getName() + ", スタッフID：" + staffId);
      continue;
    }

    //var isFixed = currentAttendanceSheet.getRange(FIXED_STATUS_RANGE_POSITION).getValue();
    var protections = currentAttendanceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var isFixed = false;

    var targetProtection;
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.getDescription() === PROTECTION_DESCRIPTION) {
        isFixed = true;
        targetProtection = protection;
        break;
      }
    }


    if (!isFixed) {
      console.log("先月分の勤務シートが未確定です。処理をスキップします。:" + file.getName());
      // unfixedCnt++;
      continue;
    }

    var total = currentAttendanceSheet.getRange(TOTAL_RANGE_POSITION).getValue();
    if (typeof (total) !== "number" || total <= 0) {
      console.warn("金額が未入力か、金額カラムが不正です。ファイル名：" + file.getName() + ", 金額：" + total);
      // otherError++;
      continue;
    }

    var fixedTotal = currentAttendanceSheet.getRange(FIXED_TOTAL_RANGE_POSITION).getValue();
    if (total !== fixedTotal) {
      console.warn("請求確定金額が勤務実績表の金額と一致しません。ファイル名：" + file.getName() + ", 金額：" + total);
      // unmatchedCnt++;
      continue;

    }

    isSucceeded = setTotalToSummarySheet(attendanceSummarySheetId, prevMonthTitle, staffId, total);

    if (isSucceeded) {
      targetProtection.remove();
      var protection = currentAttendanceSheet.protect();
      protection.setDescription(PROTECTION_DESCRIPTION);
      protection.removeEditors(protection.getEditors());
      protection.addEditors(EDITORS_LIST);

      console.log("集計が成功し、保護をかけました。ファイル名：" + file.getName());
      // succeededCnt++;
      correctEnquete(staffId,file.getName(),currentAttendanceSheet);
    }
    if (needSuspend) {
      break;
    }
    if (needRestart(start_time, CURRENT_CNT_PROPERTY_KEY, cnt)) {
      ScriptApp
        .newTrigger("collectMonthlyAttendanceSummary")
        .timeBased()
        .everyMinutes(1)
        .create();
      console.log("6 minutes restart!!");
      return;
    }
  }
  initializeProperies(CURRENT_CNT_PROPERTY_KEY)
  //特定関数のトリガーのみ削除
  delete_specific_triggers("collectMonthlyAttendanceSummary");

  if (needSuspend) {
    console.warn("処理が中断されました。");
    sendSlackSuspendMessage();
  }

  //outputEndSummaryLog();

  console.timeEnd("TOTAL TIME");
}


/* 
* シート内の特定の列内の文字列を検索する。便利。
* @param <String> attendanceSummarySheetId 「銀行振込の振込み先口座（回答）在宅スタッフの勤務実績の金額」シートのID。テストの場合にIDを変えたいので引数で渡す
* @param <String> prevMonthTitle 対象となるシートのID
* @param <String> staffId スタッフのID
* @param <String> total 金額
* 
* @return {boolean} 成功可否
*/
function setTotalToSummarySheet(attendanceSummarySheetId, prevMonthTitle, staffId, total) {
  var attendaceSummarySpreadSheet = SpreadsheetApp.openById(attendanceSummarySheetId);
  var prevMonthSheet = attendaceSummarySpreadSheet.getSheetByName(prevMonthTitle);
  if (!prevMonthSheet) {
    console.warn("「在宅スタッフの勤務実績の金額」に先月のシートがありません。検索対象ファイル名：" + prevMonthTitle)
    needSuspend = true;
    // otherError++;
    return false;
  }
  var targetRowIdx = findRow(prevMonthSheet, staffId, STAFF_ID_COL_INDEX);
  if (targetRowIdx <= 0) {
    console.warn("「在宅スタッフの勤務実績の金額」に対象の行が見つかりませんでした。StaffId : " + staffId);
    // notFindStaffIdCnt++;
    return false;
  }
  if (total === prevMonthSheet.getRange(TOTAL_COL_POSITION + targetRowIdx).getValue()) {
    console.log("集計済みファイルです。StaffId : " + staffId);
    return false;
  }
  prevMonthSheet.getRange(TOTAL_COL_POSITION + targetRowIdx).setValue(total);
  return true;
}

/* 
* シート内の特定の列内の文字列を検索する。便利。
* @param <String> attendanceSummarySheetId 「銀行振込の振込み先口座（回答）在宅スタッフの勤務実績の金額」シートのID。テストの場合にIDを変えたいので引数で渡す
* @param <String> prevMonthTitle 対象となるシートのID
* @param <String> staffId スタッフのID
* @param <String> total 金額
* 
* @return {boolean} クリア実施有無
*/
function clearTotalToSummarySheet(attendanceSummarySheetId, prevMonthTitle, staffId) {
  var attendaceSummarySpreadSheet = SpreadsheetApp.openById(attendanceSummarySheetId);
  var prevMonthSheet = attendaceSummarySpreadSheet.getSheetByName(prevMonthTitle);
  if (!prevMonthSheet) {
    console.warn("「在宅スタッフの勤務実績の金額」に先月のシートがありません。検索対象ファイル名：" + prevMonthTitle)
    needSuspend = true;
    return false;
  }
  var targetRowIdx = findRow(prevMonthSheet, staffId, STAFF_ID_COL_INDEX);
  if (targetRowIdx <= 0) {
    console.warn("対象の行が見つかりませんでした。StaffId : " + staffId);
    return false;
  }
  var targetAmountRange = prevMonthSheet.getRange(TOTAL_COL_POSITION + targetRowIdx);
  if (typeof (targetAmountRange.getValue()) === "number" && targetAmountRange.getValue() > 0) {
    console.log("金額をクリアします。StaffId : " + staffId);
    targetAmountRange.clear();
    return true;
  }
  return false;
}

function correctEnquete(staffId,fileName,currentAttendanceSheet){
  var hasSixthEnquete = (currentAttendanceSheet.getRange(ENQUETE_LAST_RANGE_POSITION).getValue() == ENQUETE_LAST_TITLE);
  
  var cm_sheet = SpreadsheetApp.openById(CM_SHEET_ID).getSheetByName(CM_ENQUETE_SHEET_NAME);
  var firstEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIRST_RANGE_POSITION).getValue();
  var secondEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SECOND_RANGE_POSITION).getValue();
  var thirdEnqueteAnswere = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_THIRD_RANGE_POSITION).getValue();
  if(hasSixthEnquete){
    var fourthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FOURTH_RANGE_POSITION).getValue();
    var fifthEnqueteAnsewer = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_FIFTH_RANGE_POSITION).getValue();
    var sixthEnqueteAnswere = currentAttendanceSheet.getRange(ENQUETE_ATTENDANCE_SIXTH_RANGE_POSITION).getValue();
    cm_sheet.appendRow([staffId, fileName, currentAttendanceSheet.getName(), firstEnqueteAnsewer, secondEnqueteAnsewer, thirdEnqueteAnswere,fourthEnqueteAnsewer,fifthEnqueteAnsewer,sixthEnqueteAnswere]);
  }else{
    cm_sheet.appendRow([staffId, fileName, currentAttendanceSheet.getName(), firstEnqueteAnsewer, secondEnqueteAnsewer, thirdEnqueteAnswere]);
  }
}