function submit() {
  var targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  var currentAttendanceSheet = targetSpreadSheet.getActiveSheet();
  if (!currentAttendanceSheet) {
    console.warn("対象月の勤務実績表を取得できませんでした。ファイル名:" + targetSpreadSheet.getName());
    return false;
  }
  
  var protections = currentAttendanceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === PROTECTION_DESCRIPTION) {
      console.log("確定処理済みのためsubmit処理はスキップ");
      return false;
    }
  }

  var hasEnquete = (currentAttendanceSheet.getRange(ENQUETE_TITLE_RANGE_POSITION).getValue() == "稼働アンケート");

  if (hasEnquete) {
    var hasSixthEnquete = (currentAttendanceSheet.getRange(ENQUETE_LAST_RANGE_POSITION).getValue() == ENQUETE_LAST_TITLE);
    var hasEighthEnquete = (currentAttendanceSheet.getRange(ENQUETE_LAST_RANGE_POSITION_V2).getValue().startsWith("⑤【④で「2.減る可能性がある」「3.増える可能性がある」を選択した方】"));
    if(hasEighthEnquete){
      submitIncludeEnquetev5();
  } else if (hasSixthEnquete) {
    
      submitIncludeEnqueteV2();
    }
  } else {
    submitWithoutEnquete();
  }
}
