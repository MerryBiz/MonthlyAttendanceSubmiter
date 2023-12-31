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
    if (hasSixthEnquete) {
      submitIncludeEnqueteV2();
    } else {
      submitIncludeEnquete();
    }
  } else {
    submitWithoutEnquete();
  }

  console.log("hasEnquete:" + hasEnquete);
}
