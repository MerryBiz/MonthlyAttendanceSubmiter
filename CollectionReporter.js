//★ 手動でレポーティングを実行するときはこれを実行する！！
function reportingTrigger() {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty("sendSlack", true);
  reportingCollectSituation();
  console.log("end of reporting");
}

function reportingTriggerWithoutSlack() {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty("sendSlack", false);
  reportingCollectSituation();
  console.log("end of reporting");
}

// 収集処理
// Reporting用のfunction。これを実行したらSlackに連携する！
function reportingCollectSituation() {
    console.time("REPORTING TOTAL TIME");
    var start_time = new Date();

    prepareConfig(REPORTING_CNT_PROPERTY_KEY);
    delete_specific_triggers("reportingCollectSituation");

    var staffSpreadSheetList = getSortedStaffAttendanceSheetList();
    var properties = PropertiesService.getScriptProperties();
    var sendSlackFlg = JSON.parse(properties.getProperty("sendSlack"));

    //各スタッフのスプシ毎の処理
    for (var cnt = countProperty; cnt < staffSpreadSheetList.length; cnt++) {

        var file = staffSpreadSheetList[cnt];
        Logger.log("start for " + file.getName());
        sheetCnt++;
        var prevMonthTitle = getPrevMonthTitle();

        var currentAttendanceSheet = SpreadsheetApp.open(file).getSheetByName(prevMonthTitle);
        if (!currentAttendanceSheet) {
            console.log("先月分の勤務シートが見つかりませんでした。処理をスキップします。:" + file.getName())
            notFindSheetCnt++;
            continue;
        }

        var staffId = currentAttendanceSheet.getRange(STAFF_ID_RANGE_POSITION).getValue();

        var regex = new RegExp(/^S[0-9]{4}$/);
        if (typeof (staffId) !== "string" || !regex.test(staffId)) {
            console.warn("スタッフIDが検知できないか、命名規則に沿っていません。ファイル名：" + file.getName() + ", スタッフID：" + staffId);
            otherError++;
            continue;
        }

        var total = currentAttendanceSheet.getRange(TOTAL_RANGE_POSITION).getValue();
        if (typeof (total) !== "number" || total < 0) {
            console.warn("金額カラムが不正です。ファイル名：" + file.getName() + ", 金額：" + total);
            otherError++;
            continue;
        } else if (total === 0) {
            console.info("実績が未入力で、金額が0円です。ファイル名：" + file.getName() + ", 金額：" + total);
            notInputedAmountCnt++;
            continue;
        }

        var hasProtection = hasProtectionsCheck(currentAttendanceSheet);

        var fixedTotal = currentAttendanceSheet.getRange(FIXED_TOTAL_RANGE_POSITION).getValue();
        if (hasProtection && fixedTotal !== total) {
            console.warn("請求確定金額と勤務実績表の集計結果が不一致です。ファイル名：" + file.getName() + ", 金額：" + total);
            unmatchedCnt++;
            continue;
        }


        countExecValue(attendanceSummarySheetId, prevMonthTitle, staffId, hasProtection, file.getName(), fixedTotal,sendSlackFlg);
        if (needSuspend) {
            break;
        }
        if (needRestart(start_time, REPORTING_CNT_PROPERTY_KEY, cnt)) {
            ScriptApp
                .newTrigger("reportingCollectSituation")
                .timeBased()
                .everyMinutes(1)
                .create();
            console.log("6 minutes restart!!");

            return;
        }
    }
    initializeProperies(REPORTING_CNT_PROPERTY_KEY)

    //特定関数のトリガーのみ削除
    delete_specific_triggers("reportingCollectSituation");

    if (needSuspend) {
        sendSlackSuspendMessage();
    } else {
        sendSlackReportingMessage();
    }

    console.log("finish");

    console.timeEnd("REPORTING TOTAL TIME");

}


/* 
* シート内の特定の列内の文字列を検索する。便利。
* @param <String> attendanceSummarySheetId 「銀行振込の振込み先口座（回答）在宅スタッフの勤務実績の金額」シートのID。テストの場合にIDを変えたいので引数で渡す
* @param <String> prevMonthTitle 対象となるシートのID
* @param <String> staffId スタッフのID
* @param <String> total 金額
* @param <boolean> sendSlackFlg 未確定者の名前をSlackに流すかどうか
* 
* @return {boolean} 成功可否
*/
function countExecValue(attendanceSummarySheetId, prevMonthTitle, staffId, hasProtection, fileName, fixedTotal,sendSlackFlg) {
    var attendaceSummarySpreadSheet = SpreadsheetApp.openById(attendanceSummarySheetId);
    var prevMonthSheet = attendaceSummarySpreadSheet.getSheetByName(prevMonthTitle);
    if (!prevMonthSheet) {
        console.warn("「在宅スタッフの勤務実績の金額」に先月のシートがありません。検索対象ファイル名：" + prevMonthTitle)
        otherError++;
        needSuspend = true;
        return;
    }
    var targetRowIdx = findRow(prevMonthSheet, staffId, STAFF_ID_COL_INDEX);
    if (targetRowIdx <= 0) {
        console.warn("「在宅スタッフの勤務実績の金額」対象の行が見つかりませんでした。StaffId : " + staffId);
        notFindStaffIdCnt++;
        return;
    }

    var targetAmountRange = prevMonthSheet.getRange(TOTAL_COL_POSITION + targetRowIdx);
    if (typeof (targetAmountRange.getValue()) === "number" && targetAmountRange.getValue() > 0) {
        if (hasProtection) {
            if (targetAmountRange.getValue() === fixedTotal) {
                succeededCnt++;
            } else {
                console.warn("勤務表の金額と「在宅スタッフの勤務実績の金額」が一致しません。ファイル名:" + fileName)
                needCheckSSNamesList.push(fileName);
                otherError++;
            }
        } else {
            console.warn("保護されていませんが、金額が集計されています。ファイル名:" + fileName)
            succeededWithoutProtectionCnt++;
            needCheckSSNamesList.push(fileName);
        }
        return;
    } else {
        if (hasProtection) {
            unexecutedCnt++;
        } else {
            console.log("先月分の勤務シートが未確定です。処理をスキップします。:" + fileName);
            if(sendSlackFlg){
              send("【先月分の勤務シートが未確定です】"+fileName);
            }
            unfixedCnt++;
        }
        return;
    }
}

function hasProtectionsCheck(currentAttendanceSheet) {
    var protections = currentAttendanceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    for (var i = 0; i < protections.length; i++) {
        var protection = protections[i];
        Logger.log(protection.getDescription())
        if (protection.getDescription() === PROTECTION_DESCRIPTION) {
            return true;
        }
    }
    return false;
}
