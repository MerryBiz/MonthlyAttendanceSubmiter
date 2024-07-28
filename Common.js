var scriptProperties = PropertiesService.getScriptProperties();

var ATTENDANCE_FOLDER_ID = scriptProperties.getProperty('ATTENDANCE_FOLDER_ID');
var ATTENDANCE_SUMMARY_SHEET_ID = scriptProperties.getProperty('ATTENDANCE_SUMMARY_SHEET_ID');
var CM_SHEET_ID = scriptProperties.getProperty('CM_SHEET_ID');
var CM_ENQUETE_SHEET_NAME = "稼働アンケート";
var ATTENDANCE_FOLDER_ID_FOR_TEST = scriptProperties.getProperty('ATTENDANCE_FOLDER_ID_FOR_TEST'); //検証テスト
var ATTENDANCE_SUMMARY_SHEET_ID_FOR_TEST = scriptProperties.getProperty('ATTENDANCE_SUMMARY_SHEET_ID_FOR_TEST');
var EDITORS_LIST = scriptProperties.getProperty('EDITORS_LIST').split(",");

var ENQUETE_FIRST_ANSWER_LIST = new Array('1.増やしたい ↑','2.今のペースで働きたい →','3.減らしたい ↓');
var ENQUETE_FOURTH_ANSWER_LIST = new Array('今のところ、変わる予定はない','減る予定・可能性がある','増える予定・可能性がある');

var ENQUETE_FIRST_ANSWER_LIST_V2 = new Array('1.積極的に追加したい','2.条件によっては追加可能','3.現状維持希望','4.減らしたい');
var ENQUETE_SEVENTH_ANSWER_LIST_V2 = new Array('1. 変わる予定はない','2. 減る予定・可能性がある','3. 増える予定・可能性がある');


// 勤務実績表のセル位置..
var FIXED_STATUS_RANGE_POSITION = "F24";
var STAFF_ID_RANGE_POSITION = "R2";
var TOTAL_RANGE_POSITION = "I22";
var TOTAL_CHECK_RANGE_POSITION = "I41";
var FIXED_MESSAGE_POSITION = "E43";
var FIXED_TOTAL_RANGE_POSITION = "I42";
var ENQUETE_ATTENDANCE_FIRST_RANGE_POSITION = "C25";
var ENQUETE_ATTENDANCE_SECOND_RANGE_POSITION = "C27";
var ENQUETE_ATTENDANCE_THIRD_RANGE_POSITION = "C29";
var ENQUETE_ATTENDANCE_FOURTH_RANGE_POSITION = "C31";
var ENQUETE_ATTENDANCE_FIFTH_RANGE_POSITION = "C33";
var ENQUETE_ATTENDANCE_SIXTH_RANGE_POSITION = "C35";
var ENQUETE_ATTENDANCE_SEVENTH_RANGE_POSITION = "C37";
var ENQUETE_ATTENDANCE_EIGHTH_RANGE_POSITION = "C39";


var ENQUETE_TITLE_RANGE_POSITION = "C23";
var ENQUETE_LAST_RANGE_POSITION = "C34";
var ENQUETE_LAST_RANGE_POSITION_V2 = "C38";
var ENQUETE_LAST_TITLE = "⑥その他、補足等あれば自由にご記入ください。 *任意";

var CHECK_OK_TEXT = "OK";

// 在宅スタッフの勤務実績の金額リストの列位置
var STAFF_ID_COL_INDEX = 2;
var TOTAL_COL_POSITION = "E";

var PROTECTION_DESCRIPTION = "勤務実績表確定による保護";
var FIXED_MESSAGE = "※請求額確定済み";

//カウンター
var sheetCnt = 0; // 総シート数
var succeededCnt = 0; // 成功したシート数
var unfixedCnt = 0; // 未確定シート数
var unmatchedCnt = 0;
var unexecutedCnt = 0; // 処理待ち件数
var succeededWithoutProtectionCnt = 0; // 保護されてないのに集計されているシート数
var notFindSheetCnt = 0; // 先月分が見つからなかったシート数
var notFindStaffIdCnt = 0; // スタッフIDが一覧シートで見つからなかった数
var notInputedAmountCnt = 0;//金額未入力シート数
var otherError = 0; // レイアウト崩れ等

var countProperty = 0;

var needCheckSSNamesList = [];

var attendanceFolderId;
var attendanceSummarySheetId;

var isTest = false; // Test flag. If you want to test, please set it to "true".
var needSuspend = false; // 処理を中断する際のフラグ

var CURRENT_CNT_PROPERTY_KEY = "CURRENT_CNT";
var REPORTING_CNT_PROPERTY_KEY = "REPORTING_CNT"

var errorMessage;




function getPrevMonthTitle() {
    var now = new Date();
    now.setMonth(now.getMonth() - 1);
    var prevMonthTitle = now.getFullYear() + "年" + (now.getMonth() + 1) + "月";
    return prevMonthTitle;
}

// スタッフのシートファイルをフォルダから取得する
function getSortedStaffAttendanceSheetList() {
    console.time("sortTime");

    var targetFolder = DriveApp.getFolderById(attendanceFolderId);


    Logger.log(targetFolder.getName());
    var files = targetFolder.searchFiles("title contains '勤務実績表'");

    //各スタッフのスプシ毎の処理
    var filesArray = [];

    //検証用のファイル制限
    // TODO Need to remove.
    // var verificationFileNameList = ["S0003_松尾 綾子様_勤務実績表", "S0006_皆見 佳子様_勤務実績表", "S0018_原田 雅美 様_勤務実績表", "S0021_近藤 昌代様_勤務実績表", "S0004_吉益 美江様_勤務実績表"];

    while (files.hasNext()) {
        var file = files.next();
        // for (var k = 0; k < verificationFileNameList.length; k++) {
        //     if (file.getName() === verificationFileNameList[k]) {
        filesArray.push(file)
        //         break;
        //     }
        // }
    }
    filesArray.sort(function (a, b) {
        if (a.getName() > b.getName()) {
            return 1;
        } else {
            return -1;
        }
    });

    // TODO Need to remove.
    for (var i = 0; i < filesArray.length; i++) {
        Logger.log(filesArray[i].getName());
    }
    console.timeEnd("sortTime");

    return filesArray;
}


/* 
* シート内の特定の列内の文字列を検索する。便利。
* @param <Sheet> sheet 検索対象のシート
* @param <String> val 検索文字列
* @param <int> col 検索列数(ex, A列 = 1)
* 
* @return {int} 行数
*/
function findRow(sheet, val, col) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

    for (var i = 1; i < dat.length; i++) {
        if (dat[i][col - 1] === val) {
            return i + 1;
        }
    }
    return 0;
}


function prepareConfig(cntKey) {
    if (isTest) {
        attendanceFolderId = ATTENDANCE_FOLDER_ID_FOR_TEST;
        attendanceSummarySheetId = ATTENDANCE_SUMMARY_SHEET_ID_FOR_TEST;
    } else {
        attendanceFolderId = ATTENDANCE_FOLDER_ID;
        attendanceSummarySheetId = ATTENDANCE_SUMMARY_SHEET_ID;

    }
    var properties = PropertiesService.getScriptProperties();
    countProperty = parseInt(properties.getProperty(cntKey));
    console.log("開始カウント：" + countProperty);

    if (cntKey === REPORTING_CNT_PROPERTY_KEY) {
        sheetCnt = parseInt(properties.getProperty("sheetCnt"));
        succeededCnt = parseInt(properties.getProperty("succeededCnt"));
        unfixedCnt = parseInt(properties.getProperty("unfixedCnt"));
        unmatchedCnt = parseInt(properties.getProperty("unmatchedCnt"));
        unexecutedCnt = parseInt(properties.getProperty("unexecutedCnt"));
        succeededWithoutProtectionCnt = parseInt(properties.getProperty("succeededWithoutProtectionCnt"));
        notFindSheetCnt = parseInt(properties.getProperty("notFindSheetCnt"));
        notFindStaffIdCnt = parseInt(properties.getProperty("notFindStaffIdCnt"));
        otherError = parseInt(properties.getProperty("otherError"));
        notInputedAmountCnt = parseInt(properties.getProperty("notInputedAmountCnt"));
    }
}

function outputEndSummaryLog() {
    console.info("シート総数(sheetCnt):" + sheetCnt);
    console.info("集計成功シート数(succeededCnt):" + succeededCnt);
    console.info("未確定シート数(unfixedCnt):" + unfixedCnt);
    console.info("実績未入力件数(notInputedAmountCnt)" + notInputedAmountCnt);
    console.info("該当月のシートが存在しない数(notFindSheetCnt):" + notFindSheetCnt);
    console.info("請求確定金額と勤務実績表の集計結果が不一致な数(unmatchedCnt):" + unmatchedCnt);
    console.info("スタッフ一覧に該当スタッフIDが見つからない数(notFindStaffIdCnt):" + notFindStaffIdCnt);
    console.info("その他のエラー(otherError):" + otherError);
    if (sheetCnt === (succeededCnt + unfixedCnt + notFindSheetCnt + notInputedAmountCnt + notFindStaffIdCnt + unmatchedCnt + otherError)) {
        console.info("[OK] シート総数と各カウント数が一致しました。")
    } else {
        console.info("[NG] シート総数と各カウント数が一致しませんでした。")
    }
}

function delete_specific_triggers(name_function) {
    var all_triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < all_triggers.length; ++i) {
        if (all_triggers[i].getHandlerFunction() == name_function)
            ScriptApp.deleteTrigger(all_triggers[i]);
    }
}

function needRestart(start_time, cntKey, currentCnt) {
    var current_time = new Date();
    var difference
        = parseInt((current_time.getTime() - start_time.getTime()) / (1000 * 60));

    //4分を超えていたら中断処理
    if (difference >= 4.5) {
        currentCnt++;
        var properties = PropertiesService.getScriptProperties();

        properties.setProperty(cntKey, currentCnt);
        console.log("次回再開カウント：" + currentCnt);
        if (cntKey === REPORTING_CNT_PROPERTY_KEY) {
            properties.setProperty("sheetCnt", sheetCnt);
            properties.setProperty("succeededCnt", succeededCnt);
            properties.setProperty("unfixedCnt", unfixedCnt);
            properties.setProperty("unmatchedCnt", unmatchedCnt);
            properties.setProperty("notInputedAmountCnt", notInputedAmountCnt);
            properties.setProperty("unexecutedCnt", unexecutedCnt);
            properties.setProperty("succeededWithoutProtectionCnt", succeededWithoutProtectionCnt);
            properties.setProperty("notFindSheetCnt", notFindSheetCnt);
            properties.setProperty("notFindStaffIdCnt", notFindStaffIdCnt);
            properties.setProperty("otherError", otherError);
        }

        return true;
    }
    return false;
}

function initializeProperies(cntKey) {
    var properties = PropertiesService.getScriptProperties();

    properties.setProperty(cntKey, 0);
    if (cntKey === REPORTING_CNT_PROPERTY_KEY) {
        properties.setProperty("sheetCnt", 0);
        properties.setProperty("succeededCnt", 0);
        properties.setProperty("unfixedCnt", 0);
        properties.setProperty("unmatchedCnt", 0);
        properties.setProperty("notInputedAmountCnt", 0);
        properties.setProperty("unexecutedCnt", 0);
        properties.setProperty("succeededWithoutProtectionCnt", 0);
        properties.setProperty("notFindSheetCnt", 0);
        properties.setProperty("notFindStaffIdCnt", 0);
        properties.setProperty("otherError", 0);
    }
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
  console.log ("showResultMessage, errorMessage=" +errorMessage);
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


function protect(currentAttendanceSheet,fixedPosition) {
  var protection = currentAttendanceSheet.protect();
  protection.setDescription(PROTECTION_DESCRIPTION);
  protection.setWarningOnly(true);

  var messageRange = currentAttendanceSheet.getRange(fixedPosition);
  messageRange.setValue(FIXED_MESSAGE);
  messageRange.setFontColor("red");
  messageRange.setFontWeight("bold");
  messageRange.setFontSize(14);
  messageRange.setHorizontalAlignment("right");

  console.log("sheetを保護しました。");
}