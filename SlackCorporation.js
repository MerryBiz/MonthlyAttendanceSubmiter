//定数
var SLACK_CHANNEL = "#3sx-3cm_スタッフ勤務実績自動集計";
var BOT_NAME = "勤務表自動集計Bot";
var MENTIONS = ["@wakaki.imamura", "@sayaka.kubota", "@yuko.kamikawa"];


/*
*/
function sendSlackReportingMessage() {
    send(generateReportingMessage());

}


/*
*/
function sendSlackSuspendMessage() {
    send(generateSuspendtMessage());

}


/*
* @param {string} mesage
*/
function send(message) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var url = scriptProperties.getProperty('SLACK_WEBHOOK_URL'); // URLをスクリプトプロパティから取得
    
    var data = { "channel": SLACK_CHANNEL, "username": BOT_NAME, "text": message };
    var payload = JSON.stringify(data);
    
    var options = {
        "method": "POST",
        "contentType": "application/json",
        "payload": payload
    };
    
    var response = UrlFetchApp.fetch(url, options);
}


function generateSuspendtMessage() {
    var message = "";

    for (var i = 0; i < MENTIONS.length; i++) {
        message += "<" + MENTIONS[i] + "> "
    }

    message += "\n";
    message = addSentense(message, "※エラーに付き、処理が中断されました。「在宅スタッフの勤務実績の金額」に集計対象シートが生成されているかご確認ください。");

    return message;
}

function generateReportingMessage() {
    var message = "";

    for (var i = 0; i < MENTIONS.length; i++) {
        message += "<" + MENTIONS[i] + "> "
    }

    message += "\n";
    message = addSentense(message, "勤務実績表に集計状況をレポートします。");
    message += "```";
    message = addSentense(message, "総実行件数；" + sheetCnt);
    message = addSentense(message, "===================================");
    message = addSentense(message, "◯ 集計成功件数               ：" + succeededCnt);
    message = addSentense(message, "▲  未確定件数                 ：" + unfixedCnt);
    message = addSentense(message, "▲  処理待ち件数               ：" + unexecutedCnt);
    message = addSentense(message, "▲  実績未入力件数(0円)        ：" + notInputedAmountCnt);
    message = addSentense(message, "！ 要チェック件数              ：" + succeededWithoutProtectionCnt);
    message = addSentense(message, "-----------------------------------");
    message = addSentense(message, "－ 先月分のシートがない件数    ：" + notFindSheetCnt);
    message = addSentense(message, "×  確定額が不一致な件数        ：" + unmatchedCnt);
    message = addSentense(message, "×  スタッフIDが見つからない件数：" + notFindStaffIdCnt);
    message = addSentense(message, "×  その他のエラー件数          ：" + otherError);
    message = addSentense(message, "===================================");
    if (sheetCnt === (succeededCnt + unfixedCnt + unexecutedCnt + notInputedAmountCnt + succeededWithoutProtectionCnt + notFindSheetCnt + notFindStaffIdCnt + otherError + unmatchedCnt)) {
        message = addSentense(message, "◯ 総件数と詳細件数のカウントが一致しました。");
    } else {
        message = addSentense(message, "× 総件数と詳細件数のカウントが一致しませんでした。調査が必要です！！");
    }
    message += "```";

    if (needCheckSSNamesList.length > 0) {
        message += "\n";
        message = addSentense(message, "【要チェックリスト】");
        for (var j = 0; j < needCheckSSNamesList.length; j++) {
            message = addSentense(message, needCheckSSNamesList[j]);
        }

    }

    return message;
}


function addSentense(message, sentense) {
    message += sentense;
    message += "\n";
    return message;
}
