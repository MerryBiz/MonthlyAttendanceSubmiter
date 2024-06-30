/**
 * 2023.9.21 作成
 * @n.yasuda
 * 各種情報シート内容更新スクリプト */

function updateBasicInfo(gAccount, sId) {
 
  console.log("version4 run");
  var spreadsheetId = sId;
  var basicInfoSpreadSheet = SpreadsheetApp.openById(spreadsheetId);
  var basicInfoSpreadSheetName = basicInfoSpreadSheet.getName();
   // スタッフIDを取り出す
  var staffId = basicInfoSpreadSheetName.slice(0,5);
  console.log(staffId);
  var basicInfoSheet = basicInfoSpreadSheet.getSheetByName("各種情報");
   if (basicInfoSheet === null) {
    console.log("処理を終了します。");
    
    // 処理を終了
    return;
   }
  // 初回フラグ
  // C2セルの値を取得
  var initflg = basicInfoSheet.getRange("C2").getValue();
  if (initflg === "") {
    basicInfoSpreadSheet.toast("初回情報取得中","処理中",0);
    basicDataWrite(basicInfo, staffId, accountInfo, billingSource, basicInfoSheet);
    basicInfoSpreadSheet.toast("初回処理中","初回情報の取得が終わりました",1);
  } else {
    basicDataWrite(basicInfo, staffId, accountInfo, billingSource, basicInfoSheet);
  }
}


function basicDataWrite(basicInfo, staffId, accountInfo, billingSource, basicInfoSheet) {
  // スタッフID,氏名、フリガナ、時間単価をCM表から配列で取得
  var basicResults = basicInfo(staffId);
  // 金融機関名、支店名、預金種目、口座番号、口座名義人を振込先口座シートから取得
  var accountResults = accountInfo(staffId);
  // 住所1（都道府県 / 海外の場合国名）,住所2（市区町村以下）,氏名または事業者名,インボイス登録番号をスタッフマスタから取得
  var billingSourceResults = billingSource(staffId);
 
  // C2からC5に基本情報を書き込む
  basicInfoSheet.getRange("C2:C5").setValues(basicResults);
  // C8からC12に口座情報を書き込む
  basicInfoSheet.getRange("C8:C12").setValues(accountResults);
  // C14からC18に口座情報を書き込む
  basicInfoSheet.getRange("C14:C18").setValues(billingSourceResults);
}



function basicInfo(staffId) {
  // CM表を開く（スプレッドシートIDを指定）
  var targetSpreadsheetId = "1FAOiTbtrqIU81RgFrZfIoxtgQHo7gH0gX-jOZXr4gXM";
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // 対象スプレッドシートのシートを指定（シート名を指定）
  var sheet = targetSpreadsheet.getSheetByName("CM表");

  // H列のデータを取得
  var staffIdColumnData = sheet.getRange("B:B").getValues();
  
  // メールアドレスを検索
  var rowIndex = -1;
  for (var i = 0; i < staffIdColumnData.length; i++) {
    if (staffIdColumnData[i][0] === staffId) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1) {
    // CM表でスタッフIDを一致する行を見つけた場合、氏名、フリガナ、時間単価の情報を取得
    var fullName = sheet.getRange(rowIndex + 1, 3).getValue(); // C列の情報
    var kanaName = sheet.getRange(rowIndex + 1, 4).getValue(); // D列の情報
    var basicPrice = sheet.getRange(rowIndex + 1, 69).getValue(); // BQ列の情報 
    console.log(staffId + "：" + fullName + "：" + kanaName + "：" + basicPrice);
    var basicResults = [[staffId], [fullName], [kanaName], [basicPrice]];

    return basicResults;
  } else {
    console.log("スタッフIDが見つかりませんでした");
  }
}

function accountInfo(staffId) {
  // 振込先口座シートを開く（スプレッドシートIDを指定）
  var targetSpreadsheetId = "1I4VY67wgE_qd87J8SCXE9eUV7zrwOueKpdi01MuccUA";
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // 対象スプレッドシートのシートを指定（シート名を指定）
  var sheet = targetSpreadsheet.getSheetByName("フォームの回答 1");

  // J列のデータを取得
  var staffIdColumnData = sheet.getRange("J:J").getValues();

  // スタッフIDを検索
  var rowIndex = -1;
  for (var i = 0; i < staffIdColumnData.length; i++) {
    if (staffIdColumnData[i][0] === staffId) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1) {
    // 振込先口座シートでスタッフIDが一致する行を見つけた場合、金融機関名、支店名、預金種目、口座番号、口座名義人の情報を取得
    var financialInstitution = sheet.getRange(rowIndex + 1, 3).getValue(); // C列の情報
    var branchName = sheet.getRange(rowIndex + 1, 5).getValue(); // E列の情報
    var accountType = sheet.getRange(rowIndex + 1, 7).getValue(); // G列の情報
    var accountNum = sheet.getRange(rowIndex + 1, 8).getValue(); // H列の情報
    var accountName = sheet.getRange(rowIndex + 1, 9).getValue(); // I列の情報 
    console.log(financialInstitution + "：" + branchName + "：" + accountType + "：" + accountNum+ "：" + accountName);
    var accountResults = [[financialInstitution], [branchName], [accountType], [accountNum], [accountName]];

    return accountResults;
  } else {
    console.log("スタッフIDが口座情報シートにが見つかりませんでした。");
  }
}

function billingSource(staffId) {
  // スタッフマスタスプレッドシートを開く（スプレッドシートIDを指定）
  var targetSpreadsheetId = "1n8pTG6Hvs2cbSudMuZBXwcyZLotNomQ_eR_QraYBmEs";
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // スタッフマスタシートのシートを指定（シート名を指定）
  var sheet = targetSpreadsheet.getSheetByName("スタッフマスタ");

  // A列のデータを取得
  var staffIdColumnData = sheet.getRange("A:A").getValues();

  // スタッフIDを検索
  var rowIndex = -1;
  for (var i = 0; i < staffIdColumnData.length; i++) {
    if (staffIdColumnData[i][0] === staffId) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1) {
    // スタッフマスタシートでスタッフIDが一致する行を見つけた場合、郵便番号、住所1（都道府県 / 海外の場合国名）,住所2（市区町村以下）,氏名または事業者名,インボイス登録番号の情報を取得
    var prefectures = sheet.getRange(rowIndex + 1, 6).getValue(); // F列の情報
    var address = sheet.getRange(rowIndex + 1, 7).getValue(); // G列の情報
    var billingSource = sheet.getRange(rowIndex + 1, 4).getValue(); // D列の情報
    var invoiceNum = sheet.getRange(rowIndex + 1, 5).getValue(); // E列の情報
    var fullAddress = prefectures + address;
    var postalCode = searchZip(fullAddress);
    console.log(prefectures + "：" + address + "：" + billingSource + "：" + invoiceNum + "：" + postalCode);
    var billingResults = [[postalCode], [prefectures], [address], [billingSource], [invoiceNum]];

    return billingResults;
  } else {
    console.log("スタッフIDがスタッフマスタシートにが見つかりませんでした。");
  }
}

/**
 * 住所から郵便番号検索
 */
function searchZip(fullAddress) {
  var response;
  var zip;
  var url = 'https://google.co.jp/maps/search/';
 
  // データをすべて配列に
  var value = fullAddress;
   
  try {
        // グーグルマップで検索
        response = UrlFetchApp.fetch(url + value).getContentText();       
        // 検索結果に郵便番号があるかチェック
        zip = response.search('〒');
        if (zip >= 0) {
        // 郵便番号を取得
        value = response.substr(zip + 1, 8);
        } else { 
          // 取得できなかった住所はログに
          console.log(value);
        }
      } catch(e) {
     
    // 何かエラーがあれば出力
    console.log(e);
  }
     
  // データを戻す
  return value;  
}