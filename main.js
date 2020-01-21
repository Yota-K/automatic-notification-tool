// 営業日のみ実行
function customTrigger() {
    var getDay = new Date().getDay();
  
    if (getDay === 0 || getDay === 6) {
      return;
    }
    else {
      sendMessage();
    }
}
  
function sendMessage() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet(),
        heading = sheet.getRange('B6:D6').getValues(),
        dayValue = sheet.getRange('E6').getValue(),
        progressValues = sheet.getRange('A7:D8').getValues(),
        headingIndex = -1,
        text = "";
    
    // 表全体のループ
    for (var r = 0; r < progressValues.length; r++) {
        // 列のループ
        for (var i = 0; i <progressValues[r].length; i++) {
            // 数字に"本"をつける
            if (typeof progressValues[r][i] === "number") {
              headingIndex += 1;
              text += '\n' + heading[0][headingIndex] + ' : ' + progressValues[r][i] + '本';
              //   一列終わったらリセット
              if (headingIndex >= 2) {
                  var headingIndex = -1;
                  continue;
                }
            }
            else {
                text += '[hr]' + '■'　+ progressValues[r][i];
            }
        }
    }
    
    // アクセストークン
    var client = ChatWorkClient.factory({
        token: 'APIトークン'
    });
  
    //送信
    client.sendMessage({
        room_id: 53710678,
        body: 
        '[info]' 
        + '[title]' + '本日の進捗状況！' + '[/title]' 
        + dayValue
        + text 
        + '[/info]'
    });
};
