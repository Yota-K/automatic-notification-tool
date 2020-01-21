# ツールの概要

## 開発した経緯
当ツールを開発する前はサイトの制作状況を、社員さんが目視で確認していたのですが、なんとかできないかと相談を受けた為、開発することを決意しました。

### 開発日数
たたき台は1日くらいで完成させました。

平日のみ起動させたり、見やすくしたり、表示する情報の追加をのちに行ったのですが、それ込みで2日かからないくらいで完成させました。

### GAS（GoogleAppsScript）とは？
グーグルが提供しているJavaScriptをベースに作られたプログラム言語です。

Googleアカウントさえあれば開発環境なしで簡単に利用できるほか、Googleスプレッドシート等の各Googleサービスと連携可能で、データ分析・グラフ作成なども効率化できます。

Excelなどを効率化するプログラミング言語としてVBAがありますが、そのGoogle版といえるでしょう。

## ツールの仕組み

### スプレッドシートの設定
JSで値を取りやすくするために、別のスプレッドシートに進捗状況をまとめた表を作成。

![img-1](https://user-images.githubusercontent.com/51050458/72822838-1ad36f00-3cb6-11ea-9189-a21f3b2167d9.jpg)

値はスプレッドシートのIMPORTXML関数を使用
スクリプト内部にはAPIトークンが書き込まれており漏洩の危険性を防ぐために、自分以外はアクセスできなくしてあります。

### 実行するスクリプトについて

#### customTrigger

```
function customTrigger() {
    var getDay = new Date().getDay();
  
    if (getDay === 0 || getDay === 6) {
      return;
    }
    else {
      sendMessage();
    }
}
```

営業日（平日）のみに送るための関数です。

休日は送らないでほしいとの要望が来たため後から実装しました。

#### sendMessage
メインのスクリプトです。
```
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

```

### 実行方法

うまく実行されると下記のような感じでチャットワークに送信されます。

![img-2](https://user-images.githubusercontent.com/51050458/72823002-5f5f0a80-3cb6-11ea-88d5-9384d8ff5d1b.jpg)

