# ツールの概要

## 開発した経緯
当ツールを開発する前はサイトの制作状況を、社員さんが目視で確認していたのですが、なんとかできないかと相談を受けた為、開発することを決意しました。

### 開発日数
たたき台は1日くらいで完成させました。

平日のみ起動させたり、見やすくしたり、表示する情報の追加をのちに行ったのですが、それ込みで2日かからないくらいで完成させました。

### GAS（GoogleAppsScript）とは？
> グーグルが提供しているJavaScriptをベースに作られたプログラム言語です。
>
> Googleアカウントさえあれば開発環境なしで簡単に利用できるほか、Googleスプレッドシート等の各Googleサービスと連携可能で、データ分析・グラフ作成なども効率化できま>す。
>
> Excelなどを効率化するプログラミング言語としてVBAがありますが、そのGoogle版といえるでしょう。
>
> https://goworkship.com/magazine/google-apps-script/

## ツールの仕組み

### スプレッドシートの設定
JSで値を取りやすくするために、別のスプレッドシートに進捗状況をまとめた表を作成します。

![img-1](https://user-images.githubusercontent.com/51050458/72822838-1ad36f00-3cb6-11ea-9189-a21f3b2167d9.jpg)

値の取得には、スプレッドシートのIMPORTXML関数を使用。
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

営業日（平日）のみにスクリプトを実行するための関数です。

getDayで実行される日にちの曜日を取得し、土日の時（0か６の時）はreturnで処理を行わないようにしてあります。

休日は送らないでほしいとの要望が来たため後から実装しました。

#### sendMessage
メインのスクリプトです。
```
function sendMessage() {
    // 現在のスプレッドシートを取得
    var sheet = SpreadsheetApp.getActiveSpreadsheet(),

        // 見出し部分のセルを取得
        heading = sheet.getRange('B6:D6').getValues(),

        // 残り営業日のセルを取得
        dayValue = sheet.getRange('E6').getValue(),

        // 表全体を取得
        progressValues = sheet.getRange('A7:D8').getValues(),

        // 残りの数の横の見出しを書き換えるためのインデックス
        headingIndex = -1,

        // テキストを代入するための変数
        text = "";
    
    // 表全体のループ
    for (var r = 0; r < progressValues.length; r++) {
        // 列のループ
        for (var i = 0; i < progressValues[r].length; i++) {
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
        room_id: ******,
        body: 
        '[info]' 
        + '[title]' + '本日の進捗状況！' + '[/title]' 
        + dayValue
        + text 
        + '[/info]'
    });
};

```

以下が基本の流れになります。

- 見出しと、残りの日数、表が書いてあるセルを取得し、for文でループ
- 数値の場合は本をつけて、文字列の場合は先頭に■をつけて変数textに代入
- 1行終わったらcotinueでスキップし、次の行の処理を行う

処理を行う変数progressValuesには下記のような配列が代入されています。

```[["\u30af\u30e9\u30a4\u30a2\u30f3\u30c8", -11, -6, -11], ["\u81ea\u793e\u6848\u4ef6", 0, 0, 0]]```

1つ目の配列に1行目、2つ目の配列に2行目の配列が入っています。

全ての処理が完了したら、client.senMessageで送信される仕組みになっています。

GASでスプレッドシートを操作するための関数については、こちらのサイトが詳しく解説しています。

> https://uxmilk.jp/25841

### 実行方法
トリガー（作成したコードを、特定の条件のときに自動で実行する仕組み）に関数を指定します。

時間は終業1時間前、18時にセットしています。

![img-3](https://user-images.githubusercontent.com/51050458/72828756-ad790b80-3cc0-11ea-9612-b4f3e55e3f2c.jpg)

うまく実行されると下記のような感じでチャットワークに送信されます。

![img-2](https://user-images.githubusercontent.com/51050458/72823002-5f5f0a80-3cb6-11ea-88d5-9384d8ff5d1b.jpg)
