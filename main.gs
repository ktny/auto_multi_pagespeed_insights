/**
 * ページの表示速度を測定する
 */
function insightPagespeed() {

  // スプレッドシート全体に関わる変数
  var API_TOKEN_PAGESPEED = getScriptProperty('API_TOKEN_PAGESPEED');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(); // スプレッドシート内の全シートを取得

  // 各シートごとに関わる変数
  var sheetIndex = getScriptProperty('sheetIndex') ? parseInt(getScriptProperty('sheetIndex')) : 0; // 何番目のシートを処理するか
  var sheet = sheets[sheetIndex];
  var lastRow = sheet.getLastRow(); // そのシートの最終行を取得
  var today = Moment.moment().format('M月D日');

  // PageSpeedInsightsAPIのリクエストに関わる変数
  var device = sheet.getName().substr(-2); // シート名の後ろ2文字を切り出してデバイスを取得
  var strategy = device === 'PC' ? 'desktop' : 'mobile';　// デバイス別にクエリの値を取得

  // 同一シート内の各処理ごとに関わる変数
  var row = getScriptProperty('row') ? parseInt(getScriptProperty('row')) : 2; // 何番目の行から処理するか
  var urls = sheet.getRange(row, 1, lastRow - row + 1, 1).getValues(); // URL配列を現在の行から最後まで取得
  var scores = [];

  // 再起動用に開始時間を取得
  var start = Moment.moment()

  // 各行のURLのページスピード書き込みの処理が途中であれば最終列を、なければその次の列を取得する
  var column;
  if (row > 2) {
    column = sheet.getLastColumn();
  } else {
    column = sheet.getLastColumn() + 1;
    sheet.getRange(1,column).setValue(today); // 今日の日付を1行目に書き込む
  }

  Logger.log(sheetIndex + 'シート目' + column + '列' + row + '行目からのURLを処理中...');

  // 取得した全URLに対して処理
  for (var i = 0; i < urls.length; i++) {

    // URLが空の場合はスキップ
    var url = urls[i][0];
    if (!url) continue;

    // リクエストURLを作成
    var request = 'https://www.googleapis.com/pagespeedonline/v2/runPagespeed?url=' + url + '&key=' + API_TOKEN_PAGESPEED + '&strategy=' + strategy;

    // URLをAPIに投げてみてエラーが返ってくる場合はログに残す
    try {
      var response = UrlFetchApp.fetch(request, {muteHttpExceptions: true });
    } catch (err) {
      Logger.log(err);
      return(err);
    }

    // 返ってきたjsonをパース
    var parsedResult = JSON.parse(response.getContentText());
    var score = parsedResult.ruleGroups ? parsedResult.ruleGroups.SPEED.score : '-';

    // ページスピードスコアをscores配列に追加
    scores.push([score]);

    // 現在時間を取得して、開始から5分経過していたらforループ処理を中断して再起動
    var now = Moment.moment()
    if (now.diff(start, 'minutes') >= 5) {
      Logger.log('5分経過しました。タイムアウト回避のため処理を中断して再起動します。')
      break;
    }
  }

  // 取得したスコアを一度に書き込む
  sheet.getRange(row, column, scores.length, 1).setValues(scores);
  Logger.log(sheetIndex + 'シート目' + column + '列' + row + '行目から' + (row + scores.length) + '行目まで入力を行いました。')

  // rowを次の再起動用に設定
  row = row + scores.length;
  setScriptProperty('row', row);

  // 最終行まで処理していない場合は次の関数を再起動。最終行まで処理している場合は保存していた行を削除
  if (row < lastRow) {
    setTrigger('insightPagespeed');

  } else {
  　　　　deleteScriptProperty('row');

    sheetIndex++;

    // sheetIndexをスクリプトプロパティにセット。最終シートまで処理した場合はスクリプトプロパティを全て削除してChatworkで結果を共有
    if (sheetIndex < sheets.length) {
      setScriptProperty('sheetIndex', sheetIndex);
      setTrigger('insightPagespeed');
    } else {
      deleteScriptProperty('sheetIndex');
      deleteTrigger();
      sendMessage();
    }
  }
}

/**
 * Chatworkでメッセージを送信
 */
function sendMessage() {
  var API_TOKEN_CHATWORK = getScriptProperty('API_TOKEN_CHATWORK');
  var client = ChatWorkClient.factory({token: API_TOKEN_CHATWORK});
  var roomId = 45368752;
  var message = '[To:1641604] 松本　有加さん\n\
[To:854940] 佐藤　宜也さん\n\
[To:1784818] 浅田　亨さん\n\
[To:1741386] 折田　洋さん\n\
[To:1508994] 山内 卓朗さん\n\
今週のページスピード結果です(F)\n\
https://docs.google.com/a/cyber-ss.co.jp/spreadsheets/d/1FKzQVrTs-P0va7Fpxvzm-es3cGHqsBjkLnhs5g2j9B0/edit?usp=sharing';

  client.sendMessage({
    room_id: roomId,
    body: message
  });
}

/**
 * トリガーをセットする
 *
 * @param {function} func トリガーさせる関数
 */
function setTrigger(func) {
  deleteTrigger(); // 保存してるトリガーがあったら削除
  var triggerId = ScriptApp.newTrigger(func).timeBased().at(Moment.moment().add('minutes', 1).toDate()).create().getUniqueId();
  PropertiesService.getScriptProperties().setProperty('triggerId', triggerId);
}

/**
 * トリガーを削除する
 */
function deleteTrigger() {
  var triggerId = PropertiesService.getScriptProperties().getProperty('triggerId');

  if (!triggerId) return;

  ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getUniqueId() === triggerId;
  })
  .forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  PropertiesService.getScriptProperties().deleteProperty('triggerId');
}

/**
 * 特定のスクリプトプロパティを取得する
 *
 * @param {string} key キー
 * @return {string} スクリプトプロパティの値
 */
function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * スクリプトプロパティにキーと値をセットする
 *
 * @param {string} key キー
 * @param {number} value 値
 */
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * 特定のスクリプトプロパティを削除する
 *
 * @param {string} key キー
 */
function deleteScriptProperty(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}
