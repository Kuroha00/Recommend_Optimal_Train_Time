/* 
柏キャンパスから自宅に帰るときに，乗り換え時間の少ない電車で帰るため，
その時間の20分前くらいにslackに通知する
ファイル→プロジェクトのプロパティ→スクリプトのプロパティ
・verifyToken: Outgoing Webhooksのトークン．　動かない可能性がある
・slackToken: slackAPIトークン
*/

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// var spreadsheet = SpreadsheetApp.openById('1JBr9r18LwwVr-mV37N_HpC19wgD64aXci_zrvvhJpWc'); //この書き方でもできる

var train_sheet　=　spreadsheet.getSheetByName("train");  //メッセージを送信する時刻，内容，媒体，宛先等を記録するシート
var button_sheet = spreadsheet.getSheetByName("button");
var slack_channel = "GJ7LQ5RT6";
//var my_slack = "DANV24ETV";
var slack_token = PropertiesService.getScriptProperties().getProperty("slackToken");
var notification_time = 20;

var dat = train_sheet.getDataRange().getValues();
for (var r=0; r<dat[0].length; r++){
  if (dat[0][r] == "flag"){
    var flagcol = r;
  }
}


function postCheck(){
  //setTrigger()によって，この関数がスプレッドシートに指定した時刻にトリガーが設定されて実行される．
  //   送信するメッセージがあるか，シート"messages"を巡回して送信していく
  
  for(var r=1; r<dat.length; r++){
    var time = dat[r][0];
    var after_time = dat[r][1];
    var today  = new Date();
    var time_c = new Date();
    
    time_c.setHours(time.getHours());
    time_c.setMinutes(time.getMinutes());
    
    today.setMinutes(today.getMinutes() + notification_time);
    var minLag = (time_c - today) / 60000;
    Logger.log(minLag);
    
    if(Math.abs(minLag)<0.5 && dat[r][flagcol]!=1){
      if (String(time.getMinutes()).length==1){
        var departure_time = time.getHours() + ":0" + time.getMinutes();
      }else{
        var departure_time = time.getHours() + ":" + time.getMinutes();
      }
//      Logger.log("aa");
      if (String(after_time.getMinutes()).length==1){
        var arrival_time = after_time.getHours() + ":0" + after_time.getMinutes();
      }else{
        var arrival_time = after_time.getHours() + ":" + after_time.getMinutes();
      }
      
//      Logger.log("bb");
      if ( (today.getDay()==0) || (today.getDay()==6) ){
        // 土日のときは送らない．
        var a;
        Logger.log("cc");
      }else{
        message = "江戸川台 " + departure_time + "発  /  牛久 " + arrival_time + "着（所要時間" + dat[r][2] + "分，" + dat[r][3] + "行き，" + dat[r][4] + "両）";
        Logger.log(message);
        postSlackMessage(slack_channel, message);
      }
      train_sheet.getRange(r+1, flagcol+1).setValue(1);
    }
  }
}


function setTrigger() {
  //postCheck()にトリガーを設定する関数．この関数が毎日0~1時に実行されるようトリガーを設定しておく．
  deleteTrigger();
  
  var triggerDay = new Date(); 
  for(var r=1; r<dat.length; r++){
    var time = dat[r][0];
    triggerDay.setHours(time.getHours());
    triggerDay.setMinutes(time.getMinutes() - notification_time);
    ScriptApp.newTrigger("postCheck").timeBased().at(triggerDay).create();
  }
  train_sheet.getRange(2, flagcol+1, dat.length-1).clear();
}


function deleteTrigger() {
  // その日のトリガーを削除する関数(消さないと残る)
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "postCheck") {
      Utilities.sleep(1000);
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function postSlackMessage(toname, message){
  // Slackでメッセージを送るメインの関数
  var slackApp = SlackApp.create(slack_token);
  var channel = toname;  //投稿先チャンネル名
  var username = "TrainBot";
  
//  var payload = {
//    "channel" : channel, //通知先チャンネル名
//    "text" : message, //送信テキスト
//    "username": username, //BOTの名前
//  };
  
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token="+token+"&channel="+toname+"&text="+message+"&username="+username+"&pretty=1");
}


//プライベートチャンネルではOutgoing Webhookは使えなさそう
//function doPost(e) {
//  //スクリプトプロパティ取得
//  var prop = PropertiesService.getScriptProperties().getProperties();
//  var bot_name = "Train_bot";
//  
//  var app = SlackApp.create(prop.slackToken);
//  //POSTされた情報を使ってメッセージを作る
//  var message = e.parameter.user_name + "さんが" + e.parameter.channel_name + "で" + e.parameter.text + "と言いました";
//  
//  //POST元が違っていたらエラー
//  if (prop.verifyToken != e.parameter.token) {
//    throw new Error("invalid token.");
//  }
//  
//  //特定のチャンネルにメッセージをPOSTする
//  return app.postMessage("train_bot", message , {
//    username: bot_name,
//  });
//}




