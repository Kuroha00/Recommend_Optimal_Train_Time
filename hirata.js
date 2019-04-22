/**
zabutonbotのコード．Slack API, またはLINE Messaging APIを使用．GASのトリガー機能を利用して指定した時間への実行を実現している．
setTrigger()が毎日0時〜1時の間に実行され，その時点でスプレッドシート上に設定されていた時刻全てに，その日のトリガーを設定する．
これは「毎日特定の時刻に」実行するトリガー設定がデフォルトで存在しないため，分単位で正確な時刻に関数を実行するためにこのように実装した．

使い方：
スプレッドシート "zabutonbot" シート "messages" のコメントに従って記入してください．
LINEの場合は最初に友達登録して，なんでもいいので1回メッセージを送信する必要があります．

流用する場合：
Slack, LINEそれぞれ送信元を変更するには，それぞれアクセストークンを書き換えてください．
Slackのアクセストークン（HEILAB_SLACK_ACCESS_TOKEN）は，ファイル→プロジェクトのプロパティ→スクリプトのプロパティにあります．
LINEはLINE_CHANNEL_ACCESS_TOKENです．LINEの場合はLINE DevelopersのサイトからWebhook URLに，このスクリプトを「公開」から「ウェブアプリケーションとして導入…」したところの「ウェブ アプリケーションの URL」を貼り付ける必要があります．
*/

var spreadSheet = SpreadsheetApp.openById('1pc5-pfcxkFCIs62HTfSTiZhzv9lddRAS0WmJpTPGptQ');//zabutonbotの情報を書き込むスプレッドシート
var mesListSheet=spreadSheet.getSheetByName('messages');//メッセージを送信する時刻，内容，媒体，宛先等を記録するシート
var tokenSheet=spreadSheet.getSheetByName('tokens');//未使用
var lineUsersSheet=spreadSheet.getSheetByName('lineusers');//LINE bot「座布団ボット」が友達登録され，最初にメッセージを受信したときにその相手のLINEのuser_id等を記録する．LINEでリマインダを送るためにプッシュするときにこのuser_idを使用．
var lineLogSheet=spreadSheet.getSheetByName('linelog');//LINEの会話内容のログをとっている

var LINE_CHANNEL_ACCESS_TOKEN = 'vxKoPd3V0fRGbbn2T6WMDIAtl0lc/OMVhIQU0SqFwkIIDmhamiYWv0FkLf+4JDfU+VuBI8uKYT7dbXQcMaj+3b6bYNWzy/FK/AqcYyVTUkYbpyOfcb5pbJgV6N7KkV8dELxZLWjkBPxy7RvZXXqa8wdB04t89/1O/w1cDnyilFU=';
//「座布団ボット」のLINE Messaging APIのアクセストークン


var line_endpoint = 'https://api.line.me/v2/bot/message/reply';
var line_endpoint_profile = 'https://api.line.me/v2/bot/profile';
var line_endpoint_push = 'https://api.line.me/v2/bot/message/push';


function getLineUserData(user_id) {
  //LINEユーザーIDを指定すると，LINEユーザーデータを返す
  var res = UrlFetchApp.fetch(line_endpoint_profile + '/' + user_id, {
                              'headers': {
                              'Content-Type': 'application/json; charset=UTF-8',
                              'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
                              },
                              'method': 'get',
                              });
  return JSON.parse(res);
}


function writeLineUserData(user_id){
  //LINEユーザーデータをスプレッドシートに追記
  var u_dat=getLineUserData(user_id);
  var writeData=[u_dat.displayName,user_id,'',u_dat.pictureUrl,u_dat.statusMessage];
  lineUsersSheet.appendRow(writeData);
  return writeData;
}



function linepushtest(){
  //LINEプッシュのテストをするための関数
  linePushMessage("hirata",["テストメッセージです","a"]);
}


function linePushMessage(toname,push_messages){
  //LINEでメッセージをプッシュするメインの関数．push_messagesが配列だと複数のメッセージを送信できる．linepushtest()参照．
  var dat = lineUsersSheet.getDataRange().getValues();
  var user_id;
  for(var r=1;r<dat.length;r++){
    if(dat[r][5]==toname){
      user_id=dat[r][1]; 
    }
  }
  var messages = push_messages.map(function (v) {
    return {'type': 'text', 'text': v};    
  });   
  UrlFetchApp.fetch(line_endpoint_push, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'to': user_id,
      'messages': messages
    }),
  });
  var today=new Date();
  lineLogSheet.appendRow([today,'text' , push_messages[0], 'bot']);
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}


function shortenUrl(longUrl) {
  var url = UrlShortener.Url.insert({
    longUrl: longUrl
  });
  return url.id;
}

function postSlackMessage(toname,message){
  Logger.log(toname)
  //Slackでメッセージを送るメインの関数
  var token = "xoxp-164093776417-346697431441-376390430051-881f06861f90dea7bd984bc9f9096c59"
  //var token = PropertiesService.getScriptProperties().getProperty('HEILAB_SLACK_ACCESS_TOKEN');
  var slackApp = SlackApp.create(token);
  var channel = "toname";  //投稿先チャンネル名
  var username = "ZabutonBot";
  //  var shortLink = shortenUrl(url);
  
  var payload = {
    "channel" : channel, //通知先チャンネル名
    "text" : message, //送信テキスト
    "username": username, //BOTの名前
  };
  
  
  var option = {
    "method" : "POST", //POST送信
    "payload" : payload //POSTデータ
  };
//  Logger.log(message)
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token="+token+"&channel="+toname+"&text="+message+"&username="+username+"&pretty=1");
  //UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + token, option)
  //slackApp.postMessage(channel, message, {username:username});
}


/*
function Slacktest(){
  var message_test1 = "こんにちは"
  var message_test2 = "お疲れ様です。アンケートへの回答をお願いします。・SRS-18 https://goo.gl/forms/xziL6DZNjNrbxMvi1 （毎日開始時と終了時に答えていただけるとありがたいです） ・PSS-14 https://goo.gl/forms/CVosizGLu2WWhqev2 （火曜日と金曜日の開始時にお願いします）"
  var message_test3 = "みなさんお時間があれば今日の分のアンケートへの回答をお願い致します。m(_ _)m"
  postSlackMessage("GA77919NE", message_test2);
}
*/

var flagcol=4;

function postCheck(){
  //setTrigger()によって，この関数がスプレッドシートに指定した時刻にトリガーが設定されて実行される．送信するメッセージがあるか，シート"messages"を巡回して送信していく．
  var dat = mesListSheet.getDataRange().getValues();
  for(var r=1;r<dat.length;r++){
    var time=dat[r][0];
    var message = dat[r][1];
    var names=[];
    for(var c=4;c<dat[r].length;c++){
      if(dat[r][c]!=""){
        names.push(dat[r][c]);
      }
    }
    var today  = new Date();
    var time_c = new Date();
    time_c.setHours(time.getHours());
    time_c.setMinutes(time.getMinutes());
    var minLag=(time_c-today)/60000;
//    Logger.log(time_c)
//    Logger.log(today)
//    Logger.log(minLag)
    if(Math.abs(minLag)<0.5 && dat[r][flagcol-1]!=1){
      for(var i=0;i<names.length;i++){
        switch (dat[r][2]){
          case "heilabslack":
            postSlackMessage(names[i],message); 
            break;
          case "line":
            linePushMessage(names[i],[message]);
            break;
        }
      }
      mesListSheet.getRange(r+1,flagcol).setValue(1);
    }
  }
}

function setTrigger() {
//postCheck()にトリガーを設定する関数．この関数が毎日0~1時に実行されるようトリガーを設定しておく．
  deleteTrigger();
  var triggerDay = new Date(); 
  var pushtime;
  var dat = mesListSheet.getDataRange().getValues();
  
  // 火曜と金曜にpssを送るときに二重で送ってしまう問題対策
  var pss_len = 3;
  var r_len = dat.length-pss_len;
  
  for(var r=1;r<r_len;r++){
    var time=dat[r][0];
    triggerDay.setHours(time.getHours());
    triggerDay.setMinutes(time.getMinutes());  
    ScriptApp.newTrigger("postCheck").timeBased().at(triggerDay).create();
  }
  mesListSheet.getRange(2,flagcol,dat.length-1).clear();
}

function deleteTrigger() {
  // その日のトリガーを削除する関数(消さないと残る)
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "postCheck") {
      Utilities.sleep(1000) 
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function findRow(sheet,val,col,mode){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  if(mode==0){
    for(var i=1;i<dat.length;i++){
      if(dat[i][col-1] === val){
        return i+1;
      }
    }}else{
      for(var i=dat.length-1;i>0;i--){
        if(dat[i][col-1] === val){
          return i+1;
        }
      }
    }
  return 0;
}

function doPost(e) {
  var json = JSON.parse(e.postData.contents);  
  var user_id = json.events[0].source.userId;
  var sourcetype = json.events[0].source.type;
  if(sourcetype=='group'){
    var group_id = json.events[0].source.groupId;
  }else if(sourcetype == 'room'){
    var room_id = json.events[0].source.roomId;    
  }
  var reply_token= json.events[0].replyToken;
  if (typeof reply_token === 'undefined') {
    return;
  }
  var user_message_type = json.events[0].message.type;
  var user_message;
  if(user_message_type=='text'){
    user_message = json.events[0].message.text;
    user_message=user_message.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 65248);
    });
  }else if(user_message_type=='sticker'){
    var packageId= json.events[0].message.packageId;
    var stickerId= json.events[0].message.stickerId;
    user_message=packageId + '-' + stickerId;
  }else if(user_message_type=='location'){
    var lat= json.events[0].message.latitude;
    var lng= json.events[0].message.longitude;
    user_message=lat + ',' + lng;
  }
  
  var dat = lineUsersSheet.getDataRange().getValues();
  var nickname;
  for(var r=1;r<dat.length;r++){
    if(dat[r][1]==user_id){
      nickname=dat[r][0]; 
    }
  }
  if(nickname===undefined){
    nickname=writeLineUserData(user_id)[0];    
  }
  var reply_messages;
  var today = new Date();
  lineLogSheet.appendRow([today, user_message_type, user_message, nickname,sourcetype]);
}