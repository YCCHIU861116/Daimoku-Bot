var Channel_access_token = 'QimQsHHol8kE3jxZds4NkgDbaqI4frciop6sts0s1atl8YOT+xIELoFJ8q+VMZPIjc7tvo9/FPpLxrCjHDuog+1TFr4NP3r/B9G+js6Bp6WD6zWhVwCDx1PFMjuMAuZ9hL6aG8Uqr6s8pWefJQ6iDQdB04t89/1O/w1cDnyilFU=';
var Sheet_Id = '1amor-fS4IzOG8pPwV5hrwNhdj647fZWu_uoaWMqDr6I';
var Test_Sheet_Id = '1-GQ4KZjf1tcELwyGICq1lvOq0iW2jq_k4E69-N-hceY';
var spreadsheet = SpreadsheetApp.openById(Sheet_Id);
var sheet = spreadsheet.getSheets()[0];
var Response_Sheet_Id = '18DMAruCVS4HUe_YrcnPidPbMCMksqUJrV8DdaPzaIcQ';
var res_spreadsheet = SpreadsheetApp.openById(Response_Sheet_Id);
var res_sheet = res_spreadsheet.getSheets()[0];

function alarm(){
  var i;
  var report_num = sheet.getRange(2,7).getValue();
  for(i = 0; i < report_num; i++){
    var userId = sheet.getRange(i+7,6).getValue();
    var Name = sheet.getRange(i+7,7).getValue();
    var pushtext = Name + res_sheet.getRange(9,10).getValue() +sheet.getRange(2,6).getValue()+res_sheet.getRange(9,11).getValue();
    var payload = {
      'to': userId,
      'messages':[{
        'type': 'text',
        'text': pushtext
      }]
    };
    var post_options = {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': 'Bearer '+ Channel_access_token
      },
      'method': 'post',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    var push_url = 'https://api.line.me/v2/bot/message/push';
    var res = UrlFetchApp.fetch(push_url, post_options);
    if(res)
      sheet.getRange(2,5).setValue(res.getContentText());
  }
}

function timeparse(timestamp){
  var days_in_month = [31,29,31,30,31,30,31,31,30,31,30,31];
  var ms_after_20200101 = parseInt(timestamp) - 1577836800000;
  var ms = ms_after_20200101 % 1000;
  var sec_after_20200101 = Math.floor(ms_after_20200101 / 1000);
  var sec = sec_after_20200101 % 60;
  var min_after_20200101 = Math.floor(sec_after_20200101 / 60);
  var min = min_after_20200101 % 60;
  var hr_after_20200101 = Math.floor(min_after_20200101 / 60)+8;
  var hr = hr_after_20200101 % 24;
  var day_after_20200101 = Math.floor(hr_after_20200101 / 24);
  var month = 1, days = day_after_20200101+1;
  var i;
  for(i = 0; i < days_in_month.length; i++){
    if(days > days_in_month[i]){
      days -= days_in_month[i];
      month ++;
    }
  }
  var timestamp_string = '2020/'+month+'/'+days+' '+ hr +':'+min+':'+sec;//+':'+ ms +'/'+ ms_after_20200101+'/'+day_after_20200101;
  return timestamp_string;
}

function find_daimoku_num(Name){
  sheet.getRange(1, 7).setValue('用戶數');
  var report_num = sheet.getRange(2,7).getValue();
  var i,complete = 0,index=0;
  for(i = 0; i < report_num && !complete; i++){
   if(Name == sheet.getRange(i+7,7).getValue()){
     complete = 1;
     index = i;
   }
  }
  if(complete){
    return sheet.getRange(index+7,8).getValue();
  }
  else{
    sheet.getRange(1, 7).setValue('error!!');
    return -1;//error!!!
  }
}

function find_reply(message){
  res_sheet.getRange(1, 7).setValue('總回應次數');
  var res_num = res_sheet.getRange(2,7).getValue();
  var new_row = res_num + 5;
  var i,complete = 0,index=0;
  for(i = 0; i < res_num && !complete; i++){
   if(message == res_sheet.getRange(i+5,6).getValue()){
     complete = 1;
     index = i;
   }
  }
  if(complete){
    return res_sheet.getRange(index+5,8).getValue();
  }
  else{
    res_sheet.getRange(2,7).setValue(res_num+1);
    res_sheet.getRange(new_row, 6).setValue(message);
    res_sheet.getRange(new_row, 7).setValue("=sumif(C:C,\""+message+"\",D:D)");
    res_sheet.getRange(new_row, 8).setValue("empty");
    return "empty"
  }
}

function Reply(Name,timestamp, userMessage){
  var replytext;
  sheet.getRange(1, 6).setValue('總計');
  if(user_daimoku_num == -1){
    return "error";
  }
  var msg =  userMessage.split(' ');
  if(isNaN(userMessage)){
    if (userMessage === "大家唱了多少" || userMessage === "大家唱多少了"){
      var all_daimoku_num = sheet.getRange(2,6).getValue();
      replytext = '大家已經共戰了 ' + all_daimoku_num + ' 分鐘囉\n我們一起前進吧！';
    }
    else if(userMessage === "我唱了多少" || userMessage === "我唱多少了" ){
      var user_daimoku_num = find_daimoku_num(Name);
      if(user_daimoku_num == 0)
        replytext = "尼根本還沒唱，快唱！";
      else
        replytext = '尼已經共戰了 ' + user_daimoku_num + ' 分鐘囉，繼續加油RRR';
      }
    else if(msg[0] === 'Undo' || msg[0] === 'undo'){
      var lastRow = sheet.getLastRow();
      var i,complete = 0;
      for(i = lastRow; i > 0 && !complete; i--){
        if(sheet.getRange(i,2).getValue() == Name){
          sheet.getRange(i,3).setValue(msg[1]);
          complete = 1; 
        }
      }
      replytext = '修正完成，下次確定再輸入可以ㄇ(怒';
    }
    else if(userMessage === "目標"){
      replytext = res_sheet.getRange(5,10).getValue();
    }
    else if(userMessage === "help"){
      replytext = res_sheet.getRange(7,10).getValue();
    }
    else{
      var newresContents = [timestamp,Name,userMessage,1];
      res_sheet.appendRow(newresContents);
      var reply_to_res = find_reply(userMessage);
      if(reply_to_res === 'empty')
        replytext = res_sheet.getRange(2,11).getValue();//Name + 
      else
        replytext = reply_to_res;
    }
  }
  else{
    sheet.getRange(1, 3).setValue('題目數');
    var daimoku_num = parseInt(userMessage);
    var newrowContents = [timestamp,Name,daimoku_num];
    if(daimoku_num >=100000){
      return res_sheet.getRange(2,9).getValue();
    }
    else if(daimoku_num<0 || userMessage.indexOf(".")>-1){
      return res_sheet.getRange(2,10).getValue();
    }
    sheet.appendRow(newrowContents);
    if(daimoku_num >= 60){
      replytext = Name + '!!!!!!!!\n';
      if(daimoku_num == 60){
        replytext += '一小時';
      }
      else{
        replytext += (daimoku_num + '分鐘');
      }
      replytext += ' 太神啦';
    }
    else{
      replytext = 'hen 棒';
    }
    var user_daimoku_num = find_daimoku_num(Name);
    replytext += '\n你現在已經和大家共戰了' + user_daimoku_num +' 分鐘囉！繼續加油~~~';
  }
  return replytext;
}

function FindNamebyProfile(userId){
  var get_options = {
    'headers': {
      'Authorization': 'Bearer '+ Channel_access_token
    },
    'method': 'get'
  };
  var profile_url = 'https://api.line.me/v2/bot/profile/' + userId;
  var response = UrlFetchApp.fetch(profile_url, get_options);
  var user = JSON.parse(response.getContentText());
  return user.displayName;
}

function FindNameinSheet(userId){
  sheet.getRange(6, 6).setValue('UserId');
  var report_num = sheet.getRange(2,7).getValue();
  var i,complete = 0,index = 0;
  for(i = 0; i < report_num && !complete; i++){
    if(sheet.getRange(i+7,6).getValue() == userId){
      complete = 1;
      index = i;
    }
  }
  if(complete){
    return sheet.getRange(index+7,7).getValue();
  }
  else{
    sheet.getRange(6, 6).setValue('error!!');
    return "error";
  }
}

function doPost(e){
  sheet.getRange(1, 1).setValue('時間');
  var msg = JSON.parse(e.postData.contents);
//  Logger.log(msg);
//  var destination = message.destination;
  var events = msg.events[0];
  var follow_event = msg.events[1];
  if (events) {
    sheet.getRange(1, 2).setValue('名字');
    var replyToken =  events.replyToken;
    var type = events.type;
    var timestamp = timeparse(events.timestamp);
    var userId = events.source && events.source.userId;
    if(type === "message"){
      var Name = FindNameinSheet(userId);        
      var userMessage = events.message.text;
      var replytext = Reply(Name,timestamp,userMessage);
      
      var payload = {
        'replyToken': replyToken,
        'messages':[{
          'type': 'text',
          'text': replytext
        }]
      };
      var post_options = {
        'headers': {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer '+ Channel_access_token
        },
        'method': 'post',
        'payload': JSON.stringify(payload)
      };
      var reply_url = 'https://api.line.me/v2/bot/message/reply';
      UrlFetchApp.fetch(reply_url, post_options);
    }
    else if(type === "follow"){
      var Name = FindNameinSheet(userId);
      if(Name == 'error'){
        sheet.getRange(5, 6).setValue('個人資料庫');
        var report_num = sheet.getRange(2,7).getValue();
        var new_row = report_num + 7;
        Name = FindNamebyProfile(userId);
        sheet.getRange(2,7).setValue(report_num+1);
        sheet.getRange(new_row, 6).setValue(userId);
        sheet.getRange(new_row, 7).setValue(Name);
        sheet.getRange(new_row, 8).setValue("=sumif(B:B,\""+Name+"\",C:C)");
      }
    }
  }  
}

function doGet(e){
//  var msg = JSON.parse(e.postData.contents);
//  Logger.log(msg);
//  var events = msg.events[0];
  console.log('Hello world!');
}
