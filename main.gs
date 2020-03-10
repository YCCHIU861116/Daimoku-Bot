var Channel_access_token = 'zJAXdKw7A9ZvWJYfFzME4kEWO3lJ/VMco2vgQZVLD/WTmF8th+X51JonNFeXGScrdbFpz7C8F5V3EHjvhEd6KNPgP+5ssrgAygCjajK6msasukLXPdqo/b8zhfTzDITf8sleKtE1ERo3pWrs/5V5RgdB04t89/1O/w1cDnyilFU=';
var Sheet_Id = '1nspvplAZLH3-JhwLNHyYFmEDi1mbi_eR5I8NP-DYVac';
var Test_Sheet_Id = '1-GQ4KZjf1tcELwyGICq1lvOq0iW2jq_k4E69-N-hceY';
var spreadsheet = SpreadsheetApp.openById(Sheet_Id);
var sheet = spreadsheet.getSheets()[0];
var reply_sheet = spreadsheet.getSheets()[1];
var Response_Sheet_Id = '18DMAruCVS4HUe_YrcnPidPbMCMksqUJrV8DdaPzaIcQ';
var res_spreadsheet = SpreadsheetApp.openById(Response_Sheet_Id);
var res_sheet = res_spreadsheet.getSheets()[0];

function alarm(){
  var i;
  var report_num = sheet.getRange(2,7).getValue();
  for(i = 0; i < report_num; i++){
    var userId = sheet.getRange(i+7,6).getValue();
    var Name = sheet.getRange(i+7,7).getValue();
    var pushtext = Name + reply_sheet.getRange(36,2).getValue() +sheet.getRange(2,6).getValue()+reply_sheet.getRange(36,3).getValue();
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
    /*if(res)
      sheet.getRange(2,5).setValue(res.getContentText());*/
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
  if(isNaN(msg[0])){
    switch(msg[0]){
      case reply_sheet.getRange(18,2).getValue(): case reply_sheet.getRange(18,3).getValue()://everyone
        var all_daimoku_num = sheet.getRange(2,6).getValue();
        replytext = reply_sheet.getRange(19,2).getValue() + all_daimoku_num + reply_sheet.getRange(19,3).getValue();
        break;
      case reply_sheet.getRange(20,2).getValue(): case reply_sheet.getRange(20,3).getValue()://myself
        var user_daimoku_num = find_daimoku_num(Name);
        if(user_daimoku_num == 0)
          replytext = reply_sheet.getRange(21,2).getValue();
        else
          replytext = reply_sheet.getRange(22,2).getValue() + user_daimoku_num + reply_sheet.getRange(22,3).getValue();
        break;
      case reply_sheet.getRange(28,2).getValue(): case reply_sheet.getRange(28,3).getValue()://Undo
        var lastRow = sheet.getLastRow();
        var i,complete = 0;
        for(i = lastRow; i > 0 && !complete; i--){
          if(sheet.getRange(i,2).getValue() == Name){
            sheet.getRange(i,3).setValue(msg[1]);
            complete = 1; 
          }
        }
        replytext = reply_sheet.getRange(29,2).getValue();
        break;
      case reply_sheet.getRange(32,2).getValue()://祈求目標
        replytext = reply_sheet.getRange(33,2).getValue();
        break;
      case reply_sheet.getRange(2,2).getValue()://使用說明
        replytext = reply_sheet.getRange(3,2).getValue();
        break;
      default:
        var newresContents = [timestamp,Name,userMessage,1];
        res_sheet.appendRow(newresContents);
        var reply_to_res = find_reply(userMessage);
        if(reply_to_res === 'empty')
          replytext = reply_sheet.getRange(2,4).getValue();//Name + 
        else
          replytext = reply_to_res;
    }
  }
  else{
    sheet.getRange(1, 3).setValue('題目數');
    var daimoku_num = parseInt(userMessage);
    var newrowContents = [timestamp,Name,daimoku_num,"G","=TODAY()"];
    if(daimoku_num >=100000){
      return reply_sheet.getRange(6,2).getValue();
    }
    else if(daimoku_num<0 || userMessage.indexOf(".")>-1){
      return reply_sheet.getRange(7,2).getValue();
    }
    sheet.appendRow(newrowContents);
    if(daimoku_num >= 60){
      replytext = Name + reply_sheet.getRange(8,2).getValue();
      if(daimoku_num == 60){
        replytext += '一小時';
      }
      else{
        replytext += (daimoku_num + '分鐘');
      }
      replytext += reply_sheet.getRange(8,3).getValue();
    }
    else{
      replytext = reply_sheet.getRange(9,2).getValue();
    }
    var user_daimoku_num = find_daimoku_num(Name);
    replytext += reply_sheet.getRange(12,2).getValue() + user_daimoku_num +reply_sheet.getRange(12,3).getValue();
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
    //sheet.getRange(6, 6).setValue('error!!');
    return "can not find";
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
      if(Name == "can not find"){
        sheet.getRange(5, 6).setValue('個人資料庫');
        var report_num = sheet.getRange(2,7).getValue();
        var new_row = report_num + 7;
        Name = FindNamebyProfile(userId);
        sheet.getRange(2,7).setValue(report_num+1);
        sheet.getRange(new_row, 6).setValue(userId);
        sheet.getRange(new_row, 7).setValue(Name);
        sheet.getRange(new_row, 8).setValue("=sumif(B:B,\""+Name+"\",C:C)");
      }
      var replytext = reply_sheet.getRange(3,2).getValue();
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
  }  
}

function doGet(e){
//  var msg = JSON.parse(e.postData.contents);
//  Logger.log(msg);
//  var events = msg.events[0];
  console.log('Hello world!');
}
