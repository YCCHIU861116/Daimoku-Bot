var Channel_access_token = 'zJAXdKw7A9ZvWJYfFzME4kEWO3lJ/VMco2vgQZVLD/WTmF8th+X51JonNFeXGScrdbFpz7C8F5V3EHjvhEd6KNPgP+5ssrgAygCjajK6msasukLXPdqo/b8zhfTzDITf8sleKtE1ERo3pWrs/5V5RgdB04t89/1O/w1cDnyilFU=';
var Sheet_Id = '1nspvplAZLH3-JhwLNHyYFmEDi1mbi_eR5I8NP-DYVac';
var Test_Sheet_Id = '1-GQ4KZjf1tcELwyGICq1lvOq0iW2jq_k4E69-N-hceY';
var spreadsheet = SpreadsheetApp.openById(Sheet_Id);
var sheet = spreadsheet.getSheets()[0];
var reply_sheet = spreadsheet.getSheets()[1];
var Response_Sheet_Id = '18DMAruCVS4HUe_YrcnPidPbMCMksqUJrV8DdaPzaIcQ';
var res_spreadsheet = SpreadsheetApp.openById(Response_Sheet_Id);
var res_sheet = res_spreadsheet.getSheets()[0];

function doPush(userId,push_messages){
  var payload = {
    'to': userId,
    'messages': push_messages
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

function doReply(replyToken,reply_messages){
  var payload = {
    'replyToken': replyToken,
    'messages': reply_messages
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
  
function alarm(){
  var i;
  var push_messages = [];
  var pushtext = {
      'type': 'text',
      'text': ""
  };
  var report_num = sheet.getRange(2,7).getValue();
  var str = reply_sheet.getRange(36,2).getValue() +sheet.getRange(2,6).getValue()+reply_sheet.getRange(36,3).getValue()+sheet.getRange(2,6).getValue()*50+reply_sheet.getRange(36,4).getValue()+sheet.getRange(4,16).getValue()+reply_sheet.getRange(36,5).getValue();
  for(i = 0; i < report_num; i++){
    var userId = sheet.getRange(i+7,6).getValue();
    var Name = sheet.getRange(i+7,7).getValue();
    pushtext.text = Name + str;
    push_messages.push(pushtext);
    doPush(userId,push_messages);
  }
}

function check_7day(){
  var report_num = sheet.getRange(2,7).getValue();
  var i;
  for(i = 0; i < report_num; i++){
   if(sheet.getRange(i+7,11).getValue() > 7){
    var userId = sheet.getRange(i+7,6).getValue();
    var Name = sheet.getRange(i+7,7).getValue();
    var pushtext = Name + reply_sheet.getRange(44,2).getValue();
    doPush(userId,pushtext);
   }
  }
}

function check_21days_in_a_row(index,Name){
  sheet.getRange(6, 13).setValue('連續天數');
  if(sheet.getRange(index,11).getValue() == 1){
    sheet.getRange(index,13).setValue(sheet.getRange(index,13).getValue()+1);
  }
  else if(sheet.getRange(index,11).getValue() !== 0){
    sheet.getRange(index,13).setValue(1);
  }
  return (sheet.getRange(index,13).getValue() == 21);
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
    return "empty";
  }
}

function FindIndexbyName(Name){
  sheet.getRange(6, 7).setValue('名字1');
  var report_num = sheet.getRange(2,7).getValue();
  var i,complete = 0,index = 0;
  for(i = 0; i < report_num && !complete; i++){
    if(sheet.getRange(i+7,7).getValue() == Name){
      complete = 1;
      index = i;
    }
  }
  if(complete){
    return index+7;
  }
  else{
    //sheet.getRange(6, 6).setValue('error!!');
    return -1;
  }
}

function FindReplyMessage(Name,timestamp, userMessage){
  sheet.getRange(1, 6).setValue('總計');
  var reply_messages=[];
  var replytext = {
    "type": "text",
    "text": ""
  };
  var replyimage = {
      'type': 'image',
      "originalContentUrl": "",
      "previewImageUrl": ""
  };
  var msg =  userMessage.split(' ');
  if(isNaN(msg[0])){
    switch(msg[0]){
      case reply_sheet.getRange(18,2).getValue(): case reply_sheet.getRange(18,3).getValue()://everyone
        var all_daimoku_num = sheet.getRange(2,6).getValue();
        replytext.text = reply_sheet.getRange(19,2).getValue() + all_daimoku_num + reply_sheet.getRange(19,3).getValue()+all_daimoku_num*50+reply_sheet.getRange(19,4).getValue();
        reply_messages.push(replytext);
        break;
      case reply_sheet.getRange(20,2).getValue(): case reply_sheet.getRange(20,3).getValue()://myself
        var user_daimoku_num = find_daimoku_num(Name);
        if(user_daimoku_num == -1)
          replytext.text = "error: can't identify your Name.";
        else if(user_daimoku_num == 0)
          replytext.text = reply_sheet.getRange(21,2).getValue();
        else
          replytext.text = reply_sheet.getRange(22,2).getValue() + user_daimoku_num + reply_sheet.getRange(22,3).getValue()+user_daimoku_num*50+reply_sheet.getRange(22,4).getValue();
        reply_messages.push(replytext);
        break;
      case reply_sheet.getRange(23,2).getValue(): case reply_sheet.getRange(23,3).getValue(): case reply_sheet.getRange(23,4).getValue():
        if(msg.length == 1){
          if(msg[0] == "S") replytext.text = reply_sheet.getRange(24,2).getValue()+sheet.getRange(2,12).getValue()+reply_sheet.getRange(24,5).getValue()+sheet.getRange(2,12).getValue()*50;
          else if(msg[0] == "D") replytext.text = reply_sheet.getRange(24,3).getValue()+sheet.getRange(2,13).getValue()+reply_sheet.getRange(24,5).getValue()+sheet.getRange(2,13).getValue()*50;
          else replytext.text = reply_sheet.getRange(24,4).getValue()+sheet.getRange(2,14).getValue()+reply_sheet.getRange(24,5).getValue()+sheet.getRange(2,14).getValue()*50;
          replytext.text += reply_sheet.getRange(24,6).getValue();
        }
        else if(msg.length == 2){
          var row = FindIndexbyName(Name);
          if(Name !== msg[1]){
            sheet.getRange(row,8).setValue("=sumif(B:B,G"+row+",C:C) + sumif(B:B,L"+row+",C:C)");
          }
          sheet.getRange(row,9).setValue(msg[0]);
          sheet.getRange(row,10).setValue("=MAX(MAXIFS(A:A,B:B,G"+ row +"),MAXIFS(A:A,B:B,L"+ row+"))");
          sheet.getRange(row,11).setValue("=DATEDIF(J"+row +",TODAY()+1,\"D\")-1");
          sheet.getRange(row,12).setValue(msg[1]);
          sheet.getRange(row,13).setValue(0);
          replytext.text = reply_sheet.getRange(3,3).getValue();
        }
        else{
          replytext.text = reply_sheet.getRange(25,2).getValue();
        }
        reply_messages.push(replytext);
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
        replytext.text = reply_sheet.getRange(29,2).getValue();
        reply_messages.push(replytext);
        break;
      case reply_sheet.getRange(32,2).getValue()://祈求目標
        replyimage.originalContentUrl = reply_sheet.getRange(33,2).getValue();
        replyimage.previewImageUrl = reply_sheet.getRange(33,3).getValue();
        reply_messages.push(replyimage);
        break;
      case reply_sheet.getRange(2,2).getValue():case reply_sheet.getRange(2,3).getValue()://使用說明
        replytext.text = reply_sheet.getRange(3,3).getValue();
        reply_messages.push(replytext);
        break;
      case reply_sheet.getRange(39,2).getValue():
        var row = FindIndexbyName(Name);
        var Date = sheet.getRange(row,10).getValue();
        if (Date.getFullYear() === 1899)
          replytext.text = reply_sheet.getRange(41,2).getValue();
        else
          replytext.text = reply_sheet.getRange(40,2).getValue()+Utilities.formatDate(Date, 'Asia/Taipei', 'yyyy/MM/dd, HH:mm:ss Z');        
        reply_messages.push(replytext);
        break;
      default:
        var newresContents = [timestamp,Name,userMessage,1];
        res_sheet.appendRow(newresContents);
        var reply_to_res = find_reply(userMessage);
        if(reply_to_res === 'empty')
          replytext.text = reply_sheet.getRange(2,4).getValue();//Name + 
        else
          replytext.text = reply_to_res;
        reply_messages.push(replytext);
    }
  }
  else{
    sheet.getRange(1, 3).setValue('題目數');
    var daimoku_num = parseInt(msg[0]);
    var row = FindIndexbyName(Name);
    var now = sheet.getRange(2,10).getValue();//new Date();
    var date = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd');
    if(msg[1])
      date = '2020/'+ msg[1];
    var newrowContents = [timestamp,Name,daimoku_num,sheet.getRange(row,9).getValue(),date];
    if(daimoku_num >=100000){
      replytext.text = reply_sheet.getRange(6,2).getValue();
      reply_messages.push(replytext);
      return reply_messages;
    }
    else if(daimoku_num<0 || userMessage.indexOf(".")>-1){
      replytext.text = reply_sheet.getRange(7,2).getValue();
      reply_messages.push(replytext);
      return reply_messages;
    }
    var con21 = check_21days_in_a_row(row,Name);
    sheet.appendRow(newrowContents);
    if(daimoku_num > 90){
      replytext.text = reply_sheet.getRange(8,2).getValue() + Name + reply_sheet.getRange(8,3).getValue() + daimoku_num + reply_sheet.getRange(8,4).getValue();
    }
    else if(daimoku_num > 60){
      replytext.text = reply_sheet.getRange(9,2).getValue() + Name + reply_sheet.getRange(9,3).getValue() + daimoku_num + reply_sheet.getRange(9,4).getValue();
    }
    else if(daimoku_num > 30){
      replytext.text = reply_sheet.getRange(10,2).getValue() + Name + reply_sheet.getRange(10,3).getValue() + daimoku_num + reply_sheet.getRange(10,4).getValue();
    }
    else{
      replytext.text = reply_sheet.getRange(11,2).getValue() + Name + reply_sheet.getRange(11,3).getValue() + daimoku_num + reply_sheet.getRange(11,4).getValue();
    }
    var user_daimoku_num = find_daimoku_num(Name);
    replytext.text += reply_sheet.getRange(12,2).getValue() + user_daimoku_num +reply_sheet.getRange(12,3).getValue()+user_daimoku_num*50 +reply_sheet.getRange(12,4).getValue();
    reply_messages.push(replytext);
    if(con21){
      var replytext2 = {
        "type": "text",
        "text": reply_sheet.getRange(47,2).getValue() + Name + reply_sheet.getRange(47,3).getValue()
      };
      reply_messages.push(replytext2);
    }
  }
  return reply_messages;
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
      var reply_messages = FindReplyMessage(Name,timestamp,userMessage);
      doReply(replyToken,reply_messages);
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
        sheet.getRange(new_row, 8).setValue("=sumif(B:B,G"+new_row+",C:C)");
      }
      var reply_messages = [{
        'type': 'text',
        'text': reply_sheet.getRange(3,2).getValue()
      }]
      doReply(replyToken,reply_messages);
    }
  }  
}

function doGet(e){
//  var msg = JSON.parse(e.postData.contents);
//  Logger.log(msg);
//  var events = msg.events[0];
  console.log('Hello world!');
}
