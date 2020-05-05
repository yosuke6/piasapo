
function sendingGroup() {
  var spreadsheet1 = SpreadsheetApp.openById('1_LwvqEM_3Y8WsOUjIbGY3Xjic1-rDZF7AzBDdSN6018');
  var usersInfo   = spreadsheet1.getDataRange().getValues();
  var idColumn = getCol(usersInfo, 0);
  var listSize = idColumn.length;
  
  for(var i=1;i<listSize;i++) {
    var str = usersInfo[i][1] + 'さんは' + usersInfo[i][6] + 'グループです';
    pushMessage(usersInfo[i][0], str);
  } 
  

}


function doPost(e) {
  var replyToken   = JSON.parse(e.postData.contents).events[0].replyToken;
  var userId      = JSON.parse(e.postData.contents).events[0].source.userId;
  var userMessage  = JSON.parse(e.postData.contents).events[0].message.text;
  var spreadsheet1 = SpreadsheetApp.openById('1_LwvqEM_3Y8WsOUjIbGY3Xjic1-rDZF7AzBDdSN6018');
  var spreadsheet2 = SpreadsheetApp.openById('1S64Y0aOqdEN92roSibMdOJgBPcVNYcqwsNy7Y4gTMjc');
  var usersInfo   = spreadsheet1.getDataRange().getValues();
  var stringInfo  = spreadsheet2.getDataRange().getValues();
  
  console.log(userId);
  
  // 現在のユーザーリストの情報を取得
  var idColumn = getCol(usersInfo, 0);
  var listSize = idColumn.length;
  var userIndex = idColumn.indexOf(userId); //名簿内における送信者の位置を検出
  
  var messageStr;
  
  if(userIndex < 0) {//初送信の場合の処理
    setValueFunc(spreadsheet1,listSize,0,userId);
    setValueFunc(spreadsheet1,listSize,1,userMessage);
    messageStr = '名前：' + userMessage;
    messageStr = messageStr + '\n言語を入力してください';
    replyFun(replyToken, messageStr);
    return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    
  var index2=2;
    
  while(!hasNoValue(spreadsheet1,userIndex,index2)) {
    if(index2>4) {
      var messageStr = '登録は完了しています。\n' + userInfoCheck2(usersInfo[userIndex]) + '\n修正したい場合はスタッフにお声かけください。';
      replyFun(replyToken, messageStr);
      return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);}
      index2++;
    } //index2に記入位置を代入
  setValueFunc(spreadsheet1,userIndex,index2,userMessage);
  
  
  switch (index2) {
    case 2:
      messageStr = '言語：' + userMessage;
      messageStr = messageStr + '\n専攻を入力してください';
      break;
    case 3:
      messageStr = '専攻：' + userMessage;
      messageStr = messageStr + '\n学年を入力してください';
      break;
    case 4:
      messageStr = '学年：' + userMessage;
      messageStr = messageStr + '\n国籍を入力してください';
      break;
    case 5:
      messageStr = userInfoCheck(usersInfo[userIndex], userMessage);
//      messageStr = messageStr + '\n受付が完了しました\n登録内容に誤りがある場合はスタッフにお声かけください';
      var random = Math.floor( Math.random() * 5 );
      var groups = ['A','B','C','D','E'];
      var group = groups[random]
      setValueFunc(spreadsheet1,userIndex,6,group);
      messageStr = messageStr + '\n受付が完了しました\n' + usersInfo[userIndex][1] + 'さんは' + group + 'グループです';
      break;
    default:
      break;
  
  
  
  }
  replyFun(replyToken, messageStr);
  

  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}



function userInfoCheck(userInfo, nationality) {
  var str;
  str =         "名前：" + userInfo[1];
  str = str + "\n言語：" + userInfo[2];
  str = str + "\n専攻：" + userInfo[3];
  str = str + "\n学年：" + userInfo[4];
  str = str + "\n国籍：" + nationality;
  return str
}

function userInfoCheck2(userInfo) {
  var str;
  str =         "名前：" + userInfo[1];
  str = str + "\n言語：" + userInfo[2];
  str = str + "\n専攻：" + userInfo[3];
  str = str + "\n学年：" + userInfo[4];
  str = str + "\n国籍：" + userInfo[5];
  str = str + "\nグループ：" + userInfo[6];
  return str
}




//配列の列ベクトル抽出
function getCol(matrix, col){
  var column = [];
  for(var i=0; i<matrix.length; i++){column.push(matrix[i][col]);}
  return column;}



//指定のセルに文字列を書き込み
function setValueFunc(sheet,index1,index2,value) { 
  var currentCell=currentCellFunc(sheet,index1,index2);
  currentCell.setValue(value);}
//カレントセルを指定
function currentCellFunc(sheet,index1,index2) { 
  var currentCell = sheet.getRange("A1"); 
  currentCell=currentCell.offset(index1,index2);
  return currentCell;} 



//該当列の2行目が空白であるか判定
function hasNoValue(sheet,index1,index2) { 
  var currentCell=currentCellFunc(sheet,index1,index2);       
  return currentCell.isBlank();}



function replyFun(replyToken, messageStr) {
  // LINE developersのメッセージ送受信設定に記載のアクセストークン
  var ACCESS_TOKEN = '5Dal3rWSTGmZdedrHVajsbImpSfGGdTxmSbTCDQyAdwaybqbtc5kqTOmuFEm9/YBoYh/g9J/I8FE1GL/jexYPlM8IZuJbEZsqETkPQf3L+uBJ66czGroXc8APQ7BfKRpTyKSsooYxKmQn9/d9aZWLgdB04t89/1O/w1cDnyilFU=';
  // 応答メッセージ用のAPI URL
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
  //    'text': userMessage + 'ンゴ',
        'text': messageStr,
      }],
    }),
    });
}

function pushMessage(USER_ID, messageStr) {
  var ACCESS_TOKEN = '5Dal3rWSTGmZdedrHVajsbImpSfGGdTxmSbTCDQyAdwaybqbtc5kqTOmuFEm9/YBoYh/g9J/I8FE1GL/jexYPlM8IZuJbEZsqETkPQf3L+uBJ66czGroXc8APQ7BfKRpTyKSsooYxKmQn9/d9aZWLgdB04t89/1O/w1cDnyilFU=';
  var url = "https://api.line.me/v2/bot/message/push";
  var postData = {"to": USER_ID,"messages": [{"type": "text","text": messageStr,}]};
  var headers = {"Content-Type": "application/json",'Authorization': 'Bearer ' + ACCESS_TOKEN,};
  var options = {"method": "post","headers": headers,"payload": JSON.stringify(postData)};
  var response = UrlFetchApp.fetch(url, options);
}
