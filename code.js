// Declaration of variables: 
// Read the sheet 
  
  var url = "https://script.google.com/macros/s/AKfycbyx3aOL2z55SDj4g0heeHAM79wzx1GtNHBQEcu2ZkKUHAWzJ7lz/exec";
  var ss = SpreadsheetApp.openById("1VO9ZU0EgMYJF9aCWpZvR3zx1n2pM-K3lQ_Fb9Cdl0ng"); // Open ENYU_DB
  var sheetENYU = ss.getSheetByName("ENYU_CONTACTS");
  var sheetSundaySub = ss.getSheetByName("SUNDAY_MEAL_SUBSCRIPTION"); 
  var ss2 = SpreadsheetApp.openById("1PynUvUxwOe0WDFZ11aw5K06Vnwn8Q4qFm6Z98f7t760");// Open the table
  var sheetCurrentYear = ss2.getSheetByName(2016);
  var sheetCurrentYearArray = sheetCurrentYear.getRange(15, 1, sheetCurrentYear.getLastRow(), 5).getValues();

  var sunday = getSunday(); // get the comming Sunday.
    
  var dateTimeString = sunday+' 01:00:00';
  var deadLine = new Date(dateTimeString);
  
  var startRow = 2;  // First row of data to process
  var numRows = sheetENYU.getLastRow()-1;  // Number of rows to process
  var numCols = 13;
  // Fetch the range of cells A2:B3
  var allData = sheetENYU.getRange(startRow, 1, numRows, numCols);
  var dataArray = allData.getValues();

function sundayMealInvitation(e) {
 
  // Fetch values for each row in the Range.
  var mealDate = replaceAll(getSunday(), "/", "");
  for (var i = 0; i < dataArray.length; ++i) {
    var row = dataArray[i];
    if ((row[10] ==1)&&(row[9]=='active')&&(row[6]!=null)) { // subscription_sunday
     var emailAddress = row[6];  // First column
     var rowId = i;
     var firstName = getFirstName(rowId);
     var personId = row[0];
     var accountLeft = getAccountLeft(personId);
     
     var accountAlert = getAccountAlert(accountLeft);
       
     var participate = url + '?opinion=true' + '&reply_id='+personId + '&reply_rowId='+rowId + '&reply_date='+mealDate;
     var refuse = url + '?opinion=false' + '&reply_id='+personId + '&reply_rowId='+rowId + '&reply_date='+mealDate;
           
     var htmlBodyObj = HtmlService.createTemplateFromFile('mail_template');
     htmlBodyObj.firstName = firstName;
     htmlBodyObj.participate = participate;
     htmlBodyObj.refuse = refuse;
     htmlBodyObj.deadLine = deadLine;
     htmlBodyObj.accountLeft = accountLeft;
     htmlBodyObj.accountAlert = accountAlert;
     
     var htmlBody = htmlBodyObj.evaluate().getContent();
     var subject = sunday+" 周日聚餐统计";
     var message = "请确认是否参加周日聚餐";
      
      MailApp.sendEmail(emailAddress, subject, message, {htmlBody:htmlBody});
      SpreadsheetApp.flush();
    }
         
  }
         
  }

// Responsive html
function doGet(e) {
  
  var replyRowId = e.parameter.reply_rowId;
  var row = dataArray[replyRowId]; 
  var firstName = getFirstName(replyRowId);
  var actualDate = new Date();
  var app = UiApp.createApplication();
  var join =1;
  var sundayStr = replaceAll(getSunday(), "/", "");
  
  if (e.parameter.reply_date == sundayStr){
    if (e.parameter.opinion == 'true') {
      if(actualDate < deadLine) { //Check the date
        if (row[0]==e.parameter.reply_id){ //Check if the response page was manually changed.
          app = responsePage(0,firstName,actualDate);
          insertRecords(row,join);
        /* Record the reservation in the spreadsheet if not yet added.*/
        }
    
        else { app = responsePage(4,firstName,actualDate);}
      } else {app = responsePage(2,firstName,actualDate);}
    
    } else if (e.parameter.opinion == 'false') {
      if(actualDate < deadLine) {
        if (row[0]==e.parameter.reply_id){
          app = responsePage(1,firstName,actualDate);
          join =0;
          insertRecords(row,join);
          // deleteRecords(row,e.parameter.reply_id);
          /* Remove the reservation in the spreadsheet if it is still there.*/
        
        } else { app = responsePage(4,firstName,actualDate);}
      
      } else {app = responsePage(3,firstName,actualDate);}  
          
    }
  } else {app = responsePage(5,firstName,actualDate);}
  
    return app;
}


// ++++++++ Tools ++++++++

function getSunday() {// Get the comming Sunday

  var sunday = new Date();
  var x = sunday.getDay();
  if (x!=0) {
    sunday.setDate(sunday.getDate()+7-x);}
    
  var dd = sunday.getDate();
  var mm = sunday.getMonth()+1; //January is 0!
  var yyyy = sunday.getFullYear();

  if(dd<10) {
      dd='0'+dd
  } 

  if(mm<10) {
      mm='0'+mm
  } 

  var sunday = mm+'/'+dd+'/'+yyyy;
  return sunday;
}


function replaceAll(str, find, replace) {
  return str.replace(new RegExp(find, 'g'), replace);
}

function getFirstName(rowId) {
  var firstName = "";
  var row = dataArray[rowId];
  
  firstName = row[1];
  if (firstName.length==1){
   firstName = row[2]+firstName;
  }
  
  return firstName;

}

function getAccountLeft(personId){
  var isExist = 0;
  var value;
  for (i in sheetCurrentYearArray){
     if (sheetCurrentYearArray[i][0] == personId){ 
        isExist=1;
        value = sheetCurrentYearArray[i][4];
        
        break;
      }
  
  }
  
  if(isExist ==0) {
    value = 10;
  }
 return value;              

}

function getAccountAlert(accountLeft){
  var alertMsg ="";
  if (accountLeft!=-9999){
    if(accountLeft >=4) {
       alertMsg = "，恩雨团契餐饮部感谢您的忠实支持！";
    } else {
       alertMsg = "，亲，您的余额已不足，记得来充值哦！";
    }
  } else {
       alertMsg = "，帐户尚未开通或异常，请联系客服！";
  
  }
  return alertMsg;
}


function responsePage(option,firstName,actualDate) {
  var app = UiApp.createApplication();
  var html ="";
  switch (option) {
      case 0:
      html = '<h2>谢谢! '+ firstName +', 你的报名已确认！ timestamp:'+actualDate+'</h2><br /><br /><br />'+
                             '小提醒：如果你临时有事不能参加，请于 '+ deadLine.toLocaleString() + ' 之前打开邮件点击按钮 NON，否则账户还是会扣款哦！';
          break;
      case 1:
          html = '<h2>'+ firstName +', 很难过你不能来参加我们的聚餐，希望下次光临！ timestamp:'+actualDate+'</h2><br /><br /><br />'+
                             '小提醒：如果你临时计划有变，请于 '+ deadLine.toLocaleString() + ' 之前打开邮件点击按钮 OUI，否则就不能参加了哦！'
          break;
      case 2:
          html = '<h2>'+ firstName +', 报名已截止！</h2><br /><br /><br />'+
                             '小提醒：下次请尽快报名！'
          break;
      case 3:
          html = '<h2>'+ firstName +', 报名已截止，你的订单已不能取消！</h2><br /><br /><br />'+
                             '小提醒：下次请在截止时间之前取消！'
          break;
      case 4:
          html = '<h2>网页有错误，请重新点击！</h2><br /><br /><br />'
          break; 
      case 5:
          html = '<h2>此邮件已过期，请从本周收到的邮件中点击！</h2><br /><br /><br />'
          break;
  }
  app.add(app.createHTML(html));
  return app;
}


function insertRecords(row, join) {
  // Initializing the insert data;
  var personId = row[0];
  var firstName = row[1];
  var lastName = row[2];
  var emailAddress = row[6];
  var mealDate = sunday;
  var timeStamp = new Date();
    
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
     
    var sheetSundaySub = ss.getSheetByName("SUNDAY_MEAL_SUBSCRIPTION"); 
    var sheetSundaySubValues = sheetSundaySub.getDataRange().getValues();
    var isExist = 0;
    var headRow = 1;
      //  var headers = sheetSundaySub.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheetSundaySub.getLastRow()+1; // get next row
    var insertRow = []; 
    
    insertRow.push(personId);
    insertRow.push(firstName);
    insertRow.push(lastName);
    insertRow.push(emailAddress);
    insertRow.push(mealDate);
    insertRow.push(timeStamp);
    insertRow.push(join);
      // loop through the header columns
    
    
    for (i in sheetSundaySubValues){
      if (sheetSundaySubValues[i][0] == personId){ //If the row exists, update the row with new values.
        sheetSundaySub.getRange(parseInt(i)+1, 1, 1, insertRow.length).setValues([insertRow]);
        isExist=1;
        break;
      }
           
    }
     
    if (isExist!=1){ //If the row doesn't exist, append to the end of spreadsheet.
      
      sheetSundaySub.getRange(nextRow, 1, 1, insertRow.length).setValues([insertRow]);
    }
    
    SpreadsheetApp.flush();
      // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
    
    
  } catch(e){
    // if error return this
         return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
    } finally { //release lock
      lock.releaseLock();
  }
  
  
}
