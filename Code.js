//most up to date code 
//variables for the NI and DD sheets
var ssNI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Intake");
var ssDD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Delivery");
var ssCC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Closed/Completed");
var ssTEC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TEC Copy - do not edit");


//upon opening the script will add a menu button to the document
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('DD Automation')
  .addItem('Run DD Automation', 'startUpMessage').addToUi();
};

//Returns the column index number by the name and sheet name
function getColByName(name, sheetname){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var headers = sheet.getDataRange().getValues().shift();
  var colindex = headers.indexOf(name);
  return colindex+1;
};

//if the DD Automation button is selected and the 'Run DD Automation' YES button is pressed, DD automation will pop up and process will begin
function startUpMessage() {
  //this message will pop up once you click on DD automation
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Please update status for all deliveries in DD sheet prior. Begin DD automation?",ui.ButtonSet.YES_NO);
  if(buttonPressed == ui.Button.YES){
    DDAutomation() 
  }
};
//DD AUTOMATION SCRIPT
function DDAutomation() {
  var ssNI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Intake");
  var ssDD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Delivery");
  var ssCC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Closed/Completed");
  var ssTEC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TEC Copy - do not edit");

  
  //prompts user to enter the number of runners for the day
  var ui = SpreadsheetApp.getUi();
  var runnersofDay = ui.prompt('Please enter the number of runners for the day', ui.ButtonSet.OK_CANCEL);
  var button = runnersofDay.getSelectedButton();
  if(button == ui.Button.OK) {
    //saves the number of RUNNERS for use within this function
    var runners = runnersofDay.getResponseText();
  }
  var now = new Date();
  var ddUID = getColByName('UID', 'Daily Delivery');
  var niUID = getColByName('UID', 'New Intake');
  var ddStatCol = getColByName('Status Update', 'Daily Delivery');
  var niStatCol = getColByName('Status', 'New Intake');
  var ddRows = [];
  var ddRows2 = [];
  var ddRows3 = []; //var x
  var urgent = []
  var pending = []
  var rowNumsToDelete = [];
  var deliveredUID;
  var deliveredRow;
  var ddUIDs;
  var totalrows = runners*4;
  var rowsToImport = (totalrows - ssDD.getLastRow());
  var weekday = new Array(7);
  weekday[0] = "Sunday";
  weekday[1] = "Monday";
  weekday[2] = "Tuesday";
  weekday[3] = "Wednesday";
  weekday[4] = "Thursday";
  weekday[5] = "Friday";
  weekday[6] = "Saturday";
  var urgentRow;
  var pendingRow;

  //adds all UID numbers from DD into an array             
  for (x = 2; x < ssDD.getLastRow(); x++) {
        ddRows3.push(ssDD.getRange(x, 1).getValue());
      }

      
  //highlights all orders that are DD form in the NI form and marks as 'scheduled'
  for (p = 1; p < ssNI.getLastRow(); p++) {
    for (ddUIDs of ddRows3){
      if (ssNI.getRange(p, niUID).getValue() == ddUIDs){
        ssNI.getRange(p, niStatCol).setBackground("#c27ba0");
        ssNI.getRange(p, niStatCol).setValue('Scheduled');
      }
    }
 }
         }
//  //adds the dd rows that are marked delivered into two arrays - one of UIDs and one of row numbers              
//  for (d = 1; d < ssDD.getLastRow(); d++) {
//      if(ssDD.getRange(d, ddStatCol).getValue() == "Delivered"){
//        ddRows.push(ssDD.getRange(d, ddUID).getValue());
//        ddRows2.push(d);
//      }
//     }
//  //highlights the rows in New Intake that are delivered in DD, changes them to Delivered and adds a delivered date
//  for (n = 1; n < ssNI.getLastRow(); n++) {
//    for (deliveredUID of ddRows){
//      if (ssNI.getRange(n, niUID).getValue() == deliveredUID){
//        ssNI.getRange(n, niStatCol).setBackground("#93c47d");
//        ssNI.getRange(n, niStatCol).setValue('Delivered');
//        ssNI.getRange(n, getColByName('Delivery date', 'New Intake')).setValue(now);
//      }
//    }
//   }
//  //deletes the delivered rows from the DD Sheet
//  for (deliveredRow of ddRows2.reverse()){
//    ssDD.deleteRow(deliveredRow);
//   }
//
//  
//  //if a row is marked delivered or closed in NI sheet it will move that entire row to the closed/completed tab and delete those rows
//  for (j = 2; j < ssNI.getMaxRows(); j++){
//    if(ssNI.getRange(j, niStatCol).getValue() == "Delivered" || ssNI.getRange(j, niStatCol).getValue() == "Closed"){
//      // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
//      ssCC.appendRow(ssNI.getRange(j, 1, 1, ssNI.getLastColumn()).getValues()[0]);
//      rowNumsToDelete.push(j);
//    }
//  }
//
//  // Delete rows from the bottom up so that you don't change row indices as you're iterating
//  // Caution, reverse() changes the actual contents of rowNumsToDelete!
//  //deletes the rows from the SSNI sheet that are closed/completed and moved to that tab
//  var rowNum;
//  for (rowNum of rowNumsToDelete.reverse()){
//    ssNI.deleteRow(rowNum);
//  }
//   //sorts the SSNI sheet by date first and then time
//   var ssNIdate = getColByName('Date', 'New Intake');
//   var ssNItime = getColByName('Time', 'New Intake');
//   ssNI.getRange(2, 1, ssNI.getLastRow(), ssNI.getLastColumn()).sort([ssNIdate, ssNItime]);
//
//  //days of the week are returned as numbers, so this  will return the string for day of week
//  var dayofweek = weekday[now.getDay()];
//  var tomorrowday = weekday[now.getDay()+1];
//  //loop that copies over the TEC row and if the status equals tomorrow's day of the week, copies that into the DD sheet
//  for (h = 2; h < ssTEC.getLastRow(); h++) {
//      if(ssTEC.getRange(h, 6).getValue() == tomorrowday){
//        ssDD.appendRow(ssTEC.getRange(h, 1, 1, ssTEC.getLastColumn()).getValues()[0]);
//      }
//     }
// //the new intake form is sorted by date and time, so as this loop flips through for needs urgent or pending orders, it will take those first that are later in date
//  for (y = 2; y < ssNI.getMaxRows(); y++){
//    if(ssNI.getRange(y, niStatCol).getValue() == "Needs Urgent"){
//      // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
//        ssDD.appendRow(ssNI.getRange(y, 1, 1, ssNI.getMaxColumns()).getValues()[0]);
//   }
//    }   
//
//  for (c = 2; c < ssNI.getMaxRows(); c++){
//    if(ssNI.getRange(c, niStatCol).getValue() == "Pending"){
//        ssDD.appendRow(ssNI.getRange(c, 1, 1, ssNI.getMaxColumns()).getValues()[0]);
//      ;
//    }
//  }
//}