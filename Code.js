// Get the spreadsheet object.
// https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Getting relevant sheets from the spreadsheet.
// https://developers.google.com/apps-script/reference/spreadsheet/sheet
const intakeSheet = spreadsheet.getSheetByName('New Intake');
const deliverySheet = spreadsheet.getSheetByName('Daily Delivery');
const completedSheet = spreadsheet.getSheetByName('Closed/Completed');
const tecSheet = spreadsheet.getSheetByName('TEC Copy - do not edit');

// Get the ui object.
// https://developers.google.com/apps-script/reference/base/ui.html
const ui = SpreadsheetApp.getUi();

const weekday = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Status values for easy reference.
const status = {
  URGENT: 'Needs Urgent',
  PENDING: 'Pending',
  SCHEDULED: 'Scheduled',
  DELIVERED: 'Delivered',
};

// Returns the column index number by the name and sheet name
function getColByName(name, sheet) {
  return sheet.getDataRange().getValues().shift().indexOf(name)++;
}

// Pull out the indices of various columns of interest.
const deliveryUidCol = getColByName('UID', deliverySheet);
const intakeUidCol = getColByName('UID', intakeSheet);
const deliveryStatusCol = getColByName('Status Update', deliverySheet);
const intakeStatusCol = getColByName('Status', intakeSheet);

// upon opening the script will add a menu button to the document
function onOpen() {
  ui.createMenu('Daily Delivery Automation')
    .addItem('Run', 'confirmStart').addToUi();
};

// if the DD Automation button is selected and the 'Run DD Automation' YES
// button is pressed, DD automation will pop up and process will begin
function confirmStart() {
  let buttonPressed = ui.alert(
    'Please update status for all deliveries in daily delivery sheet prior. Begin daily delivery automation?',
    ui.ButtonSet.YES_NO
  );
  if (buttonPressed == ui.Button.YES) {
    runAutomation();
  }
};

//function getRunners() {
//  let runnersofDay = ui.prompt('Please enter the number of runners for the day', ui.ButtonSet.OK_CANCEL);
//  let button = runnersofDay.getSelectedButton();
//  if(button == ui.Button.OK) {
//    //saves the number of RUNNERS for use within this function
//    return runnersofDay.getResponseText();
//  }
//}

// Main automation process.
function runAutomation() {
  // Get all UIDs currently in Delivery Sheet.
  let deliveryUids = getDeliveryUids();
  highlightScheduledRows(deliveryUids);
  highlightDeliveredRows();
  processDelivered();
  getNextDayRows();
  sortNI();
}

function getDeliveryUids() {
  return deliverySheet
    .getSheetValues(2, deliveryUidCol, deliverySheet.getLastRow(), 1)
    .map(d => d.shift());
}

// Mark rows in Delivery sheet as scheduled in intake sheet.
function highlightScheduledRows(deliveryUids) {
  for (let p = 2; p < intakeSheet.getLastRow(); p++) {
    for (deliveryUid of deliveryUids){
      if (intakeSheet.getRange(p, intakeUidCol).getValue() == deliveryUid){
        intakeSheet.getRange(p, intakeStatusCol).setBackground('#f163d6'); //#c27ba0
        intakeSheet.getRange(p, intakeStatusCol).setValue(status.SCHEDULED);
      }
    }
  }
}

function highlightDeliveredRows() {
  let deliveredUids = [];
  let deliveredRowIndices = [];
  let intakeDeliveryCol = getColByName('Delivery date', intakeSheet)
  let now = new Date();

  ui.alert('Marks all delivered orders as delivered in NI sheet and adds delivered date');
  //adds the dd rows that are marked delivered into two arrays - one of UIDs and one of row numbers              
  for (let d = 2; d < deliverySheet.getLastRow(); d++) {
    if (deliverySheet.getRange(d, deliveryStatusCol).getValue() == status.DELIVERED){
      deliveredUids.push(deliverySheet.getRange(d, deliveryUidCol).getValue());
      deliveredRowIndices.push(d);
    }
  }
  ui.alert(`deliveredUids: ${deliveredUids} \n deliveredRowIndices: ${deliveredRowIndices}`);

  //highlights the rows in New Intake that are delivered in DD, changes them to Delivered and adds a delivered date 
  for (n = 1; n < intakeSheet.getLastRow(); n++) {
    for (let deliveredUID of deliveredUids){
      if (intakeSheet.getRange(n, intakeUidCol).getValue() == deliveredUID){
        intakeSheet.getRange(n, intakeStatusCol).setBackground('#93c47d');
        intakeSheet.getRange(n, intakeStatusCol).setValue(status.DELIVERED);
        intakeSheet.getRange(n, intakeDeliveryCol).setValue(now);
      }
    }
  }
  ui.alert('Deleting delivered rows from DD sheet');
  //deletes the delivered rows from the DD Sheet
  for (let deliveredRow of deliveredRowIndices.reverse()){
    deliverySheet.deleteRow(deliveredRow);
  }
}

function processDelivered() {
  let rowNumsToDelete = [];
  ui.alert('Moving all closed/completed rows from NI to Closed/Completed and deleting from NI');
  //if a row is marked delivered or closed in NI sheet it will move that entire row to the closed/completed tab and delete those rows
  for (j = 2; j < intakeSheet.getMaxRows(); j++){
    if(intakeSheet.getRange(j, intakeStatusCol).getValue() == status.DELIVERED || intakeSheet.getRange(j, intakeStatusCol).getValue() == 'Closed'){
      // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
      completedSheet.appendRow(intakeSheet.getRange(j, 1, 1, intakeSheet.getLastColumn()).getValues()[0]);
      rowNumsToDelete.push(j);
    }
  }
  // Delete rows from the bottom up so that you don't change row indices as you're iterating
  // Caution, reverse() changes the actual contents of rowNumsToDelete!
  for (let rowNum of rowNumsToDelete.reverse()){
    intakeSheet.deleteRow(rowNum);
  }
}

function getNextDayRows() {
  //days of the week are returned as numbers, so this  will return the string for day of week
  let tomorrowday = weekday[now.getDay()+1];
  let tecDAY = getColByName('Status', tecSheet);
  ui.alert('Adding in TEC rows for delivery tomorrow into DD sheet');
  //loop that copies over the TEC row and if the status equals tomorrow's day of the week, copies that into the DD sheet
  for (h = 2; h < tecSheet.getLastRow(); h++) {
    if(tecSheet.getRange(h, tecDAY).getValue() == tomorrowday){
      deliverySheet.appendRow(tecSheet.getRange(h, 1, 1, tecSheet.getLastColumn()).getValues()[0]);

    }
  }
}

function sortNI() {
  ui.alert('Sorting NI sheet by date and then time');
  //sorts the SSNI sheet by date first and then time
  let intakeSheetdate = getColByName('Date', intakeSheet);
  let intakeSheettime = getColByName('Time', intakeSheet);
  intakeSheet.getRange(2, 1, intakeSheet.getLastRow(), intakeSheet.getLastColumn()).sort([intakeSheetdate, intakeSheettime]);

}

//  for (y = 2; y < intakeSheet.getMaxRows(); y++){
//    let intakeStatusColValue = intakeSheet.getRange(y, intakeStatusCol).getValue()
//    
//    if (intakeStatusColValue === 'Needs Urgent') {
//      // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
//        deliverySheet.appendRow(intakeSheet.getRange(y, 1, 1, intakeSheet.getMaxColumns()).getValues()[0]);
//    } else if (intakeStatusColValue === 'Pending') {
//      deliverySheet.appendRow(intakeSheet.getRange(y, 1, 1, intakeSheet.getMaxColumns()).getValues()[0]);
//    }
//  }   
//
//  for (y = 2; y < intakeSheet.getMaxRows(); y++){
//    if(intakeSheet.getRange(y, intakeStatusCol).getValue() == "Needs Urgent"){
//      while (deliverySheet.getLastRow() < totalrows){
//      // the [0] bit is changing a Matrix to a list because getRange().getValues returns a section of the sheet not a row
//        deliverySheet.appendRow(intakeSheet.getRange(y, 1, 1, intakeSheet.getMaxColumns()).getValues()[0]);
//   }
//    }
//  }
//}
