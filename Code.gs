/***********************************************************************
MIT License

Copyright (c) 2018 daubedesign

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
************************************************************************/

//===========================================================================GLOBAL REFERENCES
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Events');
var documentProperties = PropertiesService.getDocumentProperties();
var calendar = documentProperties.getProperty('calendar');
var calendarName = documentProperties.getProperty('calendarName');

var FY = "=if(B2=\"\",\"\",if(OR(left(B2,2)=\"01\",left(B2,2)=\"02\"),right(B2,4),right(B2,4) + 1))";
var QTR = "=if(B2=\"\",\"\",switch(WEEKDAY(B2),1,\"Sun\",2,\"Mon\",3,\"Tues\",4,\"Wed\",5,\"Thu\",6,\"Fri\",7,\"Sat\"))";
var DAY = "=SWITCH(left(B2,2),\"09\",3,\"10\",3,\"11\",3,\"12\",4,\"01\",4,\"02\",4,\"03\",1,\"04\",1,\"05\",1,\"06\",2,\"07\",2,\"08\",2,\"\",\"\",\"Error\")";

var daubeCommCal = getCalendar();


//===========================================================================SETUP PROCEDURE
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Show Sidebar', 'showSidebar')
      .addItem('Create Event Template', 'createCalendarTemplate')
      .addSeparator()
      .addItem('Configuration', 'showConfigMenu')
      .addToUi();
      showSidebar();
  if (sheet) {
    var range = sheet.getRange("A:Y");
    range.sort(2);
  }
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Calendar Synchronization');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showConfigMenu() {
  var menu = HtmlService.createTemplateFromFile('Menu').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  SpreadsheetApp.getUi().showModelessDialog(menu, 'Calendar Synchronization - Configuration Options');
}

//===========================================================================MAIN FUNCTIONS
//=========================================================================================GET CALENDAR EVENTS FROM GOOGLE

//gets calendar events for the next year and inserts new ones in the sheet
function refreshEvents() {
  var events = getEventsFromGoogle();
  var eventsFormatted = getCalendarEventDetails(events);
  var eventsFormattedAndNew = filterNewEvents(eventsFormatted);
  putNewEventsOnSheet(eventsFormattedAndNew);
  var range = sheet.getRange("A:Y");
  range.sort(2);
}

//get calendar events for today plus one year
function getEventsFromGoogle() {
  var now = new Date();
  var then = new Date(now.getTime() + (3.154e+10));
  var events = daubeCommCal.getEvents(now, then);
  return events;
}

//parse events into the desired fields
function getCalendarEventDetails(events) {
  var eventsFormatted = [];
  for (var i=0;i<events.length;i++) {
    var eventFormatted = {};
    var startTime = events[i].getStartTime();
    var description = events[i].getDescription();
    var title = events[i].getTitle();
    var id = events[i].getId();
    eventFormatted = {
      "startTime": startTime,
      "title": title,
      "description": description,
      "id": id
    };
    eventsFormatted.push(eventFormatted);  
  }
  return eventsFormatted;
}

//get all existing event ids from the sheet
function getEventIdsFromSheet() {
  
  var lastRow = sheet.getLastRow();
  var sheetEvents = sheet.getRange(2, 1, lastRow, 1);
  var EventIds = sheetEvents.getValues();
  return EventIds
}

//filter events to just new events
function filterNewEvents(eN) {
  var newEvents = [];
  var evIds = getEventIdsFromSheet();
  var stringIds = evIds.toString();
  for (var i=0;i<eN.length;i++) {
    var con = stringIds.indexOf(eN[i].id) 
    if (con == -1) {
      newEvents.push(eN[i]);
    } else {
      Logger.log(eN[i].id);
    }
  }
  return newEvents;
}


// jeff you were about to duplicate the formulas to paste to many rows HERE HERE HERE
// append the new events to the sheet
function putNewEventsOnSheet(eF) {
  var newEventCount = eF.length;
  if (newEventCount === 0) {
    return;
  }
  Logger.log('new event count: ' + newEventCount);
  var lastRow = sheet.getRange('D:D').getValues().filter(String).length;
  Logger.log('last row: ' + lastRow);
  var newRow = sheet.insertRowsAfter(lastRow, newEventCount);
  setFormulas()
  var newEventsArray = []
  for (var i = 0; i < newEventCount; i++) {
    var tempArray = [
      eF[i].id,
      eF[i].startTime,
      "Published",
      eF[i].title,
      eF[i].description
    ];
    newEventsArray.push(tempArray);
  }
  var newEventsRange = sheet.getRange('A' + (lastRow + 1) + ':E' + (lastRow + newEventCount));
  newEventsRange.setValues(newEventsArray);
}

//=========================================================================================PUBLISH NEW SHEET EVENTS TO GOOGLE CALENDAR

function publishNewEvent(pubRowNum) {
  var a1 = 'A' + pubRowNum + ':' + 'E' + pubRowNum;
  var data = sheet.getRange(a1).getValues();
  Logger.log(data);
  var eventsToPublish = [];
  var eventToPub = {}
    var id = data[0][0];
    var startDate = data[0][1];
    var topic = data[0][3];
    var details = data[0][4];
    var rowNum = pubRowNum;
    eventToPub = {
      "id": id,
      "startDate": startDate,
      "topic": topic,
      "details": details,
      "rowNum": rowNum
    };
    eventsToPublish.push(eventToPub);
    var event = createEventOnCalendar(eventsToPublish[0]);
    var id = event.getId();
    var row = eventsToPublish[0].rowNum;
    var idCell = sheet.getRange('A' + row);
    idCell.setValue(id);
    var status = sheet.getRange('C' + row);
    status.setValue('Published');
}

function publishNewEvents() {
  var data = getEventsFromSheet();
  Logger.log(data);
  var eventsToPublish = filterNonPublishedEvents(data);
  for (i=0; i<eventsToPublish.length; i++) {
    var event = createEventOnCalendar(eventsToPublish[i]);
    var id = event.getId();
    var row = eventsToPublish[i].rowNum;
    var idCell = sheet.getRange('A' + row);
    idCell.setValue(id);
    var status = sheet.getRange('C' + row);
    status.setValue('Published');    
  }
}

function getEventsFromSheet() {
  var lastRow = sheet.getLastRow() - 1;
  
  var sheetEvents = sheet.getRange(2, 1, lastRow, 5);
  var data = sheetEvents.getValues();
  return data;
}

function filterNonPublishedEvents(data) {
  var eventsToPublish = [];
  var lastRow = sheet.getLastRow() - 1;
  for (var i=0; i<lastRow; i++) {
    var status = data[i][2];
    if (status === 'Published') { continue; };
    
    var eventToPub = {}
    var id = data[i][0];
    var startDate = data[i][1];
    var topic = data[i][3];
    var details = data[i][4];
    var rowNum = i + 2;
    eventToPub = {
      "id": id,
      "startDate": startDate,
      "topic": topic,
      "details": details,
      "rowNum": rowNum
    };
    eventsToPublish.push(eventToPub);
  }
  return eventsToPublish;  
}
  
function createEventOnCalendar(ev) {
  var event = getCalendar().createAllDayEvent(
    ev.topic,
    new Date(ev.startDate),
    {description: ev.details}
  );   
  return event;                                                                                                                     
}

//=========================================================================================UPDATE CHANGED SHEET EVENTS TO CALENDAR

function updateCalendarEvent(upRowNum) {
  var eventData = sheet.getRange(upRowNum, 1, 1, 5).getValues();
  Logger.log(eventData[0][2])
  if (eventData[0][2] === 'Published') {
    displayToast( 'You cannot update an event that has a Published status.' );
  }
  
  var eventToUpdate = daubeCommCal.getEventSeriesById(eventData[0][0]);
  var sDate = new Date(eventData[0][1]);
  eventToUpdate.setTitle(eventData[0][3]);
  eventToUpdate.setDescription(eventData[0][4]);
  var recurrence = CalendarApp.newRecurrence().addDailyRule().times(1);
  eventToUpdate.setRecurrence(recurrence, sDate);
  var status = sheet.getRange('C' + upRowNum);
  status.setValue('Published');    
}

/**************** DELETE EVENT FROM CALENDAR AND SHEET *************/
function deleteEvent(delRowNum) {
  var deleteId = sheet.getRange(delRowNum, 1).getValue();
  Logger.log('row number: ' + delRowNum + 'id: ' + deleteId);
  var cal = getCalendar();
  var event = cal.getEventSeriesById(deleteId);
  event.deleteEventSeries();
  sheet.deleteRow(delRowNum);  
}

/*************** TOGGLE DETAILS ON SHEET ***************************/

function toggleDetails(state) {
  if (state) {
    sheet.showColumns(6, 20)
    sheet.getRange('Z1').setValue('Details On, click checkbox to hide ---------------->')
  } else {
    sheet.hideColumns(6, 20)
    sheet.getRange('Z1').setValue('Details Off, click checkbox to show ---------------->')
  }
}

/**** GENERAL SHEET EDIT DETECTION ****/
function onEdit(e){
  var range = e.range;
  var editColumn = range.getColumn();
  var editSheet = range.getSheet();
  var editSheetName = editSheet.getName();
  var anImportantChange = anImportantFieldChanged(editColumn, editSheetName);
  var editRow = range.getRow();
  var status = sheet.getRange('C' + editRow);
  var statusValue = status.getValue();
  if (anImportantChange) {
    if (statusValue == 'Published') {
      status.setValue('Update calendar');
    }
  }
}

//=========================================================================================UPDATE USER PROPERTIES

function updateCalendarId(cId) {
  if (cId === null) {
    displayToast('You must enter a Google Calendar ID');
  } else {
    calendar = cId
    documentProperties.setProperty('calendar', cId)
  }
}

function updateCalendarName(cName) {
  if (cName === null) {
    displayToast('You must enter a Google Calendar Name');
  } else {
    calendarName = cName
    documentProperties.setProperty('calendarName', cName)
  }
}

/**************************HELPER FUNCTIONS ********************************/
function getCalendar() {
  if (calendar === null) {
    displayToast('You must connect a calendar to this sheet.');
  } else {
    return CalendarApp.getCalendarById(calendar);
  }
}

function displayToast(m) {
  ss.toast(m);
}

function showAlert(type, prompt, message) {

  var ui = SpreadsheetApp.getUi(); // Same variations.
  if (type == 'delete') {
    var result = ui.alert(
      prompt,
      message,
      ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (result == ui.Button.YES) {
      var userResponse = "yes";
    } else {
      var userResponse = "no";
    }
    return userResponse;
  }
}

function anImportantFieldChanged(col, editSheetName) {
  if ((col == '2.0' || col == '4.0' || col == '5.0') && (editSheetName == 'Events')) {
    return true;  
  } else {
    return false;  
  }
}

function getCalendarIdAndName() {
  var calDetails = {
    id: calendar,
    name: calendarName
  }
  return calDetails
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function setFormulas() {
  var sheet = ss.getSheetByName('Events');
  var maxRow = sheet.getMaxRows();
  var f = sheet.getRange('F2:F' + maxRow);
  var g = sheet.getRange('G2:G' + maxRow);
  var h = sheet.getRange('H2:H' + maxRow);
  f.setFormula(FY);
  g.setFormula(QTR);
  h.setFormula(DAY);
}

//===================================================================================CREATE CALENDAR TEMPLATE

function createCalendarTemplate() {
  var sheet1 = ss.getSheetByName('Sheet1');
  var eventSheet = ss.insertSheet('Events');
  ss.deleteSheet(sheet1);
  eventSheet.getRange(1, 1, 1, 8)
      .setValues([['id', 'Date', 'Status', 'Topic', 'Details', 'FY', 'Qtr', 'Day']])
      .setFontWeight('bold')
  eventSheet.setColumnWidth(4, 338)
      .setColumnWidth(5, 374)
      .setColumnWidth(6, 45)
      .setColumnWidth(7, 45)
      .setColumnWidth(8, 45)
      .setColumnWidth(26, 300)
      .hideColumns(1);
  eventSheet.setFrozenRows(1);
  eventSheet.setFrozenColumns(4);
  var dateColumn = eventSheet.getRange('B2:B1000');
  var rule = SpreadsheetApp.newDataValidation().requireDate().build();
  dateColumn.setNumberFormat("mm-dd-yyyy")
      .setDataValidation(rule);
  eventSheet.getRange('A:A').setBackground('#F3F3F3')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  
  eventSheet.getRange('C:C').setBackground('#F3F3F3')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  
  eventSheet.getRange('F:H').setBackground('#F3F3F3')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  
  eventSheet.getRange('D:D').setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  
  setFormulas();
  eventSheet.getRange('Z1').setValue('Details On, click checkbox to hide ---------------->')
      .setBackground('#d9ead3');
  
  eventSheet.getRange('B2:E2').setValues([[
    new Date(),'','Example Calendar Event Title', 'Unpublished calendar event description.'
  ]]);
}