var CALENDAR_ID = "REPLACE_ME",
    HEADER_ROWS = 1,
    DATE = 0,
    DATE_COL = "A",
    EVENT_ID = 4;

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  addMenuItem(sheet);
  selectMostRecentDay(sheet);
};

function addMenuItem(sheet) {
  var entries = [{
    name : "Export New Rotations",
    functionName : "exportEvents"
  }];
  sheet.addMenu("Custom Actions", entries);
}

function selectMostRecentDay(sheet) {
  var data = sheet.getDataRange().getValues(),
      today = new Date();

  for (i=0; i<data.length; i++) {
    var row = data[i],
        date = new Date(row[DATE]);

    if (date.valueOf() > today.valueOf()) {
      sheet.setActiveSelection(sheet.getRange(DATE_COL + i));
      return;
    }
  }
}

/**
 * Export events from spreadsheet to calendar
 */
function exportEvents() {
  var sheet = SpreadsheetApp.getActiveSheet(),
      range = sheet.getDataRange(),
      data = range.getValues(),
      calId = CALENDAR_ID,
      calendar = CalendarApp.getCalendarById(calId);
  
  exportEvent(calendar, data);
  
  range.setValues(data);
}


function exportEvent(calendar, data) {
    for (i=0; i<data.length; i++) {
      var row = data[i];
      
      if (i < HEADER_ROWS) continue; // Skip header row(s)
      if (row[EVENT_ID] != "") continue; // Skip exported events
      
      var date = new Date(row[DATE]),
          title = row[1],           // Title = B
          tstart = new Date(row[2]); // Start Time = C

      tstart.setDate(date.getDate());
      tstart.setMonth(date.getMonth());
      tstart.setYear(date.getYear());

      var tstop = new Date(row[3]); // End Time = D
      tstop.setDate(date.getDate());
      tstop.setMonth(date.getMonth());
      tstop.setYear(date.getYear());
  
      Logger.log(tstart);
      var newEvent = calendar.createEvent(title, tstart, tstop);
      row[EVENT_ID] = newEvent.getId();
    }
}
