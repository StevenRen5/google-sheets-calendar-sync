/*
  Google Sheets ↔ Calendar Sync Script
  ------------------------------------
  Adds a "Calendar Sync" menu to your Google Sheet Job Tracker to create, update, and delete events
  in your default Google Calendar based on spreadsheet data.
  
  Author: Steven Ren
  License: MIT
*/



/*
Creates a custom menu in spreadsheet to run the sync script
onOpen() is a built-in function that runs when the associated Google Sheet opens
*/
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Calendar Sync') // creates a new top-level menu named 'Calendar Sync'
  .addItem('Sync Job Applications', 'createCalEvent') // adds a menu item that runs your createCalEvent() function
  .addToUi() // displays menu in the spreadsheet
}

/*
Creates calendar events
*/
function createCalEvent() {
  const cal = CalendarApp.getDefaultCalendar();

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getActiveSheet();

  let data = sheet.getDataRange().getValues(); 
  // .getDataRange() selects all cells that contain data
  // .getValues() retrieves the values from all cells
  // schedule.splice();
  // printing schedule will return a list, with each element being a list of cell values from each row
  data.splice(0,2); // removes the first row, which is the header

  // Column Indices (CAN CHANGE)
  const title_col = 0;
  const start_date_col = 3;
  const end_date_col = 4;
  const event_id_col = 12;

  // aesthetics for event (CAN CHANGE)
  const eventColor = CalendarApp.EventColor.GRAY;


  data.forEach((entry, index) => {
    // entry is the row, denoted as a list
    /* for rowInSheet:
      sheet uses 1-based indexing (row 1,2,3..),
      but array data is 0-based and is missing the header (row 1), 
      so rowInSheet = index + 2
    */
    const rowInSheet = index + 3
    const eventId = entry[event_id_col];
    const title = entry[title_col];
    const startTime = entry[start_date_col];
    const endTime = entry[end_date_col];

    /*
    EVENT DELETION
    Condition: Event ID exists, but Title and Dates are missing (user cleared the row). We will clear the row for that event ID.
    */
    if (eventId && (!title && !startTime && !endTime)) {
      try {
        const event = cal.getEventById(eventId);
        event.deleteEvent(); // deletes the event in calendar

        /* NOTE:
        rowInsheet is the current row we're on
        1 is the starting column (Column A)
        1 is number of rows (just one row)
        event_id_col is the number of columns to clear
        */
        sheet.getRange(rowInSheet, 1, 1, event_id_col + 1).clearContent();

        return; // stops processing this row
      } catch(e) {
        // happens when event is deleted from calendar or altered ID
        console.log(`ERROR: Could not find or delete event with ID ${eventId} at row ${rowInSheet}.`);
        return;
      }
    }

    // Condition: Event ID exists, and will attempt to update the calendar event.
    // If event doesn't exist in calendar, we will clear the row.
    if (eventId) {
      try {
        let event = cal.getEventById(eventId);
        event.setTitle(title);
        event.setTime(startTime, endTime);
        return;
      } catch(e) {
        console.log(`ERROR: Could not get event from Calendar with ID ${eventId} at row ${rowInSheet}`);

        sheet.getRange(rowInSheet, 1, 1, event_id_col + 1).clearContent();

        return;
      }
    }
    else { 
      /* 
      IF EVENT DOESN'T EXIST, CREATE EVENT
      */
      if (title && startTime && endTime) {
        try {
          const newEvent = cal.createEvent(title, startTime, endTime);
          const newEventId = newEvent.getId();
          newEvent.setColor(eventColor);

          // write ID back to sheet
          // .getRange(row, column) is used to work with a single cell
          // Note: event_id_col + 1 is b/c sheet uses one-based index instead of zero-based
          sheet.getRange(rowInSheet, event_id_col + 1).setValue(newEventId);
          return; 
        } catch(e) {
          console.log(`ERROR: Could not create event for row ${rowInSheet}. Check date and time format.`);
          return; 
        }
      }
    }
  });
}
