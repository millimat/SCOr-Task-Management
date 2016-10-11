// Scripts for the SCOr rehearsal scheduler.
// Matt Millican, Spring 2016.

/***************************************** GLOBALS ****************************************/
var warnings_silenced = false; // toggle when you need to work on Rows 2 and 3

var scorcal_name = "SCOr Internal";
var scor_email = "poos-comrades@lists.stanford.edu";
var scheduler_shortlink = "goo.gl/H3IELQ"

var epoch = new Date('Dec 30, 1899 00:00:00');
var ONE_DAY = 86400000; // 24h in ms
var ONE_MINUTE = 60000; // 1min in ms
var ONE_HOUR = ONE_MINUTE * 60; // 1hr in ms



/* (0-indexed) Indices of rows containing the listed information in an Object[][]
 * array created by fetching data from the top-left corner of the spreadsheet.
 * Since Javascript arrays are 0-indexed while Google sheets are 1-indexed,
 * each value is 1 less than its designated row number in the Google column.
 *
 * TO FUTURE WEBMASTER: THESE VALUES MUST BE UPDATED EVERY TIME THE SCHEDULER
 * UNDERGOES A SIGNIFICANT FORMATTING CHANGE. */
var inforows = {
EVENT_TYPE: 1,
DATE: 2,
START_TIME: 4,
END_TIME: 5,
LOCATION: 6,
};

/************************************** EDIT HANDLING ************************************/

/* Adds a menu item to propagate scheduler changes to the calendar. */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Calendar")
    .addItem("Propagate changes to calendar", "updateCalendar")
    .addToUi();
}

/* I wrote this function as to warn anyone editing the scheduler that they'll have to delete
 * obsolete events from the calendar, since there's no effective way for me to do that automatically.
 * (It would require either saving the contents of the entire calendar on startup and checking
 * them against the current version after edits -- which is extremely costly -- OR somehow getting
 * access to the pre-edit content of a cell within onEdit -- which is impossible.)
 * Unfortunately the warning is annoying as fuck, so I'm not sure whether to keep it.
 * -Matt 5/21/2016 */
function onEdit() {
    //  I tried to write these in so that warnings trigger at most once, but it seems like Google Scripts
    //  resets all variables every time the sheet is edited, so it doesn't work. -Matt
    //  if(onEdit.nameDateWarned == undefined) onEdit.nameDateWarned = false;
    //  Logger.log("User previously warned: " + onEdit.nameDateWarned);
    
    var currRow = SpreadsheetApp.getActiveRange().getRow() - 1; // 0-indexed row of edited cell
    
    
    if(currRow === inforows.DATE || currRow == inforows.EVENT_TYPE) { // date or name row was edited
        if(warnings_silenced) return;
        //  if(onEdit.nameDateWarned) return; // don't emit multiple warnings per session
        Logger.log("Warning user about editing name/date field...");
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert("Hold on!", "If you're changing the name/date of an " +
                                "existing event,\nyou'll have to delete the old " +
                                "event from the SCOr calendar\nto avoid duplicates.\n\n" +
                                "(To silence this warning, go to Tools > Script editor...\n "+
                                "and change warnings_silenced from false to true.)",
                                ui.ButtonSet.OK);
        
        //  onEdit.nameDateWarned = true;
    }
}


/*************************************** EVENT REPORTER **********************************/

/* Notify SCOr email list of today's events and enumerate
 * all absent or tardy members for the day. Runs once a day. */
function enumerateEvents() {
    var currsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // current event schedule
    var firstNameRow = getLastInfoRow(currsheet) + 2; // first (1-indexed) row of member names
    var lastRow = currsheet.getLastRow();
    var lastCol = currsheet.getLastColumn();
    var nameCol = currsheet.getSheetValues(firstNameRow, 1, lastRow, 1);
    var now = new Date();
    
    var eventString = ""; // Will contain title/time/location info for each event today.
    var absenteeString = ""; // Will contain a full list of absentees with excuses for each event today.
    
    Logger.log("Scanning scheduler for today's events...");
    for(var i = 2; i <= lastCol; i++) {
        var eventDate = currsheet.getRange(inforows.DATE + 1, i).getValue();
        var eventTitle = currsheet.getRange(inforows.EVENT_TYPE + 1, i).getValue();
        var eventLocation = currsheet.getRange(inforows.LOCATION + 1, i).getValue();
        var eventStart = currsheet.getRange(inforows.START_TIME + 1, i).getDisplayValue();
        var eventEnd = currsheet.getRange(inforows.END_TIME + 1, i).getDisplayValue();
        var absenteeInfo = currsheet.getSheetValues(firstNameRow, i, lastRow, 1); // absence excuses
        
        // Skip over previous dates or garbage dates.
        if(eventDate < dayStart(now) || isNaN(eventDate.getTime())) continue;
        // Stop processing when the listed date is in the future. No need to report.
        if(eventDate > now) break;
        
        eventString += "<b>- " + eventTitle + ":</b> " + eventLocation + ", " + eventStart + " - " + eventEnd + "<br/>";
        absenteeString += "<b>--- " + eventTitle + " ---</b><br/><br/>"; // delimiter between lists of absentee info
        absenteeString += enumerateAbsentees(nameCol, absenteeInfo);
    }
    
    // notify SCOr if anything is happening today
    if(eventString != "") {
        Logger.log("Notifying SCOr of today's events...");
        sendReport(eventString, absenteeString, now);
        Logger.log("Notification sent.");
    } else {
        Logger.log("No events found for today. No report sent.");
    }
}

/* Adds the names and excuses for all of today's absentees to the designated
 * absentee message. */
function enumerateAbsentees(names, excuses) {
    message = "";
    
    for(var i = 0; i < names.length; i++) {
        var absentee = names[i][0];
        var excuse = excuses[i][0];
        
        if(excuse != "") { // person will be absent/tardy
            message += absentee + " - \"" + excuse + "\"<br/>";
        }
    }
    
    return message + "<br/><br/>";
}

/* Sends an email to the SCOr main list notifying them of today's events and all
 * absentees for each event. */
function sendReport(eventString, absenteeString, now) {
    var recipient = scor_email;
    var subject = "SCOr Event Report for " + now.toDateString();
    var msgHtml = "Good morning, SCOr!<br/><br/>" + "According to the event scheduler (" + scheduler_shortlink +"), "
    + "there are one or more events happening today. Here is the full list:<br/><br/>"
    + eventString + "<br/>Please consult the scheduler or the SCOr calendar for more detailed "
    + "information.<br/><br/>" + "Below is a list of members who expect to be tardy or absent for today. "
    + "If you expect to miss a mandatory event today and your name is not listed below, "
    + "please notify the SCOr email list immediately to explain your situation.<br/><br/>"
    + "See you soon!<br/><br/><br/>" + absenteeString;
    
    var msgPlain = htmlToPlain(msgHtml);
    
    MailApp.sendEmail(recipient, subject, msgPlain, { htmlBody: msgHtml });
}


/*************************************** UPDATING CALENDAR **********************************/


/* Update all upcoming events on the SCOr calendar to reflect
 * what's currently on the rehearsal scheduler. This should run once an hour.
 * NOTE: If you change the date of any events on the calendar, they must
 * be deleted manually from the calendar, since this function will
 * will create duplicates. */
function updateCalendar() {
    var currsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // current event schedule
    var scorcal = CalendarApp.getCalendarsByName(scorcal_name)[0]; // The SCOr calendar
    var lastInfoRow = getLastInfoRow(currsheet);
    var lastCol = currsheet.getLastColumn();
    var titleCol = currsheet.getSheetValues(1, 1, lastInfoRow, 1);
    
    var now = new Date();
    var prevdate = new Date(0);
    var prevdate_events;
    
    /* Iterate over all columns in spreadsheet, updating their event information. */
    for(var i = 2; i <= lastCol; i++) {
        // Get list of events in SCOr calendar for this day (keep same list if day is same as prev col)
        var einfo = currsheet.getSheetValues(1, i, lastInfoRow, 1);
        var eventDate = einfo[inforows.DATE][0];
        
        // Don't edit calendar for old events or dateless events.
        if(eventDate < dayStart(now) || isNaN(eventDate.getTime())) continue;
        
        // Fetch a new list of active events on this day if the current day differs from the last one
        var thisdate_events = prevdate_events;
        if(eventDate !== prevdate) thisdate_events = scorcal.getEventsForDay(eventDate);
        
        updateEvent(scorcal, titleCol, einfo, thisdate_events, eventDate.getTime());
        
        prevdate = eventDate;
        prevdate_events = thisdate_events;
    }
    
}

/* Searches the list of SCOr events happening on the date corresponding to the current column
 * of the rehearsal scheduler. If the name of the scheduler event matches the name of an
 * event on the calendar, the calendar event has its time, location, and description updated
 * to reflect the information currently on the scheduler. If the scheduler event does not match
 * any current calendar event, a new event is added to the calendar to reflect the scheduler event.
 */
function updateEvent(calendar, titleCol, event_info, thisdate_events, eventDate_ms) {
    var eventName = event_info[inforows.EVENT_TYPE][0];
    var eventStart = new Date(0); // Initialize to NaN. If no time found, event will be all-day
    var eventEnd = new Date(0);
    
    if(event_info[inforows.START_TIME][0] != "") { // The rehearsal scheduler has time/date info
        var startTime_ms = event_info[inforows.START_TIME][0] - epoch; // time past midnight in seconds
        var endTime_ms = event_info[inforows.END_TIME][0] - epoch;
        eventStart = new Date(eventDate_ms + startTime_ms);
        eventEnd = new Date(eventDate_ms + endTime_ms);
    }
    
    var matchFound = false;
    for(var j = 0; j < thisdate_events.length; j++) { // update event info to reflect scheduler
        var calevent = thisdate_events[j];
        
        if(calevent.getTitle() == eventName) { // Scheduler event matches this calendar event
            matchFound = true;
            updateTime(calevent, eventStart, eventEnd, eventDate_ms);
            setRehearsalInfo(calevent, event_info, titleCol);
            break;
        }
    }
    
    if(!matchFound) { // create new calendar event
        var calevent = calendar.createEvent(eventName, epoch, epoch); // initialize event with garbage times b4 finalizing
        updateTime(calevent, eventStart, eventEnd, eventDate_ms);
        setRehearsalInfo(calevent, event_info, titleCol);
    }
}


/* Updates the time of a calendar event based on the time information found in
 * the scheduler. If the scheduler provided no time information for the event,
 * it is set as an all-day event on the date provided in the rehearsal scheduler.
 * If time information was provided for the event, sets the event time using that
 * information. */
function updateTime(calevent, eventStart, eventEnd, eventDate_ms) {
    if(eventStart.getTime() === 0 || eventEnd.getTime() === 0) {
        // No start time for event or no end time for event. Set as all-day event.
        calevent.setAllDayDate(new Date(eventDate_ms));
    } else {
        // Start time and end time provided. Set on calendar.
        calevent.setTime(eventStart, eventEnd);
    }
}

/* Sets the location of the event from scheduler info. Then sets the event description
 * by iterating over all rows of the scheduler that represent rehearsal segments (run-throughs
 * and slots for pieces) and piecing that segment info into a string. */
function setRehearsalInfo(calevent, event_info, titleCol) {
    var eventLocation = event_info[inforows.LOCATION][0];
    calevent.setLocation(eventLocation);
    
    // Iterate over rows containing intra-rehearsal scheduling to add to event desc.
    var description = "";
    for(var i = inforows.LOCATION + 1; i < titleCol.length; i++) {
        var portion = titleCol[i];
        var timeslot = event_info[i];
        if(timeslot != "") {
            description += portion + ": " + timeslot + "\n";
        }
    }
    
    calevent.setDescription(description);
    
}


/**************************** HELPER FUNCTIONS *******************************/

/* Get the (Google 1-indexed) number of the final row in the spreadsheet
 * that contains info usable for calendar data. Assumes the row after
 * event data is blank. */
function getLastInfoRow(currsheet) {
    var rownames = currsheet.getSheetValues(1, 1, -1, 1) // All data in first column
    //  Logger.log(rownames);
    
    var index = 0;
    while(index < rownames.length) {
        if(rownames[index] == "") { // empty value, end of info
            break; // we're 1 past the 0-indexed row of the last data row, so this is the right 1-indexed value
        }
        index++;
    }
    
    Logger.log("Last row with calendar info (1-indexed): " + index);
    return index;
}

/* Return the time in milliseconds of the midnight that began the input date.
 * e.g. inputting any date corresponding to July 4th, 2016 will return the time in
 * milliseconds of Jul 4, 2016 00:00:00 in the time zone where the date object
 * originated.  */
function dayStart(date) {
    var offsetms = date.getTimezoneOffset() * ONE_MINUTE;
    var time = date.getTime();
    var timePastMidnight = (time - offsetms) % ONE_DAY;
    
    return time - timePastMidnight;
}

/* Convert text with HTML tags into a "regular" version with no tags and
 * "\n" instead of "<br/>. Found at:
 * http://stackoverflow.com/questions/9442375/how-to-bold-specific-text-using-google-apps-script */
function htmlToPlain(string) {
    // clear html tags and convert br to new lines for plain mail
    return string.replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, ""); 
}
