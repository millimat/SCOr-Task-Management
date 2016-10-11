// Scripts for the HAL Tasklist for SCOr.
// Tasklist formatting and sorting by Allison Yuen, 2015.
// Notification system by Matt Millican, 2016.

/**********************************  GLOBALS  *************************************/
//HAL columns
var tasknameCol = 1;
var urgencyCol = 3;
var assigneeCol = 4;
var emailCol = 5;
var completeDateCol = 6;
var statusCol = 7;
var notifyCol_week = 8;
var notifyCol_day = 9;
var notifyCol_overdue = 10;

// HAL rows
var firstDataRow = 2;

// Other
var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1i_FRY8jtwKhA7-g01jygOyjuX_-CHqtk3w4UBMGv9oQ/";
var shortUrl = "https://goo.gl/mUZSdU";
var webmasterEmail = "millimat@stanford.edu";

/*******************************    MAIN BODY    **********************************/

/* Called when the sheet is opened. Adds a menu item to sort the
 * tasklist. */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Sort")
    .addItem("Sort HAL", "sortTasklist")
    .addToUi();
}

/* Called when an edit is made to the spreadsheet to sort the
 * tasklist if a column that significantly affects sorting
 * is edited, to fill in an assignee's email if their name is added,
 * or to notify an assignee of their task if a date is added to that task. */
function onEdit() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet()
    var editedColumn = sheet.getActiveRange().getColumnIndex() - 1; // subtract 1 because Google is stupid and 1-indexes
    var editedRow = sheet.getActiveRange().getRowIndex() - 1;
    
    if(sheet.getName() === "HAL" && editedRow > 0) { // Most onEdit stuff only pertains to the HAL sheet
        Logger.log("Editing HAL")
        
        if (editedColumn == completeDateCol || editedColumn == statusCol
            || editedColumn == urgencyCol ) {
            Logger.log("Sorting tasklist...")
            sortTasklist(sheet);
        } else if(editedColumn == assigneeCol) { // Fill in assignee email from name
            Logger.log("Editing assignee email col...");
            fillEmail(sheet, editedRow, editedColumn);
        }
        
        if(editedColumn == completeDateCol) { // Set notification flags for time-sensitive event
            Logger.log("Task assigned new date. Priming notifications...");
            setNotifyFlags(sheet, editedRow);
        }
    }
}

/*************************  TASKLIST SORTING  *************************/

/* Called if the edited column might affect the sorted
 * status of the tasklist in a significant way.
 */
function sortTasklist(sheet) {
    sheet = sheet || SpreadsheetApp.openByUrl(spreadsheetUrl).getSheets()[0];
    var range = sheet.getDataRange();
    var data = range.getValues();
    var headers = data.splice(0, 1)[0]; // remove and retrieve headers
    data.sort(compareTasks);
    data.splice(0, 0, headers); // put headers back in
    range.setValues(data);
}

/* Takes two arrays representing rows of data
 * from the HAL tasklist and compares. Returns
 * 1 if a should be lower down on the tasklist
 * than b, -1 if a should be higher.
 */
function compareTasks(a, b) {
    var compare = 0;
    
    // compare status (complete or incomplete)
    if (a[statusCol] == "" && b[statusCol] != "") {
        compare = -1;
    } else if (a[statusCol] != "" && b[statusCol] == "") {
        compare = 1;
    } else if (a[statusCol] != "" && b[statusCol] != "") {
        return 0; // if both are complete, then we don't care
    }
    
    if (compare != 0) return compare;
    
    // compare urgency (high/medium/low)
    compare = compareUrgency(a[urgencyCol], b[urgencyCol]);
    if (compare != 0) return compare;
    
    // compare assignee (alphabetical)
    if (a[assigneeCol] > b[assigneeCol]) {
        compare = 1;
    } else if (a[assigneeCol] < b[assigneeCol]) {
        compare = -1;
    }
    
    return compare;
}


/* Takes two strings representing urgency values
 * ("high", "medium", "low", or "") and compares.
 */
function compareUrgency(a, b) {
    a = a.toLowerCase().trim();
    b = b.toLowerCase().trim();
    
    if (a == "high") {
        if (b == "high") return 0;
        else return -1;
    } else if (a == "medium") {
        if (b == "high") return 1;
        else if (b == "medium") return 0;
        else return -1;
    } else if (a == "low") {
        if (b == "") return -1;
        else if (b == "low") return 0;
        else return 1;
    } else {
        if (b == "") return 0;
        else return 1;
    }
}

/*************************  ASSIGNEE EMAIL HANDLING  *************************/

/* When a task is assigned to someone, fill in their email in the adjacent column */
function fillEmail(sheet, editedRow, editedColumn) {
    var assignees = assigneeMap();
    var editedCell = sheet.getRange(editedRow + 1, editedColumn + 1);
    var assigneeName = toProperCase(editedCell.getValue());
    var assigneeEmail = assignees[assigneeName];
    
    if(assigneeEmail == undefined) { // Username not in the map of assignees
        assigneeEmail = "Email not found";
    }
    
    var target = sheet.getRange(editedRow+1, editedColumn+2);
    target.setValue(assigneeEmail);
}


/* Get map with list of approved HAL editors/assignees. */
function assigneeMap() {
    var result = {};
    
    var editorListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Editor Info");
    var firstEditorRow = 2;
    var firstEditorCol = 1;
    
    // Get the data range corresponding to the list of active editors.
    var editorData = editorListSheet.getSheetValues(firstEditorRow,firstEditorCol,-1,-1);
    
    // Map assignee names to Stanford emails.
    for(var i = 0; i < editorData.length; i++) {
        var fullname = editorData[i][0];
        var singlenames = fullname.split(" ");
        var email = editorData[i][1];
        Logger.log(fullname);
        Logger.log(singlenames);
        Logger.log(email);
        
        // Map the editor's full name and single names to their email address
        result[editorData[i][0]] = email;
        for(var j = 0; j < singlenames.length; j++) {
            result[singlenames[j]] = email;
        }
    }
    
    Logger.log("Assignee map processed")
    return result;
}

/*************************  NOTIFICATION HANDLING  *************************/

/* Notification-handling enums */

// Notification threshold times in milliseconds. (Date objects are times in ms.)
var notifyTimes = {
WEEK: 7 * 86400000,
DAY: 86400000,
OVERDUE: -1.5 * 86400000 // 12 hrs after the due date is totally over
}

// The status of a task as displayed in a subject line
var subjectStatusText = {
WEEK: " Task due in one week for ",
DAY: " Task due in 24 hours for ",
OVERDUE: " Task past due for ",
}

// The status of a task as displayed in the main body of an email
var notifyStatusText = {
WEEK: " is due in one week, on ",
DAY: " is due in 24 hours, on ",
OVERDUE: " is overdue. It was due to be completed by ",
}

/* Tentative notification system:
 * -Always notify assignee if a date is assigned to their task.
 * -Always notify assignee of task one week in advance.
 * -Notify assignee of task 24 hours in advance if priority >= Medium.
 * -Notify assignee of overdue task 12 hours past due date if priority >= High.
 * -If priority is "Low" or unrecognized, priority is treated as Low.
 * -Only set notification flags if the due date, assignee, and taskname have valid format.
 * 3.28.16 - Matt */
function setNotifyFlags(sheet, editedRow) {
    var dateCell = sheet.getRange(editedRow + 1, completeDateCol + 1);
    var dueDate = new Date(dateCell.getValue());
    var currDate = new Date();
    
    var oneWeekFlag = "N/A";
    var oneDayFlag = "N/A";
    var overdueFlag = "N/A";
    
    var assigneeInfo = sheet.getSheetValues(editedRow + 1, assigneeCol + 1, 1, 2) // assignee name and email
    var assigneeName = assigneeInfo[0][0];
    var assigneeEmail = assigneeInfo[0][1];
    var taskname = sheet.getRange(editedRow + 1, tasknameCol + 1).getValue();
    
    /* Only set notify tags and notify assignee if the due date, assignee name/email, and task name valid */
    if(!isNaN(dueDate.getTime()) && assigneeName !== ""
       && validateEmail(assigneeEmail) && taskname !== "") {
        var priorityCell = sheet.getRange(editedRow + 1, urgencyCol + 1);
        var priority = toProperCase(priorityCell.getValue());
        var timeTillDue = dueDate - currDate;
        
        oneWeekFlag = flagIfDateUpcoming(timeTillDue, notifyTimes.WEEK);
        if(priority === "High" || priority == "Medium") oneDayFlag = flagIfDateUpcoming(timeTillDue, notifyTimes.DAY);
        if(priority === "High") overdueFlag = flagIfDateUpcoming(timeTillDue, notifyTimes.OVERDUE);
        
        if(assigneeName !== "All") { // Do not notify all of SCOr of assignment of universal task
            initializationNotify(currDate, dueDate, assigneeName, assigneeEmail, taskname);
        }
    }
    
    // Set notification flags
    sheet.getRange(editedRow + 1, notifyCol_week + 1).setValue(oneWeekFlag);
    sheet.getRange(editedRow + 1, notifyCol_day + 1).setValue(oneDayFlag);
    sheet.getRange(editedRow + 1, notifyCol_overdue + 1).setValue(overdueFlag);
    
}

/* Notify the assignee of a task that the date of their task has been added/updated. */
function initializationNotify(currDate, dueDate, assigneeName, assigneeEmail, taskname) {
    // Only notify if task is upcoming
    if(currDate > dueDate) return;
    
    Logger.log("Notifying user " + assigneeName + " at email " + assigneeEmail + " of their assignment to new task...");
    
    var subject = "[hal-notification-system] " + getTimestamp(currDate) + " Task assignment for " + assigneeName;
    var msgHtml = "Hi, " + assigneeName + ",<br/><br/>"
    + "This is a notification from the HAL tasklist to inform you that you have been assigned "
    + "to the following task (or its due date has been updated):<br/><br/><i>"
    + taskname + "</i><br/><br/>This task is currently due on " + dueDate.toDateString()
    + ". Check it out on the HAL spreadsheet here:<br/><br/>" + shortUrl + "<br/><br/>"
    + "If you think you have received this email in error, please forward its contents to "
    + "the SCOr webmaster at the following address so they can fix the issue:<br/><br/>"
    + webmasterEmail + "<br/><br/>Have a great day!";
    
    var msgPlain = htmlToPlain(msgHtml);
    
    MailApp.sendEmail(assigneeEmail, subject, msgPlain, {htmlBody: msgHtml})
}


/* Emit emails for all pending status notifications on HAL. */
function processNotifications() {
    /* Accumulate all data that could be relevant to noticy the assignee of an unfinished task. */
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var currDate = new Date();
    var lastIncRow = firstCompleteRow(sheet) - 1; // 1-indexed row of last incomplete task
    var nIncRows = lastIncRow - firstDataRow + 1;
    
    var notifyData = sheet.getSheetValues(firstDataRow, notifyCol_week + 1, nIncRows, 3);
    var completionDates = sheet.getSheetValues(firstDataRow, completeDateCol + 1, nIncRows, 1);
    var assigneeData = sheet.getSheetValues(firstDataRow, assigneeCol + 1, nIncRows, 2);
    var taskData = sheet.getSheetValues(firstDataRow, tasknameCol + 1, nIncRows, 1);
    
    for(var i = 0; i < nIncRows; i++) {
        /* Collect data for the individual task. */
        var notified_week = notifyData[i][0];
        var notified_day = notifyData[i][1];
        var notified_overdue = notifyData[i][2];
        
        var dueDate = new Date(completionDates[i]);
        var assigneeName = assigneeData[i][0];
        var assigneeEmail = assigneeData[i][1];
        var taskName = taskData[i][0];
        
        // For debugging: output task data to log
        // Logger.log("Task %s: taskName = '%s', assigneeName = '%s', assigneeEmail = '%s', "
        //           + "dueDate = %s", i+1, taskName, assigneeName, assigneeEmail, dueDate.getTime());
        
        /* If task, due date, and assignee info are all valid, check if the spreadsheet has pending notification */
        if(!isNaN(dueDate.getTime()) && assigneeName !== ""
           && validateEmail(assigneeEmail) && taskName !== "") {
            var flaggingCell = getFlaggingCell(sheet, notified_week, notified_day, notified_overdue, i+2);
            
            /* If notification pending, check its time to see if it's time to notify */
            if(flaggingCell !== undefined) {
                var threshold_and_status = new Array(4); // Timestamp, threshold, subject line, task status
                fillThresholdStatus(flaggingCell, currDate, threshold_and_status);
                var timeTillDue = dueDate - currDate;
                
                /* If it is time to notify, do so. Email poos-comrades for 'All' or individual user for a person */
                if(timeTillDue <= threshold_and_status[1]) {
                    Logger.log("All checks passed for notification. Preparing to notify...")
                    switch(assigneeName) {
                        case "All":
                            notifySCOr(assigneeEmail, taskName, dueDate, threshold_and_status);
                            break;
                        default:
                            notifySingleAssignee(assigneeName, assigneeEmail, taskName, dueDate, threshold_and_status);
                            break;
                    }
                    // The cell has emitted its notification, so mark it as complete
                    Logger.log("Setting value of notification-emitting cell " + flaggingCell.getA1Notation() + " to 'Y'");
                    flaggingCell.setValue("Y");
                }
            }
        }
    }
}

/* Use the status of a cell flagging a notification to fill in all the information
 * necessary to send out a templatized notification email:
 * [0]: timestamp of email;
 * [1] the time threshold for sending the email;
 * [2]: the status of the task as it appears in an email subject line;
 * [3]: the status of the task as it appears in the main body of an email. */
function fillThresholdStatus(flaggingCell, currDate, arr) {
    arr[0] = getTimestamp(currDate);
    
    var col = flaggingCell.getColumn() - 1;
    if(col === notifyCol_week) {
        arr[1] = notifyTimes.WEEK;
        arr[2] = subjectStatusText.WEEK;
        arr[3] = notifyStatusText.WEEK;
    } else if(col === notifyCol_day) {
        arr[1] = notifyTimes.DAY;
        arr[2] = subjectStatusText.DAY;
        arr[3] = notifyStatusText.DAY;
    } else if(col === notifyCol_overdue) {
        arr[1] = notifyTimes.OVERDUE;
        arr[2] = subjectStatusText.OVERDUE;
        arr[3] = notifyStatusText.OVERDUE;
    }
}

/* Fetch the cell whose flag will be used to send a notification,
 * if any. The lowest-urgency notification type with a pending flag
 * is selected as the notification trigger: week before day before overdue. */
function getFlaggingCell(sheet, weekflag, dayflag, overflag, row) {
    result = undefined;
    
    if(weekflag === "N") {
        result = sheet.getRange(row, notifyCol_week + 1);
    } else if(dayflag === "N") {
        result = sheet.getRange(row, notifyCol_day + 1);
    } else if(overflag === "N") {
        result = sheet.getRange(row, notifyCol_overdue + 1);
    }
    
    return result;
}

/* Notify all of SCOr of the status of a task. */
function notifySCOr(email, task, dueDate, statusArr) {
    var subject = "[hal-notification-system] " + statusArr[0] + statusArr[2] + "SCOr";
    
    var body = "Hi, SCOr,\n\n"
    + "This is a notification from the HAL tasklist that the task "
    + "'" + task + "'" + statusArr[3] + dueDate.toDateString() + ".\n\nIf this task is already complete, "
    + "please visit the HAL spreadsheet at the following address to mark it as complete:\n\n" + shortUrl;
    
    MailApp.sendEmail(email, subject, body);
    Logger.log("Notification complete.");
}

/* Notify a specific user of the status of a task. */
function notifySingleAssignee(name, email, task, dueDate, statusArr) {
    var subject = "[hal-notification-system] " + statusArr[0] + statusArr[2] + name;
    
    var body = "Hi, " + name + ",\n\n"
    + "This is a notification from the HAL tasklist to remind you that your assigned task "
    + "'" + task + "'" + statusArr[3] + dueDate.toDateString() + ".\n\n"
    + "If you have already completed this task, please visit the HAL spreadsheet at the following "
    + "address to mark it as complete:\n\n" + shortUrl + "\n\n"
    + "If you think you have received this email in error, please forward its contents to "
    + "the SCOr webmaster at the following address so he or she can fix the issue:\n\n"
    + webmasterEmail + "\n\nHave a great day!";
    
    MailApp.sendEmail(email, subject, body);
    Logger.log("Notification complete.");
}

/*************************  UTILITY FUNCTIONS  *************************/

/* Convert a name to proper capitalization. */
function toProperCase(s) {
    return s.toLowerCase().replace(/^(.)|\s(.)/g,
                                   function($1) { return $1.toUpperCase(); });
}

/* Using regular expressions, return a boolean reflecting whether
 * the input email address is valid. */
function validateEmail(email) {
    // this did not come from my brain, it came from the interwebs
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    var result = re.test(email);
    
    // For debugging: Log decision of regex on email
    // Logger.log("Email '" + email + "' is valid? " + result);
    return result;
}

/* Military time of current date, without time zone info. */
function getTimestamp(currDate) {
    return "(" + currDate.toTimeString().split(" ")[0] + ")";
}

/* The 1-indexed row number of the first complete task in the spreadsheet. */
function firstCompleteRow(sheet) {
    var i = firstDataRow;
    var last = sheet.getLastRow();
    
    for(; i <= last; i++) {
        var cell = sheet.getRange(i, statusCol + 1);
        if(cell.getValue() !== "") break;
    }
    
    Logger.log("Row of first complete task: " + i);
    return i;
}

/* Given a valid duedate which much be used to prime a HAL notification,
 * ONLY flag as "N" (notification pending and not sent) if the time for this
 * notification has not already passed. This occurs if the duedate's time
 * minus the current date's time is above the threshold for the relevant notification. 
 * If the threshold time has already passed, mark the notification as "Y" (already sent)
 * so that HAL does not emit multiple unnecessary notifications. */
function flagIfDateUpcoming(timeTillDue, threshold) {
    return (timeTillDue > threshold ? "N" : "Y");  
}

/* Convert text with HTML tags into a "regular" version with no tags and
 * "\n" instead of "<br/>. Found at:
 * http://stackoverflow.com/questions/9442375/how-to-bold-specific-text-using-google-apps-script */
function htmlToPlain(string) {
    // clear html tags and convert br to new lines for plain mail
    return string.replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, ""); 
}
