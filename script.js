/**
 * SECTION 1: AUTOMATION LOGIC (Runs on Minute Trigger)
 */
function processTaskEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 

  const range = sheet.getRange(1, 1, lastRow, 9);
  const data = range.getValues();
  const now = new Date();
  const myEmail = Session.getActiveUser().getEmail();

  for (let i = 1; i < data.length; i++) {
    let rowNum = i + 1;
    const title     = data[i][1]; 
    const content   = data[i][2]; 
    const category  = data[i][3]; 
    const rawDate   = data[i][4]; 
    const rawTime   = data[i][5]; 
    const repeat    = data[i][6]; 
    const frequency = data[i][7]; 
    const status    = data[i][8] ? data[i][8].toString().trim() : ""; 

    if (status.toLowerCase() === "completed" || !title) continue;

    let dDate = new Date(rawDate);
    let dTime = new Date(rawTime);
    if (isNaN(dDate.getTime()) || isNaN(dTime.getTime())) continue;

    let scheduledTime = new Date(dDate.getFullYear(), dDate.getMonth(), dDate.getDate(), 
                                 dTime.getHours(), dTime.getMinutes(), 0);

    if (scheduledTime <= now) {
      try {
        MailApp.sendEmail(myEmail, `Task: ${title}`, `Category: ${category}\n\n${content}`);
        sheet.getRange(rowNum, 10).setValue("Sent: " + Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM-dd HH:mm"));

        if (repeat === true || String(repeat).toUpperCase() === "TRUE") {
          let nextDue = new Date(scheduledTime);
          let freq = String(frequency).toLowerCase().trim();
          if (freq === "everyday") nextDue.setDate(nextDue.getDate() + 1);
          else if (freq === "everyweek") nextDue.setDate(nextDue.getDate() + 7);
          else if (freq === "everymonth") nextDue.setMonth(nextDue.getMonth() + 1);
          
          sheet.getRange(rowNum, 5).setValue(nextDue); 
          sheet.getRange(rowNum, 9).setValue("Active");
        } else {
          sheet.getRange(rowNum, 9).setValue("Completed"); 
        }
      } catch (e) {
        console.log(`Row ${rowNum} Error: ${e.message}`);
      }
    }
  }
}

/**
 * SECTION 2: WEB APP INTERFACE API
 */
function doGet() {
  return HtmlService.createTemplateFromFile('indeks')
    .evaluate()
    .setTitle('Task Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getActiveTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let tasks = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][8] ? data[i][8].toString().trim() : "";
    if (status.toLowerCase() !== "completed") {
      tasks.push({
        row: i + 1,
        title: data[i][1],
        content: data[i][2],
        category: data[i][3],
        dueDate: data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][4],
        dueTime: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "HH:mm") : data[i][5],
        frequency: data[i][7]
      });
    }
  }
  // PROTOCOL: Sort by recently added (newest row first)
  return tasks.reverse(); 
}

function saveTask(taskData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nextRow = sheet.getLastRow() + 1;
  const repeatBool = (taskData.frequency !== "none");
  
  sheet.getRange(nextRow, 2, 1, 8).setValues([[
    taskData.title, taskData.content, taskData.category, 
    taskData.dueDate, taskData.dueTime, repeatBool, 
    taskData.frequency, "Active"
  ]]);
  return "Success";
}

function deleteTask(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, 9).setValue("Completed"); 
  return "Deleted";
}
