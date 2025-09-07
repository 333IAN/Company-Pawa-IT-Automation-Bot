// ==================== CONFIGURATION ====================
const SPREADSHEET_ID = '1Kye4jZqCR_dGtSCDet-yT7Af6AbbLboURyXrbnZMW40';
const SUPERVISOR_EMAIL = /* HR/SUPERVISOR EMAIL*/;
const WEB_APP_URL = 'https://script.google.com/a/macros/pawait.co.ke/s/AKfycbzrhB1Q9aerdMjnhGGiEWvAng55yg9o1EumsMlA6Hi-Ml9NEVvDPq90jls7Ll-mkTBS/exec';

// ==================== ACCESS CONTROL ====================
// Define who can access the Lunch Bot menu
const AUTHORIZED_USERS = [
  /*Emails of those who can access/run the bot*/
];

// ==================== MAIN FUNCTIONS ====================

/**
 * Creates custom menu when spreadsheet opens(it's icon on spreadsheet)
 */
function onOpen() {
  const userEmail = Session.getActiveUser().getEmail();
  
  // Check if user is authorized to see the Lunch Bot menu
  if (!AUTHORIZED_USERS.includes(userEmail)) {
    console.log(`Access denied for ${userEmail} - not in authorized users list`);
    return; // Exit early - no menu will be created
  }
  
  // User is authorized - create the menu
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍱 Pawa IT Lunch Reminder')
    .addItem('Send Email Reminders', 'sendEmailReminders')
    .addItem('View Pending Responses', 'showPendingDialog')
    .addSeparator()
    .addItem('Setup Daily Triggers', 'setupDailyTriggers')
    .addItem('View Current Triggers', 'viewCurrentTriggers')
    .addItem('Delete All Triggers', 'deleteAllTriggers')
    .addSeparator()
    .addItem('Supervisor Report', 'sendSupervisorReport')
    .addToUi();
    
  console.log(`Lunch Bot menu created for authorized user: ${userEmail}`);
}

/**
 * Get today's sheet name automatically
 */
function getTodaysSheetName() {
  const today = new Date();
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri',];
  const dayName = days[today.getDay()];
  const dayNum = today.getDate();
  
  let suffix = 'th';
  if (dayNum % 10 === 1 && dayNum !== 11) suffix = 'st';
  else if (dayNum % 10 === 2 && dayNum !== 12) suffix = 'nd';
  else if (dayNum % 10 === 3 && dayNum !== 13) suffix = 'rd';
  
  return `${dayName} ${dayNum}${suffix}`;
}

function today(){
  todaysheet=getTodaysSheetName();
  console.log('Today sheet name is:',todaysheet)
}

/**
 * Get employees who haven't responded (empty lunch choice)
 */
function getEmployeesNeedingReminders() {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log(`Sheet "${sheetName}" not found - creating new sheet`);
    return createNewDaySheet(sheetName);
  }
  
  const data = sheet.getDataRange().getValues();
  const employees = [];
  
  //Column indices based on spreadsheet structure
  const nameCol = 0;    // Column A: EMPLOYEE NAME
  const emailCol = 1;   // Column B: EMAIL ADDRESS
  const lunchCol = 2;   // Column C: TAKING LUNCH Y/N
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    const email = data[i][emailCol];
    const lunchChoice = data[i][lunchCol];
    
    if (name && email && (!lunchChoice || lunchChoice === '')) {
      employees.push({ name, email, row: i + 1 });'jah'xxx
    }
  }
  
  return employees;
}

/**
 * Create a new day sheet with template structure
 */
function createNewDaySheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const newSheet = ss.insertSheet(sheetName);
  
  const sheets = ss.getSheets();
  let templateSheet = null;
  
  for (let i = sheets.length - 1; i >= 0; i--) {
    if (sheets[i].getName() !== sheetName && isWeekdaySheet(sheets[i].getName())) {
      templateSheet = sheets[i];
      break;
    }
  }

  if (templateSheet) {
    templateSheet.getDataRange().copyTo(newSheet.getRange(1, 1));
    const lastRow = newSheet.getLastRow();
    if (lastRow > 1) {
      newSheet.getRange(2, 3, lastRow - 1, 1).clearContent();
    }
  } else {
    newSheet.getRange('A1:C1').setValues([['EMPLOYEE NAME', 'EMAIL ADDRESS', 'TAKING LUNCH Y/ ✔']]);
    newSheet.getRange('A1:C1').setFontWeight('bold');
    
    const allSheets = ss.getSheets();
    for (let sheet of allSheets) {
      if (sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] && data[i][1]) {
            newSheet.getRange(newSheet.getLastRow() + 1, 1, 1, 2).setValues([[data[i][0], data[i][1]]]);
          }
        }
        break;
      }
    }
  }
  
  console.log(`Created new sheet: ${sheetName}`);
  return [];
}

/**
 * Send email reminders to all employees who haven't responded
 */
function sendEmailReminders() {
  const employees = getEmployeesNeedingReminders();
  
  if (employees.length === 0) {
    SpreadsheetApp.getUi().alert('🎉 Everyone has already responded!');
    return;
  }
  
  employees.forEach(sendSingleEmailReminder);
  SpreadsheetApp.getUi().alert(`📧 Sent ${employees.length} email reminders`);
}

/**
 * Send individual email with clickable buttons
 */
function sendSingleEmailReminder(employee) {
  const { name, email } = employee;
  const encodedEmail = encodeURIComponent(email);
  
  const urls = {
    yes: `${WEB_APP_URL}?choice=Yes&email=${encodedEmail}`,
    no: `${WEB_APP_URL}?choice=No&email=${encodedEmail}`,
    yesWeek: `${WEB_APP_URL}?choice=Yes&weekly=true&email=${encodedEmail}`
  };
  
  const htmlBody = createEmailHtml(name, urls);
  const plainBody = createEmailText(name, urls);
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🍱 Lunch Reminder - Action Required',
      htmlBody: htmlBody,
      body: plainBody
    });
    console.log(`Email sent to ${name} (${email})`);
  } catch (error) {
    console.error(`Failed to send email to ${email}:`, error);
  }
}

/**
 * Handle web app requests (when users click email links)
 */
function doGet(e) {
  const { choice, email, weekly } = e.parameter;
  const isWeekly = weekly === 'true';
  const decodedEmail = decodeURIComponent(email);
  
  if (!choice || !email) {
    return ContentService.createTextOutput('Missing parameters');
  }
  
  try {
    const employeeName = findEmployeeByEmail(decodedEmail);
    
    if (!employeeName) {
      return ContentService.createTextOutput('Employee not found');
    }

    
    const success = updateLunchChoice(employeeName, choice, isWeekly);
    
    if (success) {
      return createSuccessResponse(employeeName, choice);
    } else {
      return ContentService.createTextOutput('Error updating choice');
    }
    
  } catch (error) {
    console.error("Error in doGet:", error);
    return ContentService.createTextOutput('Error: ' + error.message);
  }
}

// ==================== CORE HELPER FUNCTIONS ====================

/**
 * Find employee by email address in today's sheet
 */
function findEmployeeByEmail(email) {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const emailCol = 1; // Column B: EMAIL ADDRESS
  const nameCol = 0;  // Column A: EMPLOYEE NAME
  
  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][emailCol];
    if (rowEmail && rowEmail.toString().trim().toLowerCase() === email.toLowerCase()) {
      const employeeName = data[i][nameCol];
      return employeeName.toString().trim();
    }
  }
  
  return null;
}

/**
 * Update lunch choice in spreadsheet
 */
function updateLunchChoice(employeeName, choice, isWeekly = false) {
  const validatedChoice = (choice === 'YES' || choice === 'Yes') ? 'Yes' : 'No';
  
  if (isWeekly) {
    return updateWeeklyChoice(employeeName, validatedChoice);
  } else {
    return updateDailyChoice(employeeName, validatedChoice);
  }
}

/**
 * Update lunch choice for current day only
 */
function updateDailyChoice(employeeName, choice) {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const nameCol = 0; // Column A: EMPLOYEE NAME
  const lunchCol = 2; // Column C: TAKING LUNCH Y/N
  
  for (let i = 1; i < data.length; i++) {
    const rowName = data[i][nameCol];
    if (rowName && rowName.toString().trim().toLowerCase() === employeeName.toLowerCase()) {
      sheet.getRange(i + 1, lunchCol + 1).setValue(choice);
      SpreadsheetApp.flush();
      return true;
    }
  }
  
  return false;
}

/**
 * Update lunch choice for all weekday sheets
 */
function updateWeeklyChoice(employeeName, choice) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  let updatedCount = 0;
  
  for (let sheet of sheets) {
    const sheetName = sheet.getName();
    if (isWeekdaySheet(sheetName)) {
      const data = sheet.getDataRange().getValues();
      const nameCol = 0; // Column A
      const lunchCol = 2; // Column C
      
      for (let i = 1; i < data.length; i++) {
        const rowName = data[i][nameCol];
        if (rowName && rowName.toString().trim().toLowerCase() === employeeName.toLowerCase()) {
          sheet.getRange(i + 1, lunchCol + 1).setValue(choice);
          updatedCount++;
          break;
        }
      }
    }
  }
  
  return updatedCount > 0;
}

/**
 * Check if sheet name follows weekday pattern
 */
/**
 * Check if sheet name follows weekday pattern (Mon-Fri only)
 * More flexible version that handles different abbreviations
 */
/**
 * Check if sheet name follows weekday pattern (Mon-Fri only)
 * Updated to handle both "Thu" and "Thur" variations
 */
function isWeekdaySheet(sheetName) {
  // Allow both "Thu" and "Thur" variations
  const weekdayPattern = /^(Mon|Tue|Wed|Thu|Fri)\s+\d{1,2}(?:st|nd|rd|th)$/i;
  return weekdayPattern.test(sheetName.trim());
}

// ==================== EMAIL TEMPLATES ====================

function createEmailHtml(name, urls) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f9f9f9; margin: 0; padding: 20px; }
        .email-container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #2c3e50, #3498db); color: white; padding: 25px; text-align: center; }
        .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
        .content { padding: 30px; color: #333; }
        .greeting { font-size: 18px; margin-bottom: 20px; color: #2c3e50; }
        .instruction { color: #7f8c8d; margin-bottom: 25px; line-height: 1.6; }
        .button-container { text-align: center; margin: 30px 0; }
        .button { display: inline-block; padding: 16px 32px; margin: 10px; border: none; border-radius: 6px; font-size: 16px; font-weight: 600; text-decoration: none; cursor: pointer; transition: all 0.3s ease; min-width: 140px; text-align: center; }
        .button-yes { background: linear-gradient(135deg, #27ae60, #2ecc71); color: white; box-shadow: 0 4px 6px rgba(39, 174, 96, 0.2); }
        .button-yes:hover { background: linear-gradient(135deg, #229954, #27ae60); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(39, 174, 96, 0.3); }
        .button-no { background: linear-gradient(135deg, #e74c3c, #ec7063); color: white; box-shadow: 0 4px 6px rgba(231, 76, 60, 0.2); }
        .button-no:hover { background: linear-gradient(135deg, #c0392b, #e74c3c); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(231, 76, 60, 0.3); }
        .button-week { background: linear-gradient(135deg, #2980b9, #3498db); color: white; box-shadow: 0 4px 6px rgba(41, 128, 185, 0.2); }
        .button-week:hover { background: linear-gradient(135deg, #1c6ea4, #2980b9); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(41, 128, 185, 0.3); }
        .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #e9ecef; color: #6c757d; font-size: 14px; }
        .disclaimer { font-size: 12px; color: #95a5a6; margin-top: 20px; line-height: 1.5; }
        .company-info { margin-top: 15px; font-size: 12px; color: #7f8c8d; }
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="header"><h1>🍱 Daily Lunch Selection</h1></div>
        <div class="content">
          <div class="greeting">Hello ${name},</div>
          <p class="instruction">Please select your lunch preference for today. Your response helps us plan accordingly and minimize food waste.</p>
          <div class="button-container">
            <a href="${urls.yes}" class="button button-yes">✅ Yes</a>
            <a href="${urls.no}" class="button button-no">❌ No</a>
            <a href="${urls.yesWeek}" class="button button-week">📅 Whole Week</a>
          </div>
          <p class="instruction"><strong>Note:</strong> Clicking your selection will automatically update your preference in our system.</p>
        </div>
        <div class="footer">
          <p>This is an automated message. Please do not reply to this email.</p>
          <div class="disclaimer">If you experience issues with the buttons, please contact IT.</div>
          <div class="company-info">© ${new Date().getFullYear()} Pawa IT Solutions • HR Department</div>
        </div>
      </div>
    </body>
    </html>`;
}

function createEmailText(name, urls) {
  return `DAILY LUNCH SELECTION REQUEST

Hello ${name},

Please select your lunch preference for today by clicking one of the links below:

YES: ${urls.yes}
NO: ${urls.no}
YES FOR WHOLE WEEK: ${urls.yesWeek}

Note: This is an automated message. Please do not reply to this email.

If you experience issues with the links, please contact IT.

© ${new Date().getFullYear()} Pawa IT Solutions • HR Department`.trim();
}

function createSuccessResponse(name, choice) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        .success { color: #4CAF50; font-size: 24px; }
      </style>
    </head>
    <body>
      <div class="success">✅ Thank you, ${name}!</div>
      <p>Your lunch choice has been recorded: <strong>${choice}</strong></p>
      <p>You can close this window now.</p>
    </body>
    </html>`);
}

// ==================== TRIGGER FUNCTIONS ====================

function setupDailyTriggers() {
  deleteAllTriggers();
  
  ScriptApp.newTrigger('sendFirstReminder')
    .timeBased().everyDays(1).atHour(8).nearMinute(30).create();
  
  ScriptApp.newTrigger('sendFinalReminder')
    .timeBased().everyDays(1).atHour(9).nearMinute(0).create();
  
  ScriptApp.newTrigger('sendSupervisorReport')
    .timeBased().everyDays(1).atHour(9).nearMinute(15).create();
  
  SpreadsheetApp.getUi().alert('✅ Daily triggers set up successfully!');
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function viewCurrentTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert('No triggers found.');
    return;
  }
  
  const triggerInfo = triggers.map(trigger => 
    `Function: ${trigger.getHandlerFunction()} | Type: ${trigger.getEventType()}`
  ).join('\n');
  
  SpreadsheetApp.getUi().alert(`Current Triggers:\n\n${triggerInfo}`);
}

// ==================== AUTOMATED FUNCTIONS ====================

function sendFirstReminder() {
  const employees = getEmployeesNeedingReminders();
  if (employees.length === 0) return;
  
  employees.forEach(sendSingleEmailReminder);
  console.log(`8:30 AM: Sent ${employees.length} first reminders`);
}

function sendFinalReminder() {
  const employees = getEmployeesNeedingReminders();
  if (employees.length === 0) return;
  
  employees.forEach(employee => sendFinalReminderEmail(employee));
  console.log(`9:00 AM: Sent ${employees.length} final reminders`);
}

function sendFinalReminderEmail(employee) {
  const { name, email } = employee;
  const encodedEmail = encodeURIComponent(email);
  
  const urls = {
    yes: `${WEB_APP_URL}?choice=Yes&email=${encodedEmail}`,
    no: `${WEB_APP_URL}?choice=No&email=${encodedEmail}`,
    yesWeek: `${WEB_APP_URL}?choice=Yes&weekly=true&email=${encodedEmail}`
  };
  
  const htmlBody = createFinalReminderHtml(name, urls);
  const plainBody = createEmailText(name, urls);
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🚨 URGENT: Final Lunch Reminder - Response Needed Now',
      htmlBody: htmlBody,
      body: plainBody
    });
  } catch (error) {
    console.error(`Failed to send final reminder to ${email}:`, error);
  }
}

function createFinalReminderHtml(name, urls) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f9f9f9; margin: 0; padding: 20px; }
        .email-container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); overflow: hidden; border-left: 5px solid #e74c3c; }
        .header { background: linear-gradient(135deg, #e74c3c, #c0392b); color: white; padding: 25px; text-align: center; }
        .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
        .urgent-banner { background: #fff3cd; border: 1px solid #ffeaa7; color: #856404; padding: 15px; text-align: center; font-weight: bold; font-size: 16px; }
        .content { padding: 30px; color: #333; }
        .greeting { font-size: 18px; margin-bottom: 20px; color: #2c3e50; }
        .instruction { color: #e74c3c; margin-bottom: 25px; line-height: 1.6; font-weight: 500; }
        .button-container { text-align: center; margin: 30px 0; }
        .button { display: inline-block; padding: 16px 32px; margin: 10px; border: none; border-radius: 6px; font-size: 16px; font-weight: 600; text-decoration: none; cursor: pointer; transition: all 0.3s ease; min-width: 140px; text-align: center; }
        .button-yes { background: linear-gradient(135deg, #27ae60, #2ecc71); color: white; }
        .button-no { background: linear-gradient(135deg, #e74c3c, #ec7063); color: white; }
        .button-week { background: linear-gradient(135deg, #2980b9, #3498db); color: white; }
        .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #e9ecef; color: #6c757d; font-size: 14px; }
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="urgent-banner">⚠️ FINAL REMINDER - Response Required Immediately</div>
        <div class="header"><h1>🚨 Urgent Lunch Selection</h1></div>
        <div class="content">
          <div class="greeting">Hello ${name},</div>
          <p class="instruction"><strong>This is your FINAL reminder.</strong> We need your lunch preference immediately to complete today's catering order.</p>
          <div class="button-container">
            <a href="${urls.yes}" class="button button-yes">✅ Yes</a>
            <a href="${urls.no}" class="button button-no">❌ No</a>
            <a href="${urls.yesWeek}" class="button button-week">📅 Whole Week</a>
          </div>
        </div>
        <div class="footer">
          <p><strong>Time sensitive:</strong> Catering orders close soon!</p>
          <div>© ${new Date().getFullYear()} Pawa IT Solutions • HR Department</div>
        </div>
      </div>
    </body>
    </html>`;
}

// ==================== REPORTING FUNCTIONS ====================

function sendSupervisorReport() {
  const yesCount = getYesResponseCount();
  const totalCount = getTotalEmployeeCount();
  const pendingEmployees = getEmployeesNeedingReminders();
  
  const htmlBody = createSupervisorReportHtml(yesCount, totalCount, pendingEmployees);
  const plainBody = createSupervisorReportText(yesCount, totalCount, pendingEmployees);
  
  try {
    MailApp.sendEmail({
      to: SUPERVISOR_EMAIL,
      subject: `📊 Daily Lunch Report - ${new Date().toLocaleDateString()} - ${yesCount} Yes Responses`,
      htmlBody: htmlBody,
      body: plainBody
    });
  } catch (error) {
    console.error(`Failed to send supervisor report:`, error);
  }
}

function getYesResponseCount() {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let yesCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const lunchChoice = data[i][2];
    if (lunchChoice && lunchChoice.toString().toLowerCase() === 'yes') {
      yesCount++;
    }
  }
  
  return yesCount;
}

function getTotalEmployeeCount() {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let totalCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    if (name && email) {
      totalCount++;
    }
  }
  
  return totalCount;
}

function getEmployeesWhoSaidYes() {
  const sheetName = getTodaysSheetName();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const yesEmployees = [];
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    const lunchChoice = data[i][2];
    if (name && email && lunchChoice && lunchChoice.toString().toLowerCase() === 'yes') {
      yesEmployees.push({ name, email });
    }
  }
  
  return yesEmployees;
}

function createSupervisorReportHtml(yesCount, totalCount, pendingEmployees) {
  const responseRate = totalCount > 0 ? ((totalCount - pendingEmployees.length) / totalCount * 100).toFixed(1) : 0;
  const yesEmployees = getEmployeesWhoSaidYes();
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; }
        .report-container { max-width: 700px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #2c3e50, #34495e); color: white; padding: 25px; text-align: center; }
        .header h1 { margin: 0; font-size: 24px; }
        .summary-stats { display: flex; background: #ecf0f1; margin: 0; }
        .stat-box { flex: 1; padding: 20px; text-align: center; border-right: 1px solid #bdc3c7; }
        .stat-box:last-child { border-right: none; }
        .stat-number { font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 5px; }
        .stat-label { color: #7f8c8d; font-size: 14px; }
        .content { padding: 30px; }
        .section { margin-bottom: 25px; }
        .section h3 { color: #2c3e50; margin-bottom: 15px; padding-bottom: 5px; border-bottom: 2px solid #3498db; }
        .employee-list { background: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 15px; }
        .employee-item { padding: 5px 0; color: #2c3e50; }
        .yes-count { color: #27ae60; font-weight: bold; }
        .pending-count { color: #e74c3c; font-weight: bold; }
        .footer { background: #34495e; color: white; padding: 20px; text-align: center; font-size: 14px; }
      </style>
    </head>
    <body>
      <div class="report-container">
        <div class="header">
          <h1>📊 Daily Lunch Report</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.9;">${new Date().toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>
        <div class="summary-stats">
          <div class="stat-box"><div class="stat-number yes-count">${yesCount}</div><div class="stat-label">Lunch Orders</div></div>
          <div class="stat-box"><div class="stat-number">${totalCount - pendingEmployees.length}</div><div class="stat-label">Total Responses</div></div>
          <div class="stat-box"><div class="stat-number pending-count">${pendingEmployees.length}</div><div class="stat-label">Pending</div></div>
          <div class="stat-box"><div class="stat-number">${responseRate}%</div><div class="stat-label">Response Rate</div></div>
        </div>
        <div class="content">
          <div class="section">
            <h3>📋 Summary</h3>
            <p><strong class="yes-count">${yesCount} employees</strong> have requested lunch for today out of ${totalCount} total employees.</p>
            ${pendingEmployees.length > 0 ? `<p><strong class="pending-count">${pendingEmployees.length} employees</strong> have not yet responded.</p>` : '<p>✅ All employees have responded!</p>'}
          </div>
          ${yesEmployees.length > 0 ? `
          <div class="section">
            <h3>🍱 Employees Having Lunch (${yesCount})</h3>
            <div class="employee-list">${yesEmployees.map(emp => `<div class="employee-item">• ${emp.name}</div>`).join('')}</div>
          </div>` : ''}
          ${pendingEmployees.length > 0 ? `
          <div class="section">
            <h3>⚠️ Non-Responsive Employees (${pendingEmployees.length})</h3>
            <div class="employee-list">${pendingEmployees.map(emp => `<div class="employee-item">• ${emp.name} (${emp.email})</div>`).join('')}</div>
          </div>` : ''}
        </div>
        <div class="footer">Generated automatically • Pawa IT Solutions HR System</div>
      </div>
    </body>
    </html>`;
}

function createSupervisorReportText(yesCount, totalCount, pendingEmployees) {
  const responseRate = totalCount > 0 ? ((totalCount - pendingEmployees.length) / totalCount * 100).toFixed(1) : 0;
  const yesEmployees = getEmployeesWhoSaidYes();
  
  let report = `DAILY LUNCH REPORT - ${new Date().toLocaleDateString()}\n\n`;
  report += `SUMMARY:\n`;
  report += `- Lunch orders: ${yesCount}\n`;
  report += `- Total responses: ${totalCount - pendingEmployees.length}/${totalCount}\n`;
  report += `- Response rate: ${responseRate}%\n\n`;
  
  if (yesEmployees.length > 0) {
    report += `EMPLOYEES HAVING LUNCH (${yesCount}):\n`;
    yesEmployees.forEach(emp => report += `- ${emp.name}\n`);
    report += `\n`;
  }
  
  if (pendingEmployees.length > 0) {
    report += `NON-RESPONSIVE EMPLOYEES (${pendingEmployees.length}):\n`;
    pendingEmployees.forEach(emp => report += `- ${emp.name} (${emp.email})\n`);
  }
  
  report += `\nGenerated automatically by Pawa IT Solutions HR System`;
  return report;
}

// ==================== UI FUNCTIONS ====================

function showPendingDialog() {
  const employees = getEmployeesNeedingReminders();
  
  if (employees.length === 0) {
    SpreadsheetApp.getUi().alert('🎉 All employees have responded!');
    return;
  }
  
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Employees Pending Response: ${employees.length}</h2>
      <ul>${employees.map(emp => `<li>${emp.name} (${emp.email})</li>`).join('')}</ul>
      <button onclick="google.script.run.sendEmailReminders(); google.script.host.close();" 
              style="padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer;">
        Send Reminders Now
      </button>
    </div>
  `;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(300),
    'Pending Responses'
  );
}

// ==================== UTILITY FUNCTIONS ====================

function findColumnIndex(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toLowerCase().trim();
    for (let keyword of keywords) {
      if (header.includes(keyword.toLowerCase())) {
        return i;
      }
    }
  }
  return -1;
}