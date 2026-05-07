/*********************************
 * CONFIG
 *********************************/
const SHEET_NAME         = 'contact-list';
const CONFIG_SHEET_NAME  = 'config';
const SUBJECT_CELL       = 'B2';  // Subject line read from config!B2

// Column indices (1-based)
const HEADER_ROW_INDEX = 1;
const FIRST_NAME_COL   = 1; // A
const EMAIL_COL        = 3; // C
const YES_COL          = 5; // E must say YES
const TRUE_COL         = 8; // H must say TRUE
const STATUS_COL       = 9; // I will be populated with SENT + timestamp

// Safety cap: change upward if needed
const MAX_SENDS_PER_RUN = 100;

// HTML template hosted on GitHub
const TEMPLATE_URL = 'https://raw.githubusercontent.com/elchasan/momri/refs/heads/main/newsletters/Sp2026.html';

/*********************************
 * CUSTOM MENU
 *********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MOMRI Mailer')
    .addItem('Send newsletter emails', 'sendNewsletterEmails')
    .addToUi();
}

/*********************************
 * READ SUBJECT FROM CONFIG SHEET
 *********************************/
function getNewsletterSubject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!cfg) {
    Logger.log('Config sheet not found; using default subject.');
    return 'MOMRI newsletter';
  }

  const subject = cfg.getRange(SUBJECT_CELL).getDisplayValue().trim();
  return subject || 'MOMRI newsletter';
}

/*********************************
 * FETCH HTML TEMPLATE ONCE
 *********************************/
function getHtmlTemplate() {
  const response = UrlFetchApp.fetch(TEMPLATE_URL);
  return response.getContentText();
}

/*********************************
 * BUILD HTML BODY
 *********************************/
function buildHtmlBody(templateHtml, firstName) {
  const name = firstName || 'colleague';
  return templateHtml.replace(/{{FIRST_NAME}}/g, name);
}

/*********************************
 * TIMESTAMP
 *********************************/
function getSentTimestamp() {
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  return Utilities.formatDate(now, timezone, "yyyy-MM-dd HH:mm:ss z");
}

/*********************************
 * MAIN: SEND EMAILS
 *********************************/
function sendNewsletterEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= HEADER_ROW_INDEX) {
    Logger.log('No data rows found.');
    return;
  }

  const values = sheet
    .getRange(HEADER_ROW_INDEX + 1, 1, lastRow - HEADER_ROW_INDEX, lastCol)
    .getValues();

  const subject = getNewsletterSubject();
  const templateHtml = getHtmlTemplate();

  let sentCount = 0;

  values.forEach((row, idx) => {
    if (sentCount >= MAX_SENDS_PER_RUN) return;

    const firstName = row[FIRST_NAME_COL - 1];
    const email     = row[EMAIL_COL - 1];
    const yesRaw    = row[YES_COL - 1];
    const trueRaw   = row[TRUE_COL - 1];
    const status    = row[STATUS_COL - 1];

    const hasYes  = String(yesRaw).trim().toUpperCase() === 'YES';
    const hasTrue = String(trueRaw).trim().toUpperCase() === 'TRUE';

    // Skip if no email, conditions not met, or already sent
    if (!email || !hasYes || !hasTrue) return;
    if (String(status).trim().toUpperCase().startsWith('SENT')) return;

    const htmlBody = buildHtmlBody(templateHtml, firstName);
    const plainBodyFallback = subject;

    GmailApp.sendEmail(email, subject, plainBodyFallback, {
      htmlBody: htmlBody
    });

    sentCount++;

    const timestamp = getSentTimestamp();

    // Mark as SENT with timestamp in column I
    sheet
      .getRange(HEADER_ROW_INDEX + 1 + idx, STATUS_COL)
      .setValue('SENT - ' + timestamp);
  });

  ss.toast('Sent ' + sentCount + ' email(s).', 'MOMRI Mailer', 10);
}
