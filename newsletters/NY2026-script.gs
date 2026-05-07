/*********************************
 * CONFIG
 *********************************/
const SHEET_NAME         = 'contact-list';
const CONFIG_SHEET_NAME  = 'config';
const SUBJECT_CELL       = 'B1';  // Subject line will be read from config!B1

// Column indices (1-based)
const HEADER_ROW_INDEX = 1;
const FIRST_NAME_COL   = 1; // A
const EMAIL_COL        = 3; // C
const OPT_IN_COL       = 6; // F
const STATUS_COL       = 7; // G
const MAX_DRAFTS_PER_RUN = 100; // can be changed. google likely will run for ~6min before time out

// HTML template hosted on GitHub (raw URL)
const TEMPLATE_URL = 'https://raw.githubusercontent.com/elchasan/momri/refs/heads/main/newsletters/NY2026.html';

/*********************************
 * CUSTOM MENU
 *********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MOMRI Mailer')
    .addItem('Create newsletter drafts', 'createNewsletterDrafts')
    // .addItem('SEND newsletter emails (use with care)', 'sendNewsletterEmails') // uncomment when ready
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
  if (subject) {
    return subject;
  } else {
    Logger.log('No subject in ' + CONFIG_SHEET_NAME + '!' + SUBJECT_CELL + '; using default subject.');
    return 'MOMRI newsletter';
  }
}

/*********************************
 * BUILD HTML BODY FROM GITHUB TEMPLATE
 *********************************/
function buildHtmlBody(firstName) {
  var name = firstName || 'colleague';

  // Fetch HTML template from GitHub (public raw URL)
  var response = UrlFetchApp.fetch(TEMPLATE_URL);
  var html = response.getContentText();

  // Replace placeholder(s)
  html = html.replace(/{{FIRST_NAME}}/g, name);

  return html;
}

/*********************************
 * MAIN: CREATE DRAFTS
 *********************************/
function createNewsletterDrafts() {
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

  const dataRange = sheet.getRange(HEADER_ROW_INDEX + 1, 1, lastRow - HEADER_ROW_INDEX, lastCol);
  const values = dataRange.getValues();

  const subject = getNewsletterSubject();
  let draftedCount = 0;

  values.forEach((row, idx) => {

    // If we've hit the per-run limit, do nothing for the rest of the rows
  if (draftedCount >= MAX_DRAFTS_PER_RUN) return;
  
    const firstName = row[FIRST_NAME_COL - 1];  // arrays are 0-based
    const email     = row[EMAIL_COL - 1];
    const optInRaw  = row[OPT_IN_COL - 1];
    const status    = row[STATUS_COL - 1];

    const optIn = String(optInRaw).toUpperCase() === 'TRUE';

    // Skip if no email or not opted-in
    if (!email || !optIn) return;

    // Skip if already drafted or sent
    const statusUpper = String(status).toUpperCase();
    if (statusUpper === 'DRAFTED' || statusUpper === 'SENT') {
      return;
    }

    const htmlBody = buildHtmlBody(firstName);
    const plainBodyFallback = subject; // fine to reuse subject as plain text

    GmailApp.createDraft(email, subject, plainBodyFallback, {
      htmlBody: htmlBody
    });

    draftedCount++;

    // Mark as DRAFTED in column G
    sheet.getRange(HEADER_ROW_INDEX + 1 + idx, STATUS_COL).setValue('DRAFTED');
  });

  ss.toast('Created ' + draftedCount + ' draft(s).', 'MOMRI Mailer', 10);
}

/*********************************
 * SEND EMAILS (CURRENTLY NOT LINKED IN MENU)
 *********************************/

// When/if you want to send directly instead of creating drafts,
// uncomment the .addItem line in onOpen() *and* use this function.

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

  const dataRange = sheet.getRange(HEADER_ROW_INDEX + 1, 1, lastRow - HEADER_ROW_INDEX, lastCol);
  const values = dataRange.getValues();

  const subject = getNewsletterSubject();
  let sentCount = 0;

  values.forEach((row, idx) => {
    const firstName = row[FIRST_NAME_COL - 1];
    const email     = row[EMAIL_COL - 1];
    const optInRaw  = row[OPT_IN_COL - 1];
    const status    = row[STATUS_COL - 1];

    const optIn = String(optInRaw).toUpperCase() === 'TRUE';

    // Only send if opted-in and not already SENT
    if (!email || !optIn) return;
    if (String(status).toUpperCase() === 'SENT') return;

    const htmlBody = buildHtmlBody(firstName);
    const plainBodyFallback = subject;

    GmailApp.sendEmail(email, subject, plainBodyFallback, {
      htmlBody: htmlBody
    });

    sentCount++;

    // Mark as SENT in column G
    sheet.getRange(HEADER_ROW_INDEX + 1 + idx, STATUS_COL).setValue('SENT');
  });

  ss.toast('Sent ' + sentCount + ' email(s).', 'MOMRI Mailer', 10);
}
