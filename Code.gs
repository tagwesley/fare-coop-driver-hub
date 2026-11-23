const HUB_ID = 'YOUR_SPREADSHEET_ID';
const PORTAL_NAME = 'Fare Co-op Driver Hub';
const PORTAL_TIME_ZONE = 'America/Phoenix';
const PORTAL_SUPPORT_OWNER = 'support@example.org';
const PORTAL_SESSION_SECONDS = 21600; // Six hours: Apps Script cache maximum.
const PORTAL_OTP_SECONDS = 600;

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.portalName = PORTAL_NAME;
  return template.evaluate()
    .setTitle(PORTAL_NAME)
    .setFaviconUrl('https://fare.coop/wp-content/uploads/2025/08/cropped-Fare-logo-new-1.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fare Co-op Ops')
    .addItem('Validate driver portal', 'showSchemaValidation')
    .addItem('Check reminder senders', 'showReminderProviderStatus')
    .addItem('Generate driver email queue', 'generateDriverNextStepQueue')
    .addItem('Send approved email queue now', 'sendApprovedEmailQueue')
    .addSeparator()
    .addItem('Configure review-first reminder automation', 'configureReviewFirstReminderAutomation')
    .addItem('Activate driver reminder automation', 'activateDriverReminderAutomation')
    .addItem('Pause driver reminder automation', 'pauseDriverReminderAutomation')
    .addSeparator()
    .addItem('Sync tasks to Calendar', 'syncTasksToCalendar')
    .addItem('Install non-email automation triggers', 'installAutomationTriggers')
    .addToUi();
}

function mapHeaders_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function requireHeaders_(sheet, requiredHeaders) {
  if (!sheet) throw new Error('Required sheet is missing.');
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const map = mapHeaders_(headers);
  const missing = requiredHeaders.filter(header => map[header] === undefined);
  if (missing.length) {
    throw new Error(`${sheet.getName()} is missing columns: ${missing.join(', ')}`);
  }
  return map;
}

function getHub_() {
  return SpreadsheetApp.openById(HUB_ID);
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function normalizeId_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function cleanText_(value, maxLength) {
  return String(value || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, ' ').trim().slice(0, maxLength || 1000);
}

function formatDateTime_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return cleanText_(value, 100);
  return Utilities.formatDate(date, PORTAL_TIME_ZONE, 'EEE, MMM d, yyyy h:mm a');
}

function formatDate_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return cleanText_(value, 100);
  return Utilities.formatDate(date, PORTAL_TIME_ZONE, 'MMM d, yyyy');
}

function dateKey_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, PORTAL_TIME_ZONE, 'yyyy-MM-dd');
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
    .map(byte => ((byte + 256) % 256).toString(16).padStart(2, '0'))
    .join('');
}

function nowIso_() {
  return new Date().toISOString();
}
