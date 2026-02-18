/**
 * ============================================================
 * SANDAL MIST Sales & HR Management System
 * Code.gs - Main Entry Point & Web App Handler
 * ============================================================
 */

// ─── CONSTANTS ────────────────────────────────────────────────
var APP_NAME    = "Sandal Mist SHRMS";
var APP_VERSION = "1.0.0";
var DRIVE_FOLDER_NAME = "SandalMist_System";
var CACHE_EXPIRY = 300;

var SHEETS = {
  USERS:            "Users",
  DAILY_REPORTS:    "DailyReports",
  TRAVEL_PLANS:     "TravelPlans",
  LEADS:            "Leads",
  BOOKINGS:         "Bookings",
  INCENTIVES:       "Incentives",
  SYSTEM_LOGS:      "SystemLogs",
  EMAIL_RECIPIENTS: "EmailRecipients"
};

// ─── WEB APP ENTRY POINT ──────────────────────────────────────

function doGet(e) {
  try {
    // Auto-init sheets if missing
    if (!isSystemInitialized()) {
      initSystem();
    }

    var userEmail = Session.getActiveUser().getEmail();

    if (!userEmail || userEmail.trim() === '') {
      return HtmlService.createHtmlOutput(getLoginPage())
        .setTitle(APP_NAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var user = getUserByEmail(userEmail);

    if (!user) {
      return HtmlService.createHtmlOutput(getAccessDeniedPage(userEmail))
        .setTitle('Access Denied')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (user.status !== 'Active') {
      return HtmlService.createHtmlOutput(getInactivePage(user.name))
        .setTitle('Inactive Account')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    try { logActivity(userEmail, 'LOGIN', 'Accessed system'); } catch(le) {}

    var tmpl = HtmlService.createTemplateFromFile('Index');
    tmpl.userEmail  = userEmail;
    tmpl.userName   = user.name || userEmail.split('@')[0];
    tmpl.userRole   = user.role || 'Sales';
    tmpl.appName    = APP_NAME;
    tmpl.appVersion = APP_VERSION;

    return tmpl.evaluate()
      .setTitle(APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    Logger.log('doGet ERROR: ' + err.message);
    return HtmlService.createHtmlOutput(getErrorPage(err.message))
      .setTitle('Error')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ─── INCLUDE PARTIAL HTML ─────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── SYSTEM CHECK ─────────────────────────────────────────────
function isSystemInitialized() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName(SHEETS.USERS) !== null;
  } catch (e) {
    return false;
  }
}

// ─── STATIC ERROR / INFO PAGES ────────────────────────────────

function getLoginPage() {
  return '<!DOCTYPE html><html><head><title>' + APP_NAME + '</title>' +
    '<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;' +
    'min-height:100vh;margin:0;background:#1a237e;}' +
    '.card{background:#fff;border-radius:12px;padding:40px;text-align:center;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.3);}' +
    'h2{color:#1a237e;}p{color:#666;font-size:14px;}</style></head>' +
    '<body><div class="card"><div style="font-size:48px;">🏨</div>' +
    '<h2>' + APP_NAME + '</h2>' +
    '<p>Please sign in with your Google account to access this system.</p>' +
    '</div></body></html>';
}

function getAccessDeniedPage(email) {
  return '<!DOCTYPE html><html><head><title>Access Denied</title>' +
    '<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;' +
    'min-height:100vh;margin:0;background:#1a237e;}' +
    '.card{background:#fff;border-radius:12px;padding:40px;text-align:center;max-width:420px;}' +
    'h2{color:#c62828;}p{color:#555;font-size:14px;line-height:1.6;}' +
    '.em{background:#f5f5f5;padding:8px 16px;border-radius:20px;font-family:monospace;' +
    'font-size:13px;display:inline-block;margin:10px 0;}</style></head>' +
    '<body><div class="card"><div style="font-size:48px;">🚫</div>' +
    '<h2>Access Denied</h2>' +
    '<p>Your Google account is not registered in the Sandal Mist system.</p>' +
    '<div class="em">' + email + '</div>' +
    '<p>Please contact your system administrator to request access.</p>' +
    '</div></body></html>';
}

function getInactivePage(name) {
  return '<!DOCTYPE html><html><head><title>Account Inactive</title>' +
    '<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;' +
    'min-height:100vh;margin:0;background:#1a237e;}' +
    '.card{background:#fff;border-radius:12px;padding:40px;text-align:center;max-width:380px;}' +
    'h2{color:#e65100;}p{color:#555;font-size:14px;}</style></head>' +
    '<body><div class="card"><div style="font-size:48px;">⚠️</div>' +
    '<h2>Account Inactive</h2>' +
    '<p>Hello ' + (name||'User') + ', your account has been deactivated.</p>' +
    '<p>Please contact your administrator.</p>' +
    '</div></body></html>';
}

function getErrorPage(msg) {
  return '<!DOCTYPE html><html><head><title>Error</title>' +
    '<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;' +
    'min-height:100vh;margin:0;background:#1a237e;}' +
    '.card{background:#fff;border-radius:12px;padding:40px;text-align:center;max-width:500px;}' +
    'h2{color:#c62828;}pre{background:#f5f5f5;padding:12px;border-radius:6px;' +
    'font-size:12px;text-align:left;overflow:auto;white-space:pre-wrap;}</style></head>' +
    '<body><div class="card"><div style="font-size:48px;">💥</div>' +
    '<h2>System Error</h2>' +
    '<p>Something went wrong. Check the Apps Script execution log.</p>' +
    '<pre>' + msg + '</pre></div></body></html>';
}

