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
    // ── Handle email Approve / Reject button clicks ──────────────
    if (e && e.parameter && e.parameter.tvAction) {
      return handleTravelEmailAction(e.parameter);
    }

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

// ─── EMAIL ACTION HANDLER (Approve / Reject from email link) ──

function handleTravelEmailAction(params) {
  var action   = params.tvAction;          // 'approve' or 'reject'
  var id       = params.id    || '';
  var token    = params.token || '';

  var styleBase =
    'body{font-family:Segoe UI,Arial,sans-serif;background:#f0f2f5;display:flex;' +
    'align-items:center;justify-content:center;min-height:100vh;margin:0;}' +
    '.card{background:#fff;border-radius:14px;padding:40px;text-align:center;' +
    'max-width:420px;box-shadow:0 8px 32px rgba(0,0,0,.12);}' +
    'h2{margin:12px 0 8px;}p{color:#555;font-size:14px;line-height:1.6;}' +
    '.id{font-family:monospace;background:#f5f5f5;padding:4px 12px;border-radius:20px;' +
    'font-size:13px;display:inline-block;margin:6px 0;}';

  // ── Validate token ─────────────────────────────────────────
  var stored = PropertiesService.getScriptProperties().getProperty('tv_token_' + id);
  if (!stored) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><style>' + styleBase + '</style></head><body>' +
      '<div class="card"><div style="font-size:48px;">⚠️</div>' +
      '<h2 style="color:#c62828;">Invalid Link</h2>' +
      '<p>This approval link is invalid or has already been removed.</p></div></body></html>'
    ).setTitle('Invalid Link').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var storedObj;
  try { storedObj = JSON.parse(stored); } catch(e) { storedObj = {}; }

  if (storedObj.token !== token) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><style>' + styleBase + '</style></head><body>' +
      '<div class="card"><div style="font-size:48px;">🔒</div>' +
      '<h2 style="color:#c62828;">Security Error</h2>' +
      '<p>The token in this link does not match our records.</p></div></body></html>'
    ).setTitle('Security Error').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (Date.now() > storedObj.expires) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><style>' + styleBase + '</style></head><body>' +
      '<div class="card"><div style="font-size:48px;">⏰</div>' +
      '<h2 style="color:#e65100;">Link Expired</h2>' +
      '<p>This approval link has expired. Please action the request from inside the system.</p>' +
      '</div></body></html>'
    ).setTitle('Link Expired').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (storedObj.used) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><style>' + styleBase + '</style></head><body>' +
      '<div class="card"><div style="font-size:48px;">✅</div>' +
      '<h2 style="color:#2e7d32;">Already Processed</h2>' +
      '<p>Travel plan <span class="id">' + id + '</span> was already <strong>' +
      storedObj.decision + '</strong>.</p>' +
      '</div></body></html>'
    ).setTitle('Already Done').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Authenticate approver ──────────────────────────────────
  var approverEmail = Session.getActiveUser().getEmail();
  if (!approverEmail) {
    return HtmlService.createHtmlOutput(getLoginPage())
      .setTitle(APP_NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var approver = getUserByEmail(approverEmail);
  if (!approver || (approver.role !== 'Admin' && approver.role !== 'HR')) {
    return HtmlService.createHtmlOutput(getAccessDeniedPage(approverEmail))
      .setTitle('Access Denied').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Perform action ────────────────────────────────────────
  var newStatus = (action === 'approve') ? 'Approved' : 'Rejected';
  var notes     = (action === 'reject')  ? 'Rejected via email link' : '';
  try {
    updateTravelStatus(
      { rowIndex: storedObj.rowIndex, status: newStatus, notes: notes },
      approver
    );
    // Mark token used
    storedObj.used     = true;
    storedObj.decision = newStatus;
    PropertiesService.getScriptProperties().setProperty('tv_token_' + id, JSON.stringify(storedObj));

    var icon  = newStatus === 'Approved' ? '✅' : '❌';
    var color = newStatus === 'Approved' ? '#2e7d32' : '#c62828';
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><style>' + styleBase + '</style></head><body>' +
      '<div class="card">' +
      '<img src="https://sandalmistresort.com/wp-content/uploads/2024/09/logo-white.png"' +
      ' style="height:40px;background:#1565c0;border-radius:6px;padding:6px 12px;margin-bottom:16px;">' +
      '<div style="font-size:48px;">' + icon + '</div>' +
      '<h2 style="color:' + color + ';">Travel Plan ' + newStatus + '</h2>' +
      '<p>Travel plan <span class="id">' + id + '</span> has been ' +
      '<strong>' + newStatus.toLowerCase() + '</strong> by ' + approver.name + '.</p>' +
      '<p style="margin-top:16px;font-size:12px;color:#888;">The sales representative will be notified automatically.</p>' +
      '</div></body></html>'
    ).setTitle('Travel ' + newStatus).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput(getErrorPage(err.message))
      .setTitle('Error').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

