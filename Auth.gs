/**
 * ============================================================
 * Auth.gs - Authentication & Authorization
 * ============================================================
 */

/**
 * Get user record by email from Users sheet
 */
function getUserByEmail(email) {
  if (!email) return null;

  // Check cache first
  const cache = CacheService.getScriptCache();
  const cacheKey = "user_" + email.replace(/[@.]/g, "_");
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // Name, Email, Role, Status, Department, Phone, JoinDate, LastLogin

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = String(row[1]).trim().toLowerCase();
    if (rowEmail === email.toLowerCase().trim()) {
      const user = {
        name:       String(row[0]).trim(),
        email:      String(row[1]).trim(),
        role:       String(row[2]).trim(),
        status:     String(row[3]).trim(),
        department: String(row[4]).trim(),
        phone:      String(row[5]).trim(),
        joinDate:   String(row[6]).trim(),
        lastLogin:  String(row[7]).trim(),
        rowIndex:   i + 1
      };

      // Update last login
      sheet.getRange(i + 1, 8).setValue(new Date());

      // Cache for 5 minutes
      cache.put(cacheKey, JSON.stringify(user), CACHE_EXPIRY);

      return user;
    }
  }

  return null;
}

/**
 * Check if user has required role access
 */
function hasAccess(user, requiredRoles) {
  if (!user || !user.role) return false;
  if (user.status !== "Active") return false;

  const roles = Array.isArray(requiredRoles) ? requiredRoles : [requiredRoles];

  // Admin has access to everything
  if (user.role === "Admin") return true;

  return roles.includes(user.role);
}

/**
 * Require admin role or throw error
 */
function requireAdmin(user) {
  if (!hasAccess(user, ["Admin"])) {
    throw new Error("Access denied: Admin role required");
  }
}

/**
 * Require HR or Admin role
 */
function requireHROrAdmin(user) {
  if (!hasAccess(user, ["Admin", "HR"])) {
    throw new Error("Access denied: HR or Admin role required");
  }
}

// ─── USER MANAGEMENT (Admin only) ─────────────────────────────

function getUsers(user) {
  requireAdmin(user);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      rowIndex:   i + 1,
      name:       String(data[i][0]).trim(),
      email:      String(data[i][1]).trim(),
      role:       String(data[i][2]).trim(),
      status:     String(data[i][3]).trim(),
      department: String(data[i][4]).trim(),
      phone:      String(data[i][5]).trim(),
      joinDate:   String(data[i][6]).trim()
    });
  }

  return { success: true, data: users };
}

function addUser(data, user) {
  requireAdmin(user);

  // Input validation
  if (!data.name || !data.email || !data.role) {
    return { success: false, error: "Name, Email, and Role are required" };
  }
  if (!isValidEmail(data.email)) {
    return { success: false, error: "Invalid email address" };
  }
  const validRoles = ["Admin", "HR", "Sales"];
  if (!validRoles.includes(data.role)) {
    return { success: false, error: "Invalid role" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);

  // Check duplicate
  const existing = getUserByEmail(data.email);
  if (existing) {
    return { success: false, error: "User with this email already exists" };
  }

  const newRow = [
    sanitizeInput(data.name),
    sanitizeInput(data.email).toLowerCase(),
    data.role,
    data.status || "Active",
    sanitizeInput(data.department || ""),
    sanitizeInput(data.phone || ""),
    data.joinDate || Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
    ""
  ];

  sheet.appendRow(newRow);

  // Invalidate cache
  const cache = CacheService.getScriptCache();
  cache.remove("user_" + data.email.replace(/[@.]/g, "_"));

  logActivity(user.email, "ADD_USER", "Added user: " + data.email);
  return { success: true, message: "User added successfully" };
}

function updateUser(data, user) {
  requireAdmin(user);

  if (!data.rowIndex || !data.email) {
    return { success: false, error: "Row index and email required" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);

  const rowIndex = parseInt(data.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) {
    return { success: false, error: "Invalid row index" };
  }

  sheet.getRange(rowIndex, 1, 1, 7).setValues([[
    sanitizeInput(data.name || ""),
    sanitizeInput(data.email || "").toLowerCase(),
    data.role || "Sales",
    data.status || "Active",
    sanitizeInput(data.department || ""),
    sanitizeInput(data.phone || ""),
    data.joinDate || ""
  ]]);

  // Invalidate cache
  const cache = CacheService.getScriptCache();
  cache.remove("user_" + data.email.replace(/[@.]/g, "_"));

  logActivity(user.email, "UPDATE_USER", "Updated user: " + data.email);
  return { success: true, message: "User updated successfully" };
}

// ─── EMAIL RECIPIENTS MANAGEMENT ─────────────────────────────

function getEmailRecipients(user) {
  requireHROrAdmin(user);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.EMAIL_RECIPIENTS);
  const data = sheet.getDataRange().getValues();

  const recipients = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    recipients.push({
      rowIndex: i + 1,
      type:     String(data[i][0]).trim(),
      name:     String(data[i][1]).trim(),
      email:    String(data[i][2]).trim(),
      active:   String(data[i][3]).trim()
    });
  }

  return { success: true, data: recipients };
}

function updateEmailRecipient(data, user) {
  requireHROrAdmin(user);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.EMAIL_RECIPIENTS);

  if (data.action === "add") {
    sheet.appendRow([
      sanitizeInput(data.type || "HR"),
      sanitizeInput(data.name || ""),
      sanitizeInput(data.email || "").toLowerCase(),
      "Yes"
    ]);
  } else if (data.action === "update" && data.rowIndex) {
    sheet.getRange(parseInt(data.rowIndex), 1, 1, 4).setValues([[
      sanitizeInput(data.type || "HR"),
      sanitizeInput(data.name || ""),
      sanitizeInput(data.email || "").toLowerCase(),
      data.active || "Yes"
    ]]);
  } else if (data.action === "delete" && data.rowIndex) {
    sheet.deleteRow(parseInt(data.rowIndex));
  }

  logActivity(user.email, "UPDATE_RECIPIENTS", "Updated email recipients");
  return { success: true, message: "Recipients updated" };
}

// ─── INPUT VALIDATION & SANITIZATION ─────────────────────────

function sanitizeInput(str) {
  if (typeof str !== "string") return String(str || "");
  // Remove HTML tags and dangerous characters
  return str
    .replace(/<[^>]*>/g, "")
    .replace(/['"`;]/g, "")
    .trim()
    .substring(0, 500); // Max length
}

function isValidEmail(email) {
  const re = /^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$/;
  return re.test(String(email).trim());
}

function isValidDate(dateStr) {
  const d = new Date(dateStr);
  return !isNaN(d.getTime());
}

function generateID(prefix) {
  return prefix + "-" + Utilities.formatDate(new Date(), "UTC", "yyyyMMdd") +
         "-" + Math.floor(Math.random() * 9000 + 1000);
}

// ─── LOGGING ─────────────────────────────────────────────────

function logActivity(email, action, details, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SYSTEM_LOGS);
    if (!sheet) return;

    sheet.appendRow([
      new Date(),
      email || "system",
      action || "ACTION",
      sanitizeInput(details || ""),
      "",
      status || "Success"
    ]);
  } catch (e) {
    Logger.log("Log error: " + e.message);
  }
}

function logError(context, err) {
  try {
    Logger.log("ERROR [" + context + "]: " + err.message);
    logActivity("system", "ERROR", "[" + context + "] " + err.message, "Error");
  } catch (e) {}
}

function getSystemLogs(user) {
  requireAdmin(user);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SYSTEM_LOGS);
  const data = sheet.getDataRange().getValues();

  const logs = [];
  for (let i = data.length - 1; i >= 1 && logs.length < 200; i--) {
    if (!data[i][0]) continue;
    logs.push({
      timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss") : "",
      email:     String(data[i][1]).trim(),
      action:    String(data[i][2]).trim(),
      details:   String(data[i][3]).trim(),
      status:    String(data[i][5]).trim()
    });
  }

  return { success: true, data: logs };
}

