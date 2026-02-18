/**
 * ============================================================
 * Init.gs - System Initialization & Setup
 * ============================================================
 */

/**
 * initSystem() - Run this ONCE to set up the entire system
 * Creates all sheets, demo data, triggers, and Drive folders.
 */
function initSystem() {
  const ui = SpreadsheetApp.getUi ? SpreadsheetApp.getUi() : null;

  try {
    Logger.log("=== SANDAL MIST SYSTEM INITIALIZATION ===");

    // Step 1: Create Drive Folder
    Logger.log("Step 1: Creating Drive folder...");
    const folder = createSystemFolder();

    // Step 2: Create all sheets
    Logger.log("Step 2: Creating sheets...");
    createAllSheets();

    // Step 3: Populate demo data
    Logger.log("Step 3: Adding demo data...");
    addDemoData();

    // Step 4: Set up triggers
    Logger.log("Step 4: Setting up triggers...");
    setupTriggers();

    // Step 5: Apply formatting
    Logger.log("Step 5: Formatting sheets...");
    formatAllSheets();

    Logger.log("=== INITIALIZATION COMPLETE ===");

    if (ui) {
      ui.alert("✅ System Initialized!", 
        "Sandal Mist SHRMS has been set up successfully.\n\n" +
        "✓ All sheets created\n✓ Demo data loaded\n✓ Triggers configured\n" +
        "✓ Drive folder created\n\nDeploy the web app to share with your team.",
        ui.ButtonSet.OK);
    }

    return { success: true, message: "System initialized successfully" };

  } catch (err) {
    Logger.log("INIT ERROR: " + err.message);
    if (ui) ui.alert("❌ Error", err.message, ui.ButtonSet.OK);
    throw err;
  }
}

// ─── DRIVE FOLDER SETUP ───────────────────────────────────────

function createSystemFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();

  const root = DriveApp.getRootFolder();
  const mainFolder = root.createFolder(DRIVE_FOLDER_NAME);
  mainFolder.createFolder("DSR_PDFs");
  mainFolder.createFolder("Reports");
  mainFolder.createFolder("Travel_Documents");

  PropertiesService.getScriptProperties().setProperty("DRIVE_FOLDER_ID", mainFolder.getId());
  Logger.log("Drive folder created: " + mainFolder.getId());
  return mainFolder;
}

function getSystemFolder() {
  const folderId = PropertiesService.getScriptProperties().getProperty("DRIVE_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) {}
  }
  return createSystemFolder();
}

function getSubFolder(name) {
  const parent = getSystemFolder();
  const subs = parent.getFoldersByName(name);
  if (subs.hasNext()) return subs.next();
  return parent.createFolder(name);
}

// ─── SHEET CREATION ───────────────────────────────────────────

function createAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDefs = [
    {
      name: SHEETS.USERS,
      headers: ["Name", "Email", "Role", "Status", "Department", "Phone", "JoinDate", "LastLogin"],
      color: "#4285F4"
    },
    {
      name: SHEETS.DAILY_REPORTS,
      headers: ["ID", "Date", "SalesRep", "Email", "Location", "ClientName", "ContactNo",
                "Purpose", "Result", "NextAction", "ProofLink", "PDFLink", "Status", "SubmittedAt"],
      color: "#34A853"
    },
    {
      name: SHEETS.TRAVEL_PLANS,
      headers: ["ID", "SalesRep", "Email", "TravelDate", "ReturnDate", "City", "Clients",
                "Purpose", "ExpectedRevenue", "EstimatedDays", "Transport", "Accommodation",
                "Status", "ApprovedBy", "ApprovedAt", "Notes", "SubmittedAt"],
      color: "#FBBC04"
    },
    {
      name: SHEETS.LEADS,
      headers: ["ID", "Date", "LeadSource", "ClientName", "ContactPerson", "Phone", "Email",
                "PropertyType", "Budget", "AssignedTo", "Status", "FollowUpDate",
                "Notes", "Value", "ConversionDate", "CreatedBy", "CreatedAt"],
      color: "#EA4335"
    },
    {
      name: SHEETS.BOOKINGS,
      headers: ["ID", "BookingDate", "ClientName", "Phone", "Email", "PropertyName",
                "PropertyType", "CheckIn", "CheckOut", "Nights", "RoomType", "RatePerNight",
                "TotalValue", "Commission", "LeadID", "SalesRep", "Status",
                "PaymentStatus", "Notes", "CreatedAt"],
      color: "#9C27B0"
    },
    {
      name: SHEETS.INCENTIVES,
      headers: ["ID", "SalesRep", "Email", "Month", "Year", "TotalSales",
                "BaseThreshold", "EligibleAmount", "IncentiveRate", "IncentiveAmount",
                "Status", "PaidDate", "CalculatedAt"],
      color: "#FF6D00"
    },
    {
      name: SHEETS.SYSTEM_LOGS,
      headers: ["Timestamp", "UserEmail", "Action", "Details", "IPAddress", "Status"],
      color: "#607D8B"
    },
    {
      name: SHEETS.EMAIL_RECIPIENTS,
      headers: ["Type", "Name", "Email", "Active"],
      color: "#00BCD4"
    }
  ];

  sheetDefs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      Logger.log("Created sheet: " + def.name);
    }
    setupSheetHeaders(sheet, def.headers, def.color);
  });
}

function setupSheetHeaders(sheet, headers, color) {
  const header = sheet.getRange(1, 1, 1, headers.length);
  header.setValues([headers]);
  header.setBackground(color);
  header.setFontColor("#FFFFFF");
  header.setFontWeight("bold");
  header.setFontSize(10);
  header.setFontFamily("Google Sans, Arial, sans-serif");
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 140);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length)
       .setVerticalAlignment("middle");
}

// ─── DEMO DATA ────────────────────────────────────────────────

function addDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Users ──
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    const users = [
      ["Administrator",   "admin@sandalm ist.com",  "Admin", "Active", "Management",  "+60-12-000-0001", "2023-01-01", ""],
      ["HR Manager",      "hr@sandalmist.com",       "HR",    "Active", "HR",          "+60-12-000-0002", "2023-01-15", ""],
      ["HR Assistant",    "hr2@sandalmist.com",      "HR",    "Active", "HR",          "+60-12-000-0003", "2023-02-01", ""],
      ["Ahmad Faris",     "ahmad@sandalmist.com",    "Sales", "Active", "Sales",       "+60-12-111-1111", "2023-03-01", ""],
      ["Nurul Ain",       "nurul@sandalmist.com",    "Sales", "Active", "Sales",       "+60-12-222-2222", "2023-03-15", ""],
      ["Razif Haron",     "razif@sandalmist.com",    "Sales", "Active", "Sales",       "+60-12-333-3333", "2023-04-01", ""],
      ["Siti Zulaikha",   "siti@sandalmist.com",     "Sales", "Inactive","Sales",      "+60-12-444-4444", "2023-04-15", ""]
    ];
    usersSheet.getRange(2, 1, users.length, 8).setValues(users);
  }

  // ── Email Recipients ──
  const emailSheet = ss.getSheetByName(SHEETS.EMAIL_RECIPIENTS);
  if (emailSheet.getLastRow() <= 1) {
    const recipients = [
      ["HR",      "HR Manager",   "hr@sandalmist.com",   "Yes"],
      ["HR",      "HR Assistant", "hr2@sandalmist.com",  "Yes"],
      ["Admin",   "Admin",        "admin@sandalmist.com","Yes"],
      ["Report",  "Director",     "director@sandalmist.com","Yes"]
    ];
    emailSheet.getRange(2, 1, recipients.length, 4).setValues(recipients);
  }

  // ── Daily Reports (Sample) ──
  const dsrSheet = ss.getSheetByName(SHEETS.DAILY_REPORTS);
  if (dsrSheet.getLastRow() <= 1) {
    const today = new Date();
    const dsr = [
      ["DSR-001", Utilities.formatDate(today, "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
       "Ahmad Faris", "ahmad@sandalmist.com", "Kuala Lumpur", "Berjaya Hotels Group",
       "+60-3-2141-8888", "Property Presentation", "Positive - site visit scheduled",
       "Follow up next week", "", "", "Submitted", Utilities.formatDate(today, "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")],
      ["DSR-002", Utilities.formatDate(new Date(today - 86400000), "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
       "Nurul Ain", "nurul@sandalmist.com", "Penang", "Eastern & Oriental Hotel",
       "+60-4-222-2000", "Contract Renewal Discussion", "Contract renewed - RM85,000",
       "Send invoice", "", "", "Submitted", Utilities.formatDate(today, "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")]
    ];
    dsrSheet.getRange(2, 1, dsr.length, 14).setValues(dsr);
  }

  // ── Travel Plans (Sample) ──
  const travelSheet = ss.getSheetByName(SHEETS.TRAVEL_PLANS);
  if (travelSheet.getLastRow() <= 1) {
    const nextWeek = new Date(Date.now() + 7 * 86400000);
    const travel = [
      ["TP-001", "Ahmad Faris", "ahmad@sandalmist.com",
       Utilities.formatDate(nextWeek, "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
       Utilities.formatDate(new Date(nextWeek.getTime() + 2*86400000), "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
       "Johor Bahru", "Anantara Desaru, Lotus Desaru", "New Property Partnership",
       120000, 3, "Flight", "Client-Hosted", "Pending", "", "", "", 
       Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")]
    ];
    travelSheet.getRange(2, 1, travel.length, 17).setValues(travel);
  }

  // ── Leads (Sample) ──
  const leadsSheet = ss.getSheetByName(SHEETS.LEADS);
  if (leadsSheet.getLastRow() <= 1) {
    const leads = [
      ["LEAD-001", "2024-01-10", "Referral", "Grand Hyatt KL", "Mr. Tan Wei Ming",
       "+60-3-2182-1234", "tan@grandhyattkl.com", "5-Star Hotel", 250000,
       "Ahmad Faris", "Qualified", "2024-02-01", "Interested in full package", 250000, "", "Ahmad Faris",
       "2024-01-10 09:00:00"],
      ["LEAD-002", "2024-01-15", "Cold Call", "The Majestic Hotel KL", "Ms. Rose Abdullah",
       "+60-3-2785-8000", "rose@majestichotel.com", "Heritage Hotel", 180000,
       "Nurul Ain", "Proposal Sent", "2024-02-05", "Requesting revised proposal", 180000, "", "Nurul Ain",
       "2024-01-15 14:30:00"],
      ["LEAD-003", "2024-01-20", "Exhibition", "Sunway Resort Hotel", "Mr. Lim Kah Hoong",
       "+60-3-7492-8000", "lim@sunwayresort.com", "Resort", 320000,
       "Razif Haron", "New", "2024-02-10", "Met at MATTA Fair", 320000, "", "Razif Haron",
       "2024-01-20 10:15:00"]
    ];
    leadsSheet.getRange(2, 1, leads.length, 17).setValues(leads);
  }

  // ── Bookings (Sample) ──
  const bookingsSheet = ss.getSheetByName(SHEETS.BOOKINGS);
  if (bookingsSheet.getLastRow() <= 1) {
    const bookings = [
      ["BK-001", "2024-01-05", "Petronas MICE Division", "+60-3-2051-8888",
       "mice@petronas.com", "The Majestic Hotel KL", "Meeting Room Package",
       "2024-02-15", "2024-02-17", 2, "Deluxe", 850, 1700, 170,
       "LEAD-002", "Nurul Ain", "Confirmed", "Paid", "Annual sales conference", "2024-01-05 11:00:00"],
      ["BK-002", "2024-01-08", "Axiata Group Corporate", "+60-3-5000-2266",
       "corp@axiata.com", "Sunway Resort Hotel", "Executive Suite",
       "2024-02-20", "2024-02-23", 3, "Executive Suite", 1200, 3600, 360,
       "", "Razif Haron", "Confirmed", "Deposit Paid", "Board retreat", "2024-01-08 15:30:00"]
    ];
    bookingsSheet.getRange(2, 1, bookings.length, 20).setValues(bookings);
  }

  // ── Incentives (Sample) ──
  const incentivesSheet = ss.getSheetByName(SHEETS.INCENTIVES);
  if (incentivesSheet.getLastRow() <= 1) {
    const incentives = [
      ["INC-001", "Ahmad Faris", "ahmad@sandalmist.com", "January", 2024,
       680000, 500000, 180000, 0.01, 1800, "Paid", "2024-02-10", "2024-02-01 09:00:00"],
      ["INC-002", "Nurul Ain", "nurul@sandalmist.com", "January", 2024,
       520000, 500000, 20000, 0.01, 200, "Paid", "2024-02-10", "2024-02-01 09:00:00"],
      ["INC-003", "Razif Haron", "razif@sandalmist.com", "January", 2024,
       490000, 500000, 0, 0.01, 0, "Not Eligible", "", "2024-02-01 09:00:00"]
    ];
    incentivesSheet.getRange(2, 1, incentives.length, 13).setValues(incentives);
  }

  // ── System Logs (Initial) ──
  const logsSheet = ss.getSheetByName(SHEETS.SYSTEM_LOGS);
  if (logsSheet.getLastRow() <= 1) {
    const logRow = [[
      new Date(), "system@sandalmist.com", "INIT",
      "System initialized successfully", "127.0.0.1", "Success"
    ]];
    logsSheet.getRange(2, 1, 1, 6).setValues(logRow);
  }
}

// ─── TRIGGERS SETUP ───────────────────────────────────────────

function setupTriggers() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Monthly incentive calculation - 1st of each month at 8am
  ScriptApp.newTrigger("runMonthlyIncentiveCalc")
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  // Monthly report email - last day check (28th to cover all months)
  ScriptApp.newTrigger("sendMonthlyReportEmail")
    .timeBased()
    .onMonthDay(28)
    .atHour(18)
    .create();

  // Daily cleanup log - run at midnight
  ScriptApp.newTrigger("dailyMaintenance")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  Logger.log("Triggers configured successfully");
}

// ─── SHEET FORMATTING ─────────────────────────────────────────

function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(SHEETS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Alternate row colors
    const lastRow = Math.max(sheet.getLastRow(), 2);
    for (let i = 2; i <= Math.min(lastRow, 100); i++) {
      const color = i % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
      sheet.getRange(i, 1, 1, sheet.getLastColumn() || 1).setBackground(color);
    }

    // Auto-resize columns
    sheet.autoResizeColumns(1, sheet.getLastColumn());

    // Set row height
    sheet.setRowHeights(1, Math.max(sheet.getLastRow(), 1), 30);
  });
}

// ─── SCHEDULED TASKS ─────────────────────────────────────────

function runMonthlyIncentiveCalc() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthName = Utilities.formatDate(lastMonth, "UTC", "MMMM");
  const year = lastMonth.getFullYear();

  Logger.log("Running monthly incentive calc for " + monthName + " " + year);
  calculateMonthlyIncentives(monthName, year);
  logActivity("system", "MONTHLY_CALC", "Incentive calculation for " + monthName + " " + year);
}

function sendMonthlyReportEmail() {
  try {
    const now = new Date();
    if (now.getDate() !== 28) return; // Simplistic check

    const recipients = getHREmailList();
    if (recipients.length === 0) return;

    const stats = getDashboardStats({ role: "Admin", email: "system" });
    if (!stats.success) return;

    const html = buildMonthlyReportEmail(stats.data);
    const subject = APP_NAME + " - Monthly Summary Report";

    recipients.forEach(email => {
      GmailApp.sendEmail(email, subject, "Please view in HTML", { htmlBody: html });
    });

    logActivity("system", "MONTHLY_EMAIL", "Monthly report sent to " + recipients.join(", "));
  } catch (err) {
    logError("sendMonthlyReportEmail", err);
  }
}

function dailyMaintenance() {
  // Clean old logs (keep last 1000 rows)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEETS.SYSTEM_LOGS);
  if (logSheet && logSheet.getLastRow() > 1001) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1001);
  }
  Logger.log("Daily maintenance completed");
}

