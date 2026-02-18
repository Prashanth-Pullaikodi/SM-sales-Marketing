/**
 * ============================================================
 * Database.gs - Data Operations for All Modules
 * ============================================================
 */

// ─── DAILY SALES REPORTS (DSR) ────────────────────────────────

function submitDailyReport(data, user) {
  // Input validation
  const required = ["date", "location", "clientName", "purpose", "result"];
  for (const field of required) {
    if (!data[field]) return { success: false, error: "Field required: " + field };
  }
  if (!isValidDate(data.date)) return { success: false, error: "Invalid date" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DAILY_REPORTS);

  const id = generateID("DSR");
  const now = new Date();

  // Handle proof upload (base64 → Drive)
  let proofLink = "";
  if (data.proofBase64 && data.proofFileName) {
    proofLink = saveFileToDrive(data.proofBase64, data.proofFileName, "DSR_PDFs");
  }

  const row = [
    id,
    data.date,
    user.name,
    user.email,
    sanitizeInput(data.location),
    sanitizeInput(data.clientName),
    sanitizeInput(data.contact || ""),
    sanitizeInput(data.purpose),
    sanitizeInput(data.result),
    sanitizeInput(data.nextAction || ""),
    proofLink,
    "",  // PDF link - filled after PDF generation
    "Submitted",
    Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
  ];

  sheet.appendRow(row);
  const newRowIndex = sheet.getLastRow();

  // Generate PDF
  try {
    const pdfLink = generateDSRPdf(id, data, user);
    sheet.getRange(newRowIndex, 12).setValue(pdfLink);

    // Send email to HR
    sendDSREmail(id, data, user, pdfLink);
  } catch (emailErr) {
    Logger.log("DSR post-processing error: " + emailErr.message);
  }

  logActivity(user.email, "SUBMIT_DSR", "Submitted DSR: " + id);
  return { success: true, message: "Daily report submitted successfully", id: id };
}

function getDailyReports(user, filters) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEETS.DAILY_REPORTS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [], count: 0, total: 0, totalPages: 0 };

  // Read only needed columns (14) instead of the entire wide sheet
  const raw  = sheet.getRange(1, 1, lastRow, 14).getValues();
  const PAGE = parseInt((filters && filters.pageSize) || 200);
  const page = parseInt((filters && filters.page)     || 0);

  const mFlt = filters && filters.month ? parseInt(filters.month) : 0;
  const yFlt = filters && filters.year  ? parseInt(filters.year)  : 0;

  // Pass 1 — collect matching row indices only (fast, no object allocation)
  const matchIdx = [];
  for (let i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    if (user.role === "Sales" && String(raw[i][3]).toLowerCase() !== user.email.toLowerCase()) continue;
    if (mFlt && yFlt) {
      const d = new Date(raw[i][1]);
      if (d.getMonth() + 1 !== mFlt || d.getFullYear() !== yFlt) continue;
    }
    matchIdx.push(i);
  }

  const total      = matchIdx.length;
  const totalPages = total > 0 ? Math.ceil(total / PAGE) : 0;

  // Reverse (newest first) then slice the requested page
  matchIdx.reverse();
  const pageIdx = matchIdx.slice(page * PAGE, (page + 1) * PAGE);

  // Pass 2 — build objects only for the current page slice
  const reports = pageIdx.map(function(i) {
    const r = raw[i];
    return {
      id:          String(r[0]).trim(),
      date:        String(r[1]).trim(),
      salesRep:    String(r[2]).trim(),
      email:       String(r[3]).trim(),
      location:    String(r[4]).trim(),
      clientName:  String(r[5]).trim(),
      contact:     String(r[6]).trim(),
      purpose:     String(r[7]).trim(),
      result:      String(r[8]).trim(),
      nextAction:  String(r[9]).trim(),
      proofLink:   String(r[10]).trim(),
      pdfLink:     String(r[11]).trim(),
      status:      String(r[12]).trim(),
      submittedAt: String(r[13]).trim()
    };
  });

  return { success: true, data: reports, count: reports.length, total, page, pageSize: PAGE, totalPages };
}

// ─── TRAVEL PLANS ─────────────────────────────────────────────

function submitTravelPlan(data, user) {
  const required = ["travelDate", "city", "purpose", "expectedRevenue", "estimatedDays"];
  for (const f of required) {
    if (!data[f]) return { success: false, error: "Field required: " + f };
  }
  if (parseFloat(data.expectedRevenue) < 0) {
    return { success: false, error: "Invalid expected revenue" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TRAVEL_PLANS);

  const id = generateID("TP");
  const now = new Date();

  const row = [
    id,
    user.name,
    user.email,
    data.travelDate,
    data.returnDate || "",
    sanitizeInput(data.city),
    sanitizeInput(data.clients || ""),
    sanitizeInput(data.purpose),
    parseFloat(data.expectedRevenue) || 0,
    parseInt(data.estimatedDays) || 1,
    sanitizeInput(data.transport || ""),
    sanitizeInput(data.accommodation || ""),
    "Pending",
    "", "", // approvedBy, approvedAt
    sanitizeInput(data.notes || ""),
    Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
  ];

  sheet.appendRow(row);
  const newRowIndex = sheet.getLastRow();

  // Email HR for approval (pass rowIndex so email buttons work directly)
  try {
    sendTravelApprovalEmail(id, data, user, newRowIndex);
  } catch (e) {
    Logger.log("Travel email error: " + e.message);
  }

  logActivity(user.email, "SUBMIT_TRAVEL", "Submitted travel plan: " + id);
  return { success: true, message: "Travel plan submitted. Pending HR approval.", id: id };
}

function getTravelPlans(user, filters) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEETS.TRAVEL_PLANS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [], count: 0, total: 0, totalPages: 0 };

  const raw  = sheet.getRange(1, 1, lastRow, 17).getValues();
  const PAGE = parseInt((filters && filters.pageSize) || 200);
  const page = parseInt((filters && filters.page)     || 0);
  const sFlt      = filters && filters.status   ? String(filters.status).toLowerCase() : '';
  const dateFrom  = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
  const dateTo    = filters && filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  const matchIdx = [];
  for (let i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    if (user.role === "Sales" && String(raw[i][2]).toLowerCase() !== user.email.toLowerCase()) continue;
    if (sFlt && String(raw[i][12]).toLowerCase() !== sFlt) continue;
    if (dateFrom || dateTo) {
      const d = new Date(raw[i][3]);
      if (isNaN(d)) { matchIdx.push(i); continue; }
      if (dateFrom && d < dateFrom) continue;
      if (dateTo   && d > dateTo)   continue;
    }
    matchIdx.push(i);
  }

  const total      = matchIdx.length;
  const totalPages = total > 0 ? Math.ceil(total / PAGE) : 0;

  matchIdx.reverse();
  const pageIdx = matchIdx.slice(page * PAGE, (page + 1) * PAGE);

  const plans = pageIdx.map(function(i) {
    const r = raw[i];
    return {
      id:              String(r[0]).trim(),
      salesRep:        String(r[1]).trim(),
      email:           String(r[2]).trim(),
      travelDate:      String(r[3]).trim(),
      returnDate:      String(r[4]).trim(),
      city:            String(r[5]).trim(),
      clients:         String(r[6]).trim(),
      purpose:         String(r[7]).trim(),
      expectedRevenue: r[8],
      estimatedDays:   r[9],
      transport:       String(r[10]).trim(),
      accommodation:   String(r[11]).trim(),
      status:          String(r[12]).trim(),
      approvedBy:      String(r[13]).trim(),
      approvedAt:      String(r[14]).trim(),
      notes:           String(r[15]).trim(),
      submittedAt:     String(r[16]).trim(),
      rowIndex:        i + 1
    };
  });

  return { success: true, data: plans, count: plans.length, total, page, pageSize: PAGE, totalPages };
}

function updateTravelStatus(data, user) {
  requireHROrAdmin(user);

  if (!data.rowIndex || !data.status) return { success: false, error: "Row and status required" };

  const validStatuses = ["Approved", "Rejected", "Pending"];
  if (!validStatuses.includes(data.status)) return { success: false, error: "Invalid status" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.TRAVEL_PLANS);
  const row = parseInt(data.rowIndex);

  sheet.getRange(row, 13).setValue(data.status);
  sheet.getRange(row, 14).setValue(user.name);
  sheet.getRange(row, 15).setValue(new Date());

  if (data.notes) sheet.getRange(row, 16).setValue(sanitizeInput(data.notes));

  // Notify the sales rep
  try {
    const rowData = sheet.getRange(row, 1, 1, 17).getValues()[0];
    sendTravelStatusEmail(rowData, data.status, user, data.notes);
  } catch (e) {
    Logger.log("Travel status email error: " + e.message);
  }

  logActivity(user.email, "UPDATE_TRAVEL_STATUS",
    "Travel " + data.rowIndex + " marked " + data.status);
  return { success: true, message: "Travel plan " + data.status };
}

// ─── LEADS ────────────────────────────────────────────────────

function submitLead(data, user) {
  const required = ["clientName", "leadSource", "status"];
  for (const f of required) {
    if (!data[f]) return { success: false, error: "Field required: " + f };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LEADS);

  const id = generateID("LEAD");
  const today = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd");

  const row = [
    id,
    today,
    sanitizeInput(data.leadSource),
    sanitizeInput(data.clientName),
    sanitizeInput(data.contactPerson || ""),
    sanitizeInput(data.phone || ""),
    sanitizeInput(data.email || ""),
    sanitizeInput(data.propertyType || ""),
    parseFloat(data.budget) || 0,
    sanitizeInput(data.assignedTo || user.name),
    sanitizeInput(data.status),
    data.followUpDate || "",
    sanitizeInput(data.notes || ""),
    parseFloat(data.value) || 0,
    "",
    user.name,
    Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
  ];

  sheet.appendRow(row);

  logActivity(user.email, "SUBMIT_LEAD", "New lead: " + id + " - " + data.clientName);
  return { success: true, message: "Lead added successfully", id: id };
}

function getLeads(user, filters) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEETS.LEADS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [], count: 0, total: 0, totalPages: 0 };

  const raw  = sheet.getRange(1, 1, lastRow, 17).getValues();
  const PAGE = parseInt((filters && filters.pageSize) || 200);
  const page = parseInt((filters && filters.page)     || 0);
  const sFlt     = filters && filters.status   ? String(filters.status) : '';
  const dateFrom = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
  const dateTo   = filters && filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  const matchIdx = [];
  for (let i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    if (user.role === "Sales" && String(raw[i][9]).toLowerCase() !== user.name.toLowerCase()) continue;
    if (sFlt && raw[i][10] !== sFlt) continue;
    if (dateFrom || dateTo) {
      const d = new Date(raw[i][1]);
      if (isNaN(d)) { matchIdx.push(i); continue; }
      if (dateFrom && d < dateFrom) continue;
      if (dateTo   && d > dateTo)   continue;
    }
    matchIdx.push(i);
  }

  const total      = matchIdx.length;
  const totalPages = total > 0 ? Math.ceil(total / PAGE) : 0;

  matchIdx.reverse();
  const pageIdx = matchIdx.slice(page * PAGE, (page + 1) * PAGE);

  const leads = pageIdx.map(function(i) {
    const r = raw[i];
    return {
      id:             String(r[0]).trim(),
      date:           String(r[1]).trim(),
      leadSource:     String(r[2]).trim(),
      clientName:     String(r[3]).trim(),
      contactPerson:  String(r[4]).trim(),
      phone:          String(r[5]).trim(),
      email:          String(r[6]).trim(),
      propertyType:   String(r[7]).trim(),
      budget:         r[8],
      assignedTo:     String(r[9]).trim(),
      status:         String(r[10]).trim(),
      followUpDate:   String(r[11]).trim(),
      notes:          String(r[12]).trim(),
      value:          r[13],
      conversionDate: String(r[14]).trim(),
      createdBy:      String(r[15]).trim(),
      createdAt:      String(r[16]).trim(),
      rowIndex:       i + 1
    };
  });

  return { success: true, data: leads, count: leads.length, total, page, pageSize: PAGE, totalPages };
}

function updateLeadStatus(data, user) {
  if (!hasAccess(user, ["Admin", "HR"])) {
    // Sales can update their own leads' status
    if (user.role === "Sales" && data.ownLead) {
      // Allowed
    } else {
      return { success: false, error: "Insufficient permissions" };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LEADS);
  const row = parseInt(data.rowIndex);

  if (data.status) sheet.getRange(row, 11).setValue(sanitizeInput(data.status));
  if (data.notes) sheet.getRange(row, 13).setValue(sanitizeInput(data.notes));
  if (data.value) sheet.getRange(row, 14).setValue(parseFloat(data.value) || 0);
  if (data.followUpDate) sheet.getRange(row, 12).setValue(data.followUpDate);

  if (data.status === "Won") {
    sheet.getRange(row, 15).setValue(
      Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd")
    );
  }

  logActivity(user.email, "UPDATE_LEAD",
    "Lead " + (data.leadId || row) + " → " + data.status);
  return { success: true, message: "Lead updated" };
}

// ─── BOOKINGS ─────────────────────────────────────────────────

function submitBooking(data, user) {
  requireHROrAdmin(user);

  const required = ["clientName", "propertyName", "checkIn", "checkOut", "totalValue"];
  for (const f of required) {
    if (!data[f]) return { success: false, error: "Field required: " + f };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BOOKINGS);

  const id = generateID("BK");
  const nights = Math.ceil(
    (new Date(data.checkOut) - new Date(data.checkIn)) / 86400000
  );
  const commission = (parseFloat(data.totalValue) || 0) * 0.10; // 10%

  const row = [
    id,
    Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd"),
    sanitizeInput(data.clientName),
    sanitizeInput(data.phone || ""),
    sanitizeInput(data.email || ""),
    sanitizeInput(data.propertyName),
    sanitizeInput(data.propertyType || ""),
    data.checkIn,
    data.checkOut,
    nights,
    sanitizeInput(data.roomType || ""),
    parseFloat(data.ratePerNight) || 0,
    parseFloat(data.totalValue) || 0,
    commission,
    sanitizeInput(data.leadId || ""),
    sanitizeInput(data.salesRep || user.name),
    sanitizeInput(data.status || "Confirmed"),
    sanitizeInput(data.paymentStatus || "Pending"),
    sanitizeInput(data.notes || ""),
    Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
  ];

  sheet.appendRow(row);

  // Touch timestamp so live-poll clients know to refresh
  PropertiesService.getScriptProperties().setProperty(
    'bookings_modified', new Date().toISOString()
  );

  logActivity(user.email, "SUBMIT_BOOKING", "Booking created: " + id);
  return { success: true, message: "Booking created successfully", id: id };
}

function getLastModified(data) {
  var sheet = (data && data.sheet) || 'bookings';
  var key = sheet + '_modified';
  var ts  = PropertiesService.getScriptProperties().getProperty(key) || '';
  return { success: true, timestamp: ts };
}

function getBookings(user, filters) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEETS.BOOKINGS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [], count: 0, total: 0, totalPages: 0 };

  const raw  = sheet.getRange(1, 1, lastRow, 20).getValues();
  const PAGE = parseInt((filters && filters.pageSize) || 200);
  const page = parseInt((filters && filters.page)     || 0);
  const sFlt     = filters && filters.status   ? String(filters.status) : '';
  const dateFrom = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
  const dateTo   = filters && filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  const matchIdx = [];
  for (let i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    if (user.role === "Sales" && String(raw[i][15]).toLowerCase() !== user.name.toLowerCase()) continue;
    if (sFlt && raw[i][16] !== sFlt) continue;
    if (dateFrom || dateTo) {
      const d = new Date(raw[i][1]);
      if (isNaN(d)) { matchIdx.push(i); continue; }
      if (dateFrom && d < dateFrom) continue;
      if (dateTo   && d > dateTo)   continue;
    }
    matchIdx.push(i);
  }

  const total      = matchIdx.length;
  const totalPages = total > 0 ? Math.ceil(total / PAGE) : 0;

  matchIdx.reverse();
  const pageIdx = matchIdx.slice(page * PAGE, (page + 1) * PAGE);

  const bookings = pageIdx.map(function(i) {
    const r = raw[i];
    return {
      id:            String(r[0]).trim(),
      bookingDate:   String(r[1]).trim(),
      clientName:    String(r[2]).trim(),
      phone:         String(r[3]).trim(),
      email:         String(r[4]).trim(),
      propertyName:  String(r[5]).trim(),
      propertyType:  String(r[6]).trim(),
      checkIn:       String(r[7]).trim(),
      checkOut:      String(r[8]).trim(),
      nights:        r[9],
      roomType:      String(r[10]).trim(),
      ratePerNight:  r[11],
      totalValue:    r[12],
      commission:    r[13],
      leadId:        String(r[14]).trim(),
      salesRep:      String(r[15]).trim(),
      status:        String(r[16]).trim(),
      paymentStatus: String(r[17]).trim(),
      notes:         String(r[18]).trim(),
      createdAt:     String(r[19]).trim(),
      rowIndex:      i + 1
    };
  });

  return { success: true, data: bookings, count: bookings.length, total, page, pageSize: PAGE, totalPages };
}

// ─── INCENTIVES ───────────────────────────────────────────────

function getIncentives(user, filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INCENTIVES);
  const data = sheet.getDataRange().getValues();

  const records = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const row = data[i];

    if (user.role === "Sales" && String(row[2]).toLowerCase() !== user.email.toLowerCase()) continue;

    if (filters && filters.year && String(row[4]) !== String(filters.year)) continue;

    records.push({
      id:              String(row[0]).trim(),
      salesRep:        String(row[1]).trim(),
      email:           String(row[2]).trim(),
      month:           String(row[3]).trim(),
      year:            row[4],
      totalSales:      row[5],
      baseThreshold:   row[6],
      eligibleAmount:  row[7],
      incentiveRate:   row[8],
      incentiveAmount: row[9],
      status:          String(row[10]).trim(),
      paidDate:        String(row[11]).trim(),
      calculatedAt:    String(row[12]).trim()
    });
  }

  return { success: true, data: records, count: records.length };
}

/**
 * Multi-tier incentive formula sourced from flexibleIncentives.gs:
 *   Tier 0: 0 – 500k   @ 0%
 *   Tier 1: 500k – 1M  @ 1%
 *   Tier 2: 1M+         @ 1.5%
 * Tiers are configurable via Admin › Incentive Tiers.
 */
function calculateIncentive(data, user) {
  requireHROrAdmin(user);

  const { month, year } = data;
  if (!month || !year) return { success: false, error: "Month and year required" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingsSheet = ss.getSheetByName(SHEETS.BOOKINGS);
  const incentivesSheet = ss.getSheetByName(SHEETS.INCENTIVES);
  const bookingsData = bookingsSheet.getDataRange().getValues();

  // Get all confirmed bookings for the month
  const salesTotals = {};
  for (let i = 1; i < bookingsData.length; i++) {
    const row = bookingsData[i];
    if (!row[0]) continue;
    if (row[16] !== "Confirmed" && row[16] !== "Completed") continue;

    const bookingDate = new Date(row[1]);
    if (!bookingDate) continue;

    const bMonth = Utilities.formatDate(bookingDate, "UTC", "MMMM");
    const bYear = bookingDate.getFullYear();

    if (bMonth === month && String(bYear) === String(year)) {
      const rep = String(row[15]).trim();
      salesTotals[rep] = (salesTotals[rep] || 0) + (parseFloat(row[12]) || 0);
    }
  }

  // Fetch tiers from flexibleIncentives.gs (uses PropertiesService)
  const tiersResult = getIncentiveTiers();
  const tiers = tiersResult.success ? tiersResult.data : [
    { from: 0,       to: 500000,  rate: 0    },
    { from: 500000,  to: 1000000, rate: 0.01 },
    { from: 1000000, to: null,    rate: 0.015 }
  ];

  // Derive a representative rate label for display (first non-zero tier rate)
  const displayRate = (tiers.find(function(t) { return t.rate > 0; }) || { rate: 0.01 }).rate;

  let saved = 0;

  Object.keys(salesTotals).forEach(function(rep) {
    const total = salesTotals[rep];
    // Use tiered calculation from flexibleIncentives.gs
    const incentive = calculateTieredIncentive(total);
    const eligible  = Math.max(0, total - (tiers[0] ? tiers[0].to || 0 : 500000));

    const repUser  = getUserByName(rep);
    const repEmail = repUser ? repUser.email : "";

    const id = "INC-" + month.substring(0,3).toUpperCase() + year + "-" + rep.replace(/\s/g,"");

    incentivesSheet.appendRow([
      id, rep, repEmail, month, parseInt(year),
      total, tiers[0] ? (tiers[0].to || 500000) : 500000, eligible, displayRate, incentive,
      incentive > 0 ? "Pending Payment" : "Not Eligible",
      "",
      Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
    ]);
    saved++;
  });

  logActivity(user.email, "CALC_INCENTIVE",
    "Calculated tiered incentives for " + month + " " + year + " (" + tiers.length + " tiers)");
  return { success: true, message: saved + " incentive records calculated for " + month + " " + year + " (multi-tier)" };
}

function getUserByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === name.toLowerCase()) {
      return { name: data[i][0], email: data[i][1], role: data[i][2] };
    }
  }
  return null;
}

// ─── DASHBOARD STATS ──────────────────────────────────────────

function getDashboardStats(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const thisMonth = now.getMonth() + 1;
  const thisYear = now.getFullYear();

  // Bookings stats
  const bookingsData = ss.getSheetByName(SHEETS.BOOKINGS).getDataRange().getValues();
  let totalRevenue = 0, monthlyRevenue = 0, confirmedBookings = 0;

  for (let i = 1; i < bookingsData.length; i++) {
    const row = bookingsData[i];
    if (!row[0]) continue;
    if (user.role === "Sales" && String(row[15]).toLowerCase() !== user.name.toLowerCase()) continue;

    const val = parseFloat(row[12]) || 0;
    totalRevenue += val;

    const bookDate = new Date(row[1]);
    if (bookDate.getMonth() + 1 === thisMonth && bookDate.getFullYear() === thisYear) {
      monthlyRevenue += val;
    }
    if (row[16] === "Confirmed" || row[16] === "Completed") confirmedBookings++;
  }

  // Leads stats
  const leadsData = ss.getSheetByName(SHEETS.LEADS).getDataRange().getValues();
  let totalLeads = 0, wonLeads = 0, pendingLeads = 0;

  for (let i = 1; i < leadsData.length; i++) {
    const row = leadsData[i];
    if (!row[0]) continue;
    if (user.role === "Sales" && String(row[9]).toLowerCase() !== user.name.toLowerCase()) continue;
    totalLeads++;
    if (row[10] === "Won") wonLeads++;
    if (row[10] === "Qualified" || row[10] === "Proposal Sent" || row[10] === "New") pendingLeads++;
  }

  // DSR count
  const dsrData = ss.getSheetByName(SHEETS.DAILY_REPORTS).getDataRange().getValues();
  let totalDSR = 0, monthlyDSR = 0;
  for (let i = 1; i < dsrData.length; i++) {
    const row = dsrData[i];
    if (!row[0]) continue;
    if (user.role === "Sales" && String(row[3]).toLowerCase() !== user.email.toLowerCase()) continue;
    totalDSR++;
    const d = new Date(row[1]);
    if (d.getMonth() + 1 === thisMonth && d.getFullYear() === thisYear) monthlyDSR++;
  }

  // Travel pending
  const travelData = ss.getSheetByName(SHEETS.TRAVEL_PLANS).getDataRange().getValues();
  let pendingTravel = 0;
  for (let i = 1; i < travelData.length; i++) {
    if (travelData[i][12] === "Pending") pendingTravel++;
  }

  // Incentives this year
  const incentivesData = ss.getSheetByName(SHEETS.INCENTIVES).getDataRange().getValues();
  let totalIncentives = 0;
  for (let i = 1; i < incentivesData.length; i++) {
    const row = incentivesData[i];
    if (!row[0]) continue;
    if (user.role === "Sales" && String(row[2]).toLowerCase() !== user.email.toLowerCase()) continue;
    if (String(row[4]) === String(thisYear)) {
      totalIncentives += parseFloat(row[9]) || 0;
    }
  }

  const conversionRate = totalLeads > 0 ? Math.round((wonLeads / totalLeads) * 100) : 0;

  return {
    success: true,
    data: {
      totalRevenue,
      monthlyRevenue,
      confirmedBookings,
      totalLeads,
      wonLeads,
      pendingLeads,
      conversionRate,
      totalDSR,
      monthlyDSR,
      pendingTravel,
      totalIncentives,
      thisMonth,
      thisYear
    }
  };
}

// ─── MONTHLY REPORT ───────────────────────────────────────────

function getMonthlyReport(data, user) {
  requireHROrAdmin(user);

  const { month, year } = data;
  if (!month || !year) return { success: false, error: "Month and year required" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Bookings by sales rep
  const bookings = ss.getSheetByName(SHEETS.BOOKINGS).getDataRange().getValues();
  const salesPerformance = {};

  for (let i = 1; i < bookings.length; i++) {
    const row = bookings[i];
    if (!row[0]) continue;

    const d = new Date(row[1]);
    const bMonth = d.getMonth() + 1;
    const bYear = d.getFullYear();

    if (bMonth !== parseInt(month) || bYear !== parseInt(year)) continue;

    const rep = String(row[15]).trim();
    if (!salesPerformance[rep]) {
      salesPerformance[rep] = { rep, bookings: 0, revenue: 0, commission: 0 };
    }
    salesPerformance[rep].bookings++;
    salesPerformance[rep].revenue += parseFloat(row[12]) || 0;
    salesPerformance[rep].commission += parseFloat(row[13]) || 0;
  }

  // Leads this month
  const leads = ss.getSheetByName(SHEETS.LEADS).getDataRange().getValues();
  const leadsData = {};

  for (let i = 1; i < leads.length; i++) {
    const row = leads[i];
    if (!row[0]) continue;

    const d = new Date(row[1]);
    if (d.getMonth() + 1 !== parseInt(month) || d.getFullYear() !== parseInt(year)) continue;

    const rep = String(row[9]).trim();
    if (!leadsData[rep]) leadsData[rep] = { new: 0, won: 0, lost: 0 };
    leadsData[rep].new++;
    if (row[10] === "Won") leadsData[rep].won++;
    if (row[10] === "Lost") leadsData[rep].lost++;
  }

  // DSR this month
  const dsr = ss.getSheetByName(SHEETS.DAILY_REPORTS).getDataRange().getValues();
  const dsrData = {};

  for (let i = 1; i < dsr.length; i++) {
    const row = dsr[i];
    if (!row[0]) continue;

    const d = new Date(row[1]);
    if (d.getMonth() + 1 !== parseInt(month) || d.getFullYear() !== parseInt(year)) continue;

    const rep = String(row[2]).trim();
    if (!dsrData[rep]) dsrData[rep] = 0;
    dsrData[rep]++;
  }

  // Combine
  const allReps = new Set([
    ...Object.keys(salesPerformance),
    ...Object.keys(leadsData),
    ...Object.keys(dsrData)
  ]);

  const report = Array.from(allReps).map(rep => ({
    rep,
    revenue:         salesPerformance[rep]?.revenue || 0,
    bookings:        salesPerformance[rep]?.bookings || 0,
    commission:      salesPerformance[rep]?.commission || 0,
    newLeads:        leadsData[rep]?.new || 0,
    wonLeads:        leadsData[rep]?.won || 0,
    lostLeads:       leadsData[rep]?.lost || 0,
    conversionRate:  leadsData[rep] ?
                     Math.round((leadsData[rep].won / leadsData[rep].new) * 100) : 0,
    dsrCount:        dsrData[rep] || 0
  }));

  report.sort((a, b) => b.revenue - a.revenue);

  return { success: true, data: report, month, year };
}

// ─── EXPORT ───────────────────────────────────────────────────

function exportReport(data, user) {
  requireHROrAdmin(user);

  const { reportType, format, month, year } = data;
  let csvContent = "";
  let filename = reportType + "_" + (month || "") + "_" + (year || "") + "." + (format || "csv");

  if (reportType === "monthly") {
    const result = getMonthlyReport({ month, year }, user);
    if (!result.success) return result;

    csvContent = "Sales Rep,Revenue (₹),Bookings,Commission (₹),New Leads,Won Leads,Conversion %,DSR Count\n";
    result.data.forEach(r => {
      csvContent += `"${r.rep}",${r.revenue},${r.bookings},${r.commission},${r.newLeads},${r.wonLeads},${r.conversionRate}%,${r.dsrCount}\n`;
    });
  } else if (reportType === "leads") {
    const result = getLeads(user, { month, year });
    if (!result.success) return result;

    csvContent = "ID,Date,Source,Client,Contact,Phone,Property Type,Budget,Assigned To,Status,Value\n";
    result.data.forEach(r => {
      csvContent += `"${r.id}","${r.date}","${r.leadSource}","${r.clientName}","${r.contactPerson}","${r.phone}","${r.propertyType}",${r.budget},"${r.assignedTo}","${r.status}",${r.value}\n`;
    });
  } else if (reportType === "bookings") {
    const result = getBookings(user, {});
    if (!result.success) return result;

    csvContent = "ID,Date,Client,Property,Check-In,Check-Out,Nights,Total Value,Commission,Sales Rep,Status\n";
    result.data.forEach(r => {
      csvContent += `"${r.id}","${r.bookingDate}","${r.clientName}","${r.propertyName}","${r.checkIn}","${r.checkOut}",${r.nights},${r.totalValue},${r.commission},"${r.salesRep}","${r.status}"\n`;
    });
  }

  // Save to Drive and get link
  const folder = getSubFolder("Reports");
  const file = folder.createFile(filename, csvContent, MimeType.CSV);

  logActivity(user.email, "EXPORT_REPORT", "Exported " + reportType + " report");
  return {
    success: true,
    downloadUrl: file.getDownloadUrl(),
    filename: filename,
    message: "Report exported successfully"
  };
}

// ─── HR EMAIL LIST ────────────────────────────────────────────

function getHREmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.EMAIL_RECIPIENTS);
  const data = sheet.getDataRange().getValues();

  const emails = [];
  for (let i = 1; i < data.length; i++) {
    if (
      data[i][0] &&
      ["HR", "Admin"].includes(String(data[i][0]).trim()) &&
      String(data[i][3]).trim().toLowerCase() === "yes" &&
      isValidEmail(String(data[i][2]).trim())
    ) {
      emails.push(String(data[i][2]).trim());
    }
  }
  return emails;
}

