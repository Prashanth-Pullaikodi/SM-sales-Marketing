/**
 * ============================================================
 * Proposals.gs - Proposal & Calculator Engine
 * ============================================================
 * Handles: sheet setup, CRUD, discount validation, approval
 * workflow, and PDF generation for customer + internal views.
 *
 * Access roles:
 *   Admin  – full access (create, edit, approve, delete, COGS config)
 *   Exec   – create, view all, approve/reject
 *   Sales  – create own, view own, submit
 *   HR     – read-only view
 * ============================================================
 */

// ─── SHEET NAMES ─────────────────────────────────────────────
var PROPOSAL_SHEETS = {
  PROPOSALS: "Proposals",
  COGS:      "ProposalCOGS"
};

// ─── DISCOUNT TIERS ──────────────────────────────────────────
var DISCOUNT_TIERS = {
  AUTO:   { max: 5,   label: "Standard",  approvalRequired: false,  approvalRole: null },
  MID:    { max: 15,  label: "Manager",   approvalRequired: true,   approvalRole: ["Admin","Exec"] },
  HIGH:   { max: 100, label: "Executive", approvalRequired: true,   approvalRole: ["Admin"] }
};

// ─── PROPOSAL STATUS ─────────────────────────────────────────
var PROPOSAL_STATUS = {
  DRAFT:    "Draft",
  PENDING:  "Pending Approval",
  APPROVED: "Approved",
  REJECTED: "Rejected",
  SENT:     "Sent to Client",
  ACCEPTED: "Client Accepted",
  DECLINED: "Client Declined"
};

// ─── INIT SHEETS ─────────────────────────────────────────────

function initProposalSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Proposals sheet ──────────────────────────────────────
  var propHeaders = [
    "ID", "ProposalDate", "ClientName", "ClientEmail", "ClientPhone",
    "PropertyName", "RoomType", "CheckIn", "CheckOut", "Nights",
    "BaseRatePerNight", "DiscountPct", "DiscountReason",
    "FinalRatePerNight", "TotalValue",
    "COGSPerNight", "TotalCOGS", "GrossMargin", "MarginPct",
    "Status", "ApprovalLevel", "ApprovedBy", "ApprovedAt",
    "Notes", "CreatedBy", "CreatedEmail", "CreatedAt"
  ];
  var propSheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  if (!propSheet) {
    propSheet = ss.insertSheet(PROPOSAL_SHEETS.PROPOSALS);
    Logger.log("Created sheet: " + PROPOSAL_SHEETS.PROPOSALS);
  }
  _setupProposalSheetHeaders(propSheet, propHeaders, "#1565c0");

  // ── COGS config sheet ────────────────────────────────────
  var cogsHeaders = [
    "RoomType", "COGSPerNight", "MinMarginPct", "BaseRate", "Description", "LastUpdated"
  ];
  var cogsSheet = ss.getSheetByName(PROPOSAL_SHEETS.COGS);
  if (!cogsSheet) {
    cogsSheet = ss.insertSheet(PROPOSAL_SHEETS.COGS);
    Logger.log("Created sheet: " + PROPOSAL_SHEETS.COGS);
    // Seed default COGS rates
    var defaultCOGS = [
      ["Deluxe Room",      2800, 35, 4500, "Standard deluxe room package",       new Date()],
      ["Superior Room",    2200, 35, 3500, "Superior room package",               new Date()],
      ["Executive Suite",  4500, 38, 7500, "Executive suite with lounge access",  new Date()],
      ["Presidential Suite",8000,40,14000, "Presidential suite full package",     new Date()],
      ["Family Room",      3200, 35, 5200, "Family room – 2 adults + 2 children", new Date()],
      ["Villa",            6500, 40,11000, "Private villa",                       new Date()]
    ];
    cogsSheet.getRange(2, 1, defaultCOGS.length, 6).setValues(defaultCOGS);
  }
  _setupProposalSheetHeaders(cogsSheet, cogsHeaders, "#1565c0");
}

function _setupProposalSheetHeaders(sheet, headers, color) {
  var header = sheet.getRange(1, 1, 1, headers.length);
  header.setValues([headers]);
  header.setBackground(color);
  header.setFontColor("#FFFFFF");
  header.setFontWeight("bold");
  header.setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 140);
}

// ─── PROPOSAL ID GENERATOR ────────────────────────────────────

function _generateProposalId() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var last  = sheet ? sheet.getLastRow() : 1;
  var num   = String(last).padStart(4, "0");
  return "PROP-" + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd") + "-" + num;
}

// ─── COGS LOOKUP ─────────────────────────────────────────────

function _getCOGSForRoomType(roomType) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.COGS);
  if (!sheet) return { cogsPerNight: 0, minMarginPct: 35, baseRate: 0 };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === roomType.trim().toLowerCase()) {
      return {
        cogsPerNight: Number(data[i][1]) || 0,
        minMarginPct: Number(data[i][2]) || 35,
        baseRate:     Number(data[i][3]) || 0
      };
    }
  }
  return { cogsPerNight: 0, minMarginPct: 35, baseRate: 0 };
}

// ─── DISCOUNT TIER ENGINE ─────────────────────────────────────

function getDiscountTier(discountPct) {
  var d = parseFloat(discountPct) || 0;
  if (d <= 0)                            return { tier: "none",      label: "No Discount",  requiresApproval: false, approvalRole: null };
  if (d <= DISCOUNT_TIERS.AUTO.max)      return { tier: "standard",  label: "Standard",     requiresApproval: false, approvalRole: null };
  if (d <= DISCOUNT_TIERS.MID.max)       return { tier: "manager",   label: "Manager Approval Required",   requiresApproval: true, approvalRole: ["Admin","Exec"] };
  return                                        { tier: "executive", label: "Executive Approval Required", requiresApproval: true, approvalRole: ["Admin"] };
}

// ─── CALCULATE PROPOSAL NUMBERS ──────────────────────────────

function calculateProposalNumbers(data) {
  var baseRate    = parseFloat(data.baseRatePerNight)  || 0;
  var discountPct = parseFloat(data.discountPct)       || 0;
  var nights      = parseInt(data.nights)              || 1;
  var roomType    = data.roomType                      || "";

  var cogs        = _getCOGSForRoomType(roomType);
  var finalRate   = baseRate * (1 - discountPct / 100);
  var totalValue  = finalRate * nights;
  var totalCOGS   = cogs.cogsPerNight * nights;
  var grossMargin = totalValue - totalCOGS;
  var marginPct   = totalValue > 0 ? (grossMargin / totalValue * 100) : 0;
  var tier        = getDiscountTier(discountPct);

  return {
    success:         true,
    baseRate:        baseRate,
    discountPct:     discountPct,
    finalRate:       Math.round(finalRate * 100) / 100,
    nights:          nights,
    totalValue:      Math.round(totalValue * 100) / 100,
    cogsPerNight:    cogs.cogsPerNight,
    totalCOGS:       totalCOGS,
    grossMargin:     Math.round(grossMargin * 100) / 100,
    marginPct:       Math.round(marginPct * 100) / 100,
    minMarginPct:    cogs.minMarginPct,
    marginWarning:   marginPct < cogs.minMarginPct,
    discountTier:    tier,
    cogsConfig:      cogs
  };
}

// ─── GET PROPOSALS ────────────────────────────────────────────

function getProposals(user, data) {
  if (!hasAccess(user, ["Admin", "Exec", "HR", "Sales"])) {
    return { success: false, error: "Access denied" };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, data: [] };
  }

  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var result  = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;

    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    obj.rowIndex = i + 1;

    // Sales can only see own proposals
    if (user.role === "Sales" && obj.CreatedEmail !== user.email) continue;

    // Status filter
    if (data && data.status && data.status !== "all" && obj.Status !== data.status) continue;

    // Format dates
    if (obj.CheckIn  instanceof Date) obj.CheckIn  = Utilities.formatDate(obj.CheckIn,  "Asia/Kolkata", "yyyy-MM-dd");
    if (obj.CheckOut instanceof Date) obj.CheckOut = Utilities.formatDate(obj.CheckOut, "Asia/Kolkata", "yyyy-MM-dd");
    if (obj.ProposalDate instanceof Date) obj.ProposalDate = Utilities.formatDate(obj.ProposalDate, "Asia/Kolkata", "yyyy-MM-dd");
    if (obj.CreatedAt   instanceof Date) obj.CreatedAt    = Utilities.formatDate(obj.CreatedAt,    "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    if (obj.ApprovedAt  instanceof Date) obj.ApprovedAt   = Utilities.formatDate(obj.ApprovedAt,   "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    result.push(obj);
  }

  // Most recent first
  result.sort(function(a, b) { return (b.CreatedAt || "").localeCompare(a.CreatedAt || ""); });

  return { success: true, data: result, total: result.length };
}

// ─── CREATE PROPOSAL ─────────────────────────────────────────

function createProposal(data, user) {
  if (!hasAccess(user, ["Admin", "Exec", "Sales"])) {
    return { success: false, error: "Access denied" };
  }
  _validateProposalData(data);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  if (!sheet) initProposalSheets();
  sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);

  var calc  = calculateProposalNumbers(data);
  var tier  = calc.discountTier;
  var id    = _generateProposalId();
  var now   = new Date();
  var status = tier.requiresApproval ? PROPOSAL_STATUS.PENDING : PROPOSAL_STATUS.DRAFT;

  var row = [
    id,
    Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd"),
    data.clientName    || "",
    data.clientEmail   || "",
    data.clientPhone   || "",
    data.propertyName  || "SandalMist Resort",
    data.roomType      || "",
    data.checkIn       || "",
    data.checkOut      || "",
    parseInt(data.nights) || 1,
    calc.baseRate,
    calc.discountPct,
    data.discountReason || "",
    calc.finalRate,
    calc.totalValue,
    calc.cogsPerNight,
    calc.totalCOGS,
    calc.grossMargin,
    calc.marginPct,
    status,
    tier.tier,
    "",   // ApprovedBy
    "",   // ApprovedAt
    data.notes || "",
    user.name,
    user.email,
    Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss")
  ];

  sheet.appendRow(row);
  logActivity(user.email, "PROPOSAL_CREATE", "Created proposal " + id);

  return { success: true, id: id, status: status, calc: calc, message: "Proposal created successfully" };
}

// ─── UPDATE PROPOSAL ─────────────────────────────────────────

function updateProposal(data, user) {
  if (!hasAccess(user, ["Admin", "Exec", "Sales"])) {
    return { success: false, error: "Access denied" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 27).getValues()[0];

  // Sales can only edit own proposals in Draft/Pending
  if (user.role === "Sales") {
    if (row[25] !== user.email) return { success: false, error: "Cannot edit another user's proposal" };
    if (row[19] !== PROPOSAL_STATUS.DRAFT && row[19] !== PROPOSAL_STATUS.PENDING) {
      return { success: false, error: "Can only edit Draft or Pending proposals" };
    }
  }

  var calc  = calculateProposalNumbers(data);
  var tier  = calc.discountTier;
  var status = tier.requiresApproval ? PROPOSAL_STATUS.PENDING : (row[19] === PROPOSAL_STATUS.DRAFT ? PROPOSAL_STATUS.DRAFT : row[19]);

  // Update columns
  sheet.getRange(data.rowIndex, 3).setValue(data.clientName    || row[2]);
  sheet.getRange(data.rowIndex, 4).setValue(data.clientEmail   || row[3]);
  sheet.getRange(data.rowIndex, 5).setValue(data.clientPhone   || row[4]);
  sheet.getRange(data.rowIndex, 6).setValue(data.propertyName  || row[5]);
  sheet.getRange(data.rowIndex, 7).setValue(data.roomType      || row[6]);
  sheet.getRange(data.rowIndex, 8).setValue(data.checkIn       || row[7]);
  sheet.getRange(data.rowIndex, 9).setValue(data.checkOut      || row[8]);
  sheet.getRange(data.rowIndex, 10).setValue(parseInt(data.nights) || row[9]);
  sheet.getRange(data.rowIndex, 11).setValue(calc.baseRate);
  sheet.getRange(data.rowIndex, 12).setValue(calc.discountPct);
  sheet.getRange(data.rowIndex, 13).setValue(data.discountReason || row[12]);
  sheet.getRange(data.rowIndex, 14).setValue(calc.finalRate);
  sheet.getRange(data.rowIndex, 15).setValue(calc.totalValue);
  sheet.getRange(data.rowIndex, 16).setValue(calc.cogsPerNight);
  sheet.getRange(data.rowIndex, 17).setValue(calc.totalCOGS);
  sheet.getRange(data.rowIndex, 18).setValue(calc.grossMargin);
  sheet.getRange(data.rowIndex, 19).setValue(calc.marginPct);
  sheet.getRange(data.rowIndex, 20).setValue(status);
  sheet.getRange(data.rowIndex, 21).setValue(tier.tier);
  sheet.getRange(data.rowIndex, 24).setValue(data.notes || row[23]);

  logActivity(user.email, "PROPOSAL_UPDATE", "Updated proposal " + row[0]);
  return { success: true, calc: calc, status: status, message: "Proposal updated" };
}

// ─── APPROVE / REJECT ─────────────────────────────────────────

function approveProposal(data, user) {
  if (!hasAccess(user, ["Admin", "Exec"])) {
    return { success: false, error: "Admin or Exec role required to approve" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 27).getValues()[0];
  var tier  = row[20]; // ApprovalLevel

  // High-discount (executive) proposals need Admin
  if (tier === "executive" && user.role !== "Admin") {
    return { success: false, error: "Executive-level discounts require Admin approval" };
  }

  var now = new Date();
  sheet.getRange(data.rowIndex, 20).setValue(PROPOSAL_STATUS.APPROVED);
  sheet.getRange(data.rowIndex, 22).setValue(user.name);
  sheet.getRange(data.rowIndex, 23).setValue(Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss"));
  if (data.notes) sheet.getRange(data.rowIndex, 24).setValue(data.notes);

  logActivity(user.email, "PROPOSAL_APPROVE", "Approved proposal " + row[0]);
  return { success: true, message: "Proposal approved by " + user.name };
}

function rejectProposal(data, user) {
  if (!hasAccess(user, ["Admin", "Exec"])) {
    return { success: false, error: "Admin or Exec role required" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 1).getValues()[0];

  sheet.getRange(data.rowIndex, 20).setValue(PROPOSAL_STATUS.REJECTED);
  sheet.getRange(data.rowIndex, 22).setValue(user.name);
  sheet.getRange(data.rowIndex, 23).setValue(Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss"));
  if (data.notes) sheet.getRange(data.rowIndex, 24).setValue(data.notes);

  logActivity(user.email, "PROPOSAL_REJECT", "Rejected proposal " + row[0]);
  return { success: true, message: "Proposal rejected" };
}

// ─── SEND TO CLIENT ──────────────────────────────────────────

function sendProposalToClient(data, user) {
  if (!hasAccess(user, ["Admin", "Exec", "Sales"])) {
    return { success: false, error: "Access denied" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 27).getValues()[0];

  if (row[19] === PROPOSAL_STATUS.PENDING) {
    return { success: false, error: "Proposal is pending approval. Cannot send yet." };
  }

  sheet.getRange(data.rowIndex, 20).setValue(PROPOSAL_STATUS.SENT);
  logActivity(user.email, "PROPOSAL_SENT", "Sent proposal " + row[0] + " to client " + row[2]);
  return { success: true, message: "Proposal marked as Sent to Client" };
}

// ─── DELETE PROPOSAL ─────────────────────────────────────────

function deleteProposal(data, user) {
  if (!hasAccess(user, ["Admin"])) {
    return { success: false, error: "Admin role required to delete proposals" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var id    = sheet.getRange(data.rowIndex, 1).getValue();
  sheet.deleteRow(data.rowIndex);

  logActivity(user.email, "PROPOSAL_DELETE", "Deleted proposal " + id);
  return { success: true, message: "Proposal deleted" };
}

// ─── GET / UPDATE COGS RATES ──────────────────────────────────

function getCOGSRates(user) {
  if (!hasAccess(user, ["Admin", "Exec", "HR", "Sales"])) {
    return { success: false, error: "Access denied" };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.COGS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var result  = [];

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = rows[i][idx]; });
    obj.rowIndex = i + 1;
    if (obj.LastUpdated instanceof Date) {
      obj.LastUpdated = Utilities.formatDate(obj.LastUpdated, "Asia/Kolkata", "yyyy-MM-dd");
    }
    result.push(obj);
  }
  return { success: true, data: result };
}

function updateCOGSRate(data, user) {
  if (!hasAccess(user, ["Admin"])) {
    return { success: false, error: "Admin role required to update COGS rates" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.COGS);

  sheet.getRange(data.rowIndex, 2).setValue(parseFloat(data.COGSPerNight)  || 0);
  sheet.getRange(data.rowIndex, 3).setValue(parseFloat(data.MinMarginPct)  || 35);
  sheet.getRange(data.rowIndex, 4).setValue(parseFloat(data.BaseRate)      || 0);
  sheet.getRange(data.rowIndex, 5).setValue(data.Description || "");
  sheet.getRange(data.rowIndex, 6).setValue(new Date());

  logActivity(user.email, "COGS_UPDATE", "Updated COGS for " + data.RoomType);
  return { success: true, message: "COGS rate updated" };
}

// ─── GENERATE CUSTOMER PDF ────────────────────────────────────

function generateCustomerProposalPDF(data, user) {
  if (!hasAccess(user, ["Admin", "Exec", "Sales"])) {
    return { success: false, error: "Access denied" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 27).getValues()[0];

  var proposal = {
    id:             row[0],
    proposalDate:   row[1] instanceof Date ? Utilities.formatDate(row[1], "Asia/Kolkata", "dd MMM yyyy") : row[1],
    clientName:     row[2],
    clientEmail:    row[3],
    clientPhone:    row[4],
    propertyName:   row[5],
    roomType:       row[6],
    checkIn:        row[7] instanceof Date ? Utilities.formatDate(row[7], "Asia/Kolkata", "dd MMM yyyy") : row[7],
    checkOut:       row[8] instanceof Date ? Utilities.formatDate(row[8], "Asia/Kolkata", "dd MMM yyyy") : row[8],
    nights:         row[9],
    baseRate:       row[10],
    discountPct:    row[11],
    discountReason: row[12],
    finalRate:      row[13],
    totalValue:     row[14],
    status:         row[19],
    notes:          row[23],
    createdBy:      row[24]
  };

  var tmpl = HtmlService.createTemplateFromFile("CustomerProposalTemplate");
  tmpl.proposal = proposal;
  var html = tmpl.evaluate().getContent();

  var blob   = Utilities.newBlob(html, "text/html", proposal.id + "_Customer.html");
  var folder = getSubFolder("Proposals");
  var file   = folder.createFile(blob.setName(proposal.id + "_Customer_Proposal.html"));

  logActivity(user.email, "PROPOSAL_PDF_CUSTOMER", "Generated customer PDF for " + proposal.id);
  return {
    success:     true,
    downloadUrl: file.getDownloadUrl(),
    fileId:      file.getId(),
    message:     "Customer proposal generated"
  };
}

// ─── GENERATE INTERNAL PDF ────────────────────────────────────

function generateInternalProposalPDF(data, user) {
  if (!hasAccess(user, ["Admin", "Exec"])) {
    return { success: false, error: "Admin or Exec role required for internal PDF" };
  }
  if (!data.rowIndex) return { success: false, error: "Row index required" };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROPOSAL_SHEETS.PROPOSALS);
  var row   = sheet.getRange(data.rowIndex, 1, 1, 27).getValues()[0];

  var proposal = {
    id:             row[0],
    proposalDate:   row[1] instanceof Date ? Utilities.formatDate(row[1], "Asia/Kolkata", "dd MMM yyyy") : row[1],
    clientName:     row[2],
    clientEmail:    row[3],
    clientPhone:    row[4],
    propertyName:   row[5],
    roomType:       row[6],
    checkIn:        row[7] instanceof Date ? Utilities.formatDate(row[7], "Asia/Kolkata", "dd MMM yyyy") : row[7],
    checkOut:       row[8] instanceof Date ? Utilities.formatDate(row[8], "Asia/Kolkata", "dd MMM yyyy") : row[8],
    nights:         row[9],
    baseRate:       row[10],
    discountPct:    row[11],
    discountReason: row[12],
    finalRate:      row[13],
    totalValue:     row[14],
    cogsPerNight:   row[15],
    totalCOGS:      row[16],
    grossMargin:    row[17],
    marginPct:      row[18],
    status:         row[19],
    approvalLevel:  row[20],
    approvedBy:     row[21],
    approvedAt:     row[22] instanceof Date ? Utilities.formatDate(row[22], "Asia/Kolkata", "dd MMM yyyy HH:mm") : row[22],
    notes:          row[23],
    createdBy:      row[24],
    createdEmail:   row[25],
    createdAt:      row[26] instanceof Date ? Utilities.formatDate(row[26], "Asia/Kolkata", "dd MMM yyyy HH:mm") : row[26]
  };

  var tmpl = HtmlService.createTemplateFromFile("InternalProposalTemplate");
  tmpl.proposal = proposal;
  var html = tmpl.evaluate().getContent();

  var blob   = Utilities.newBlob(html, "text/html", proposal.id + "_Internal.html");
  var folder = getSubFolder("Proposals");
  var file   = folder.createFile(blob.setName(proposal.id + "_Internal_Breakdown.html"));

  logActivity(user.email, "PROPOSAL_PDF_INTERNAL", "Generated internal PDF for " + proposal.id);
  return {
    success:     true,
    downloadUrl: file.getDownloadUrl(),
    fileId:      file.getId(),
    message:     "Internal breakdown generated"
  };
}

// ─── VALIDATION ──────────────────────────────────────────────

function _validateProposalData(data) {
  if (!data.clientName)  throw new Error("Client name is required");
  if (!data.roomType)    throw new Error("Room type is required");
  if (!data.checkIn)     throw new Error("Check-in date is required");
  if (!data.checkOut)    throw new Error("Check-out date is required");
  if (!data.nights || parseInt(data.nights) < 1) throw new Error("Nights must be at least 1");
  if (!data.baseRatePerNight || parseFloat(data.baseRatePerNight) <= 0) throw new Error("Base rate must be greater than 0");
  var disc = parseFloat(data.discountPct) || 0;
  if (disc < 0 || disc > 100) throw new Error("Discount must be between 0 and 100");
}
