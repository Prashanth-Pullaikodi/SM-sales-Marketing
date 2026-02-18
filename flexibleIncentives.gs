/**
 * ============================================================
 * FlexibleIncentive.gs - Advanced Incentive Configuration
 * Supports multi-tier commission structures
 * ============================================================
 */

// ─── INCENTIVE TIER STRUCTURE ─────────────────────────────────
/**
 * Example tier structure:
 * [
 *   { from: 0,      to: 500000,  rate: 0    }, // First 500k: 0%
 *   { from: 500000, to: 1000000, rate: 0.01 }, // Next 500k: 1%
 *   { from: 1000000, to: null,   rate: 0.015}  // Above 1M: 1.5%
 * ]
 */

function getIncentiveTiers() {
  var props = PropertiesService.getScriptProperties();
  var tiersJson = props.getProperty('incentive_tiers');
  
  if (!tiersJson) {
    // Default structure
    var defaultTiers = [
      { from: 0,      to: 500000,  rate: 0,     label: 'Base (0%)' },
      { from: 500000, to: 1000000, rate: 0.01,  label: 'Tier 1 (1%)' },
      { from: 1000000, to: null,   rate: 0.015, label: 'Tier 2 (1.5%)' }
    ];
    props.setProperty('incentive_tiers', JSON.stringify(defaultTiers));
    return { success: true, data: defaultTiers };
  }
  
  try {
    var tiers = JSON.parse(tiersJson);
    return { success: true, data: tiers };
  } catch (e) {
    return { success: false, error: 'Invalid tier configuration' };
  }
}

function updateIncentiveTiers(data, user) {
  requireAdmin(user);
  
  var tiers = data.tiers;
  if (!tiers || !Array.isArray(tiers)) {
    return { success: false, error: 'Invalid tiers array' };
  }
  
  // Validate tiers
  for (var i = 0; i < tiers.length; i++) {
    var tier = tiers[i];
    if (typeof tier.from !== 'number' || tier.from < 0) {
      return { success: false, error: 'Invalid tier ' + i + ': from must be >= 0' };
    }
    if (tier.to !== null && (typeof tier.to !== 'number' || tier.to <= tier.from)) {
      return { success: false, error: 'Invalid tier ' + i + ': to must be > from or null' };
    }
    if (typeof tier.rate !== 'number' || tier.rate < 0 || tier.rate > 1) {
      return { success: false, error: 'Invalid tier ' + i + ': rate must be 0-1' };
    }
  }
  
  // Sort tiers by 'from' value
  tiers.sort(function(a, b) { return a.from - b.from; });
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('incentive_tiers', JSON.stringify(tiers));
  props.setProperty('incentive_tiers_updated', new Date().toISOString());
  props.setProperty('incentive_tiers_updated_by', user.email);
  
  logActivity(user.email, 'UPDATE_INCENTIVE_TIERS', tiers.length + ' tiers configured');
  
  return { success: true, message: 'Incentive tiers updated successfully' };
}

// ─── CALCULATE INCENTIVE WITH TIERS ───────────────────────────

function calculateTieredIncentive(totalSales) {
  var tiersResult = getIncentiveTiers();
  if (!tiersResult.success) return 0;
  
  var tiers = tiersResult.data;
  var incentive = 0;
  
  for (var i = 0; i < tiers.length; i++) {
    var tier = tiers[i];
    var from = tier.from;
    var to = tier.to || Infinity;
    var rate = tier.rate;
    
    if (totalSales <= from) break;
    
    var applicableAmount = Math.min(totalSales, to) - from;
    if (applicableAmount > 0) {
      incentive += applicableAmount * rate;
    }
  }
  
  return incentive;
}

// ─── RECALCULATE INCENTIVES WITH NEW TIERS ────────────────────

function recalculateIncentivesForMonth(data, user) {
  requireHROrAdmin(user);
  
  var month = data.month;
  var year = data.year;
  
  if (!month || !year) {
    return { success: false, error: 'Month and year required' };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName(SHEETS.BOOKINGS);
  var incentivesSheet = ss.getSheetByName(SHEETS.INCENTIVES);
  
  var bookingsData = bookingsSheet.getDataRange().getValues();
  
  // Calculate sales by rep
  var salesTotals = {};
  for (var i = 1; i < bookingsData.length; i++) {
    var row = bookingsData[i];
    if (!row[0]) continue;
    if (row[16] !== 'Confirmed' && row[16] !== 'Completed') continue;
    
    var bookingDate = new Date(row[1]);
    var bMonth = Utilities.formatDate(bookingDate, 'UTC', 'MMMM');
    var bYear = bookingDate.getFullYear();
    
    if (bMonth === month && String(bYear) === String(year)) {
      var rep = String(row[15]).trim();
      salesTotals[rep] = (salesTotals[rep] || 0) + (parseFloat(row[12]) || 0);
    }
  }
  
  // Get tier info
  var tiersResult = getIncentiveTiers();
  var tiers = tiersResult.success ? tiersResult.data : [];
  
  // Clear existing records for this month
  var incentivesData = incentivesSheet.getDataRange().getValues();
  for (var j = incentivesData.length - 1; j >= 1; j--) {
    if (String(incentivesData[j][3]) === month && String(incentivesData[j][4]) === String(year)) {
      incentivesSheet.deleteRow(j + 1);
    }
  }
  
  // Create new records
  var count = 0;
  Object.keys(salesTotals).forEach(function(rep) {
    var totalSales = salesTotals[rep];
    var incentive = calculateTieredIncentive(totalSales);
    
    var repUser = getUserByName(rep);
    var repEmail = repUser ? repUser.email : '';
    
    var id = 'INC-' + month.substring(0, 3).toUpperCase() + year + '-' + 
             rep.replace(/\s+/g, '').substring(0, 10);
    
    var tierBreakdown = getTierBreakdown(totalSales);
    
    incentivesSheet.appendRow([
      id,
      rep,
      repEmail,
      month,
      parseInt(year),
      totalSales,
      tierBreakdown.baseThreshold || 0,
      tierBreakdown.eligibleAmount || (totalSales > 500000 ? totalSales - 500000 : 0),
      tierBreakdown.effectiveRate || 0.01,
      incentive,
      incentive > 0 ? 'Pending Payment' : 'Not Eligible',
      '',
      Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm:ss')
    ]);
    count++;
  });
  
  logActivity(user.email, 'RECALC_INCENTIVES', month + ' ' + year + ': ' + count + ' records');
  
  return {
    success: true,
    message: count + ' incentive records calculated for ' + month + ' ' + year,
    tiers: tiers
  };
}

function getTierBreakdown(totalSales) {
  var tiersResult = getIncentiveTiers();
  if (!tiersResult.success) {
    return {
      baseThreshold: 500000,
      eligibleAmount: Math.max(0, totalSales - 500000),
      effectiveRate: 0.01
    };
  }
  
  var tiers = tiersResult.data;
  var baseThreshold = tiers[0] ? tiers[0].to || 0 : 0;
  var eligibleAmount = Math.max(0, totalSales - baseThreshold);
  
  // Calculate weighted average rate
  var totalIncentive = calculateTieredIncentive(totalSales);
  var effectiveRate = eligibleAmount > 0 ? totalIncentive / eligibleAmount : 0;
  
  return {
    baseThreshold: baseThreshold,
    eligibleAmount: eligibleAmount,
    effectiveRate: effectiveRate
  };
}

// ─── INCENTIVE BREAKDOWN FOR DISPLAY ──────────────────────────

function getIncentiveBreakdownForSales(totalSales) {
  var tiersResult = getIncentiveTiers();
  if (!tiersResult.success) return { total: 0, breakdown: [] };
  
  var tiers = tiersResult.data;
  var breakdown = [];
  var totalIncentive = 0;
  
  for (var i = 0; i < tiers.length; i++) {
    var tier = tiers[i];
    var from = tier.from;
    var to = tier.to || Infinity;
    var rate = tier.rate;
    
    if (totalSales <= from) break;
    
    var applicableAmount = Math.min(totalSales, to) - from;
    if (applicableAmount > 0) {
      var tierIncentive = applicableAmount * rate;
      totalIncentive += tierIncentive;
      
      breakdown.push({
        tier: tier.label || ('RM ' + from.toLocaleString() + ' - ' + (to ? 'RM ' + to.toLocaleString() : 'Above')),
        amount: applicableAmount,
        rate: (rate * 100).toFixed(2) + '%',
        incentive: tierIncentive
      });
    }
  }
  
  return {
    success: true,
    total: totalIncentive,
    breakdown: breakdown
  };
}

// ─── GENERATE DETAILED INCENTIVE PDF ──────────────────────────

function generateDetailedIncentivePDF(filters, user) {
  requireHROrAdmin(user);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCENTIVES);
  var data = sheet.getDataRange().getValues();
  
  var records = [];
  var filterYear = filters.year ? String(filters.year) : '';
  var filterRep = filters.salesRep ? String(filters.salesRep).toLowerCase() : '';
  var filterMonth = filters.month ? String(filters.month) : '';
  var dateFrom = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var dateTo = filters.dateTo ? new Date(filters.dateTo) : null;
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = data[i];
    
    // Apply filters
    if (filterYear && String(row[4]) !== filterYear) continue;
    if (filterRep && String(row[1]).toLowerCase().indexOf(filterRep) === -1) continue;
    if (filterMonth && String(row[3]) !== filterMonth) continue;
    
    // Date range filter on calculatedAt
    if (dateFrom || dateTo) {
      var calcDate = new Date(row[12]);
      if (dateFrom && calcDate < dateFrom) continue;
      if (dateTo && calcDate > dateTo) continue;
    }
    
    records.push({
      id: String(row[0]),
      salesRep: String(row[1]),
      email: String(row[2]),
      month: String(row[3]),
      year: row[4],
      totalSales: row[5],
      baseThreshold: row[6],
      eligibleAmount: row[7],
      incentiveRate: row[8],
      incentiveAmount: row[9],
      status: String(row[10]),
      paidDate: String(row[11]),
      calculatedAt: String(row[12])
    });
  }
  
  if (records.length === 0) {
    return { success: false, error: 'No records found for the selected filters' };
  }
  
  // Get current tier config
  var tiersResult = getIncentiveTiers();
  var tiers = tiersResult.success ? tiersResult.data : [];
  
  // Build enhanced PDF with tier breakdown
  var html = buildIncentivePDFWithTiers(records, tiers, filters);
  
  var blob = Utilities.newBlob(html, 'text/html', 'incentive_report.html');
  var pdfBlob = blob.getAs('application/pdf').setName('Incentive_Report_' +
    Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf');
  
  var folder = getSubFolder('Reports');
  var file = folder.createFile(pdfBlob);
  
  logActivity(user.email, 'EXPORT_DETAILED_INCENTIVE_PDF', records.length + ' records');
  
  return {
    success: true,
    downloadUrl: file.getUrl(),
    filename: file.getName(),
    recordCount: records.length
  };
}

function buildIncentivePDFWithTiers(records, tiers, filters) {
  var totalSales = 0;
  var totalIncentives = 0;
  
  records.forEach(function(r) {
    totalSales += parseFloat(r.totalSales || 0);
    totalIncentives += parseFloat(r.incentiveAmount || 0);
  });
  
  var filterStr = '';
  if (filters.year) filterStr += 'Year: ' + filters.year + ' ';
  if (filters.month) filterStr += 'Month: ' + filters.month + ' ';
  if (filters.salesRep) filterStr += 'Sales Rep: ' + filters.salesRep + ' ';
  if (filters.dateFrom) filterStr += 'From: ' + filters.dateFrom + ' ';
  if (filters.dateTo) filterStr += 'To: ' + filters.dateTo;
  if (!filterStr) filterStr = 'All Records';
  
  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;margin:25px;font-size:10px;}' +
    '.header{text-align:center;border-bottom:3px solid #1565c0;padding-bottom:12px;margin-bottom:18px;}' +
    '.header h1{color:#1a237e;margin:0;font-size:18px;}' +
    '.filters{background:#f0f4ff;padding:10px;border-radius:6px;margin-bottom:12px;font-size:9px;}' +
    'table{width:100%;border-collapse:collapse;margin:10px 0;font-size:9px;}' +
    'th{background:#1565c0;color:#fff;padding:7px;text-align:left;}' +
    'td{padding:6px 7px;border-bottom:1px solid #eee;}' +
    '.summary{background:#f8fafc;font-weight:bold;}' +
    '.tier-box{background:#e8f4fd;padding:8px;border-radius:4px;margin:10px 0;font-size:9px;}' +
    '.tier-box table{margin:5px 0;}' +
    '.tier-box th{background:#42a5f5;font-size:8px;padding:4px;}' +
    '.tier-box td{padding:3px 5px;font-size:8px;}' +
    '.amount{text-align:right;font-weight:bold;color:#2e7d32;}' +
    '</style></head><body>' +
    '<div class="header">' +
      '<h1>SandalMist - Detailed Incentive Report</h1>' +
      '<p style="margin:5px 0 0;color:#666;">Generated: ' +
      Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy HH:mm') + '</p>' +
    '</div>' +
    '<div class="filters"><strong>Filters:</strong> ' + filterStr + '</div>';
  
  // Tier structure table
  if (tiers && tiers.length > 0) {
    html += '<div class="tier-box">' +
      '<strong>Current Incentive Tier Structure</strong>' +
      '<table>' +
        '<tr><th>From (RM)</th><th>To (RM)</th><th>Rate</th><th>Label</th></tr>';
    
    tiers.forEach(function(t) {
      html += '<tr>' +
        '<td>' + formatNumber(t.from) + '</td>' +
        '<td>' + (t.to ? formatNumber(t.to) : 'Above') + '</td>' +
        '<td>' + (t.rate * 100).toFixed(2) + '%</td>' +
        '<td>' + (t.label || '-') + '</td>' +
      '</tr>';
    });
    
    html += '</table></div>';
  }
  
  // Records table
  html += '<table>' +
    '<thead><tr>' +
      '<th>ID</th><th>Sales Rep</th><th>Period</th>' +
      '<th>Total Sales (RM)</th><th>Incentive (RM)</th><th>Status</th>' +
    '</tr></thead><tbody>';
  
  records.forEach(function(r) {
    html += '<tr>' +
      '<td>' + r.id + '</td>' +
      '<td>' + r.salesRep + '</td>' +
      '<td>' + r.month + ' ' + r.year + '</td>' +
      '<td class="amount">' + formatNumber(r.totalSales) + '</td>' +
      '<td class="amount">' + formatNumber(r.incentiveAmount) + '</td>' +
      '<td>' + r.status + '</td>' +
    '</tr>';
  });
  
  html += '<tr class="summary">' +
    '<td colspan="3">TOTAL</td>' +
    '<td class="amount">' + formatNumber(totalSales) + '</td>' +
    '<td class="amount">' + formatNumber(totalIncentives) + '</td>' +
    '<td></td>' +
  '</tr></tbody></table>' +
  '<div style="margin-top:15px;font-size:8px;color:#999;text-align:center;border-top:1px solid #eee;padding-top:8px;">' +
    'SandalMist SM SalesBoard | Confidential | ' + records.length + ' record(s)' +
  '</div></body></html>';
  
  return html;
}

function formatNumber(n) {
  return Number(n || 0).toLocaleString('en-MY', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

