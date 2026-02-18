/**
 * ============================================================
 * Enhanced.gs - Advanced Features
 * - Filtering & Sorting for all tables
 * - Commission Settings (Admin Panel)
 * - PDF Report Generation for Incentives
 * - Improved notification management
 * ============================================================
 */

// ─── COMMISSION SETTINGS (Admin Feature) ──────────────────────

function getCommissionSettings() {
  var props = PropertiesService.getScriptProperties();
  var settings = {
    baseThreshold:  parseFloat(props.getProperty('commission_baseThreshold')  || '500000'),
    incentiveRate:  parseFloat(props.getProperty('commission_incentiveRate')  || '0.01'),
    commissionRate: parseFloat(props.getProperty('commission_commissionRate') || '0.10'),
    lastUpdated:    props.getProperty('commission_lastUpdated') || '',
    updatedBy:      props.getProperty('commission_updatedBy') || ''
  };
  return { success: true, data: settings };
}

function updateCommissionSettings(data, user) {
  requireAdmin(user);
  
  var baseThreshold  = parseFloat(data.baseThreshold);
  var incentiveRate  = parseFloat(data.incentiveRate);
  var commissionRate = parseFloat(data.commissionRate);
  
  if (isNaN(baseThreshold) || isNaN(incentiveRate) || isNaN(commissionRate)) {
    return { success: false, error: 'Invalid numeric values' };
  }
  
  if (baseThreshold < 0 || incentiveRate < 0 || incentiveRate > 1 || 
      commissionRate < 0 || commissionRate > 1) {
    return { success: false, error: 'Values out of acceptable range' };
  }
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('commission_baseThreshold',  String(baseThreshold));
  props.setProperty('commission_incentiveRate',  String(incentiveRate));
  props.setProperty('commission_commissionRate', String(commissionRate));
  props.setProperty('commission_lastUpdated', new Date().toISOString());
  props.setProperty('commission_updatedBy', user.email);
  
  logActivity(user.email, 'UPDATE_COMMISSION', 
    'Base: ' + baseThreshold + ', Incentive: ' + (incentiveRate*100) + '%, Commission: ' + (commissionRate*100) + '%');
  
  return { 
    success: true, 
    message: 'Commission settings updated successfully'
  };
}

// ─── ENHANCED INCENTIVE REPORT WITH PDF ───────────────────────

function generateIncentivePDF(filters, user) {
  requireHROrAdmin(user);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCENTIVES);
  var data = sheet.getDataRange().getValues();
  
  var records = [];
  var filterYear = filters.year ? String(filters.year) : '';
  var filterRep  = filters.salesRep ? String(filters.salesRep).toLowerCase() : '';
  var filterMonth = filters.month ? String(filters.month) : '';
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = data[i];
    
    // Apply filters
    if (filterYear && String(row[4]) !== filterYear) continue;
    if (filterRep && String(row[1]).toLowerCase().indexOf(filterRep) === -1) continue;
    if (filterMonth && String(row[3]) !== filterMonth) continue;
    
    records.push({
      id:              String(row[0]),
      salesRep:        String(row[1]),
      email:           String(row[2]),
      month:           String(row[3]),
      year:            row[4],
      totalSales:      row[5],
      baseThreshold:   row[6],
      eligibleAmount:  row[7],
      incentiveRate:   row[8],
      incentiveAmount: row[9],
      status:          String(row[10]),
      paidDate:        String(row[11]),
      calculatedAt:    String(row[12])
    });
  }
  
  if (records.length === 0) {
    return { success: false, error: 'No records found for the selected filters' };
  }
  
  // Calculate totals
  var totalSales = 0, totalIncentives = 0;
  for (var j = 0; j < records.length; j++) {
    totalSales += parseFloat(records[j].totalSales || 0);
    totalIncentives += parseFloat(records[j].incentiveAmount || 0);
  }
  
  // Build HTML for PDF
  var filterStr = '';
  if (filterYear) filterStr += 'Year: ' + filterYear + ' ';
  if (filterMonth) filterStr += 'Month: ' + filterMonth + ' ';
  if (filterRep) filterStr += 'Sales Rep: ' + filterRep;
  if (!filterStr) filterStr = 'All Records';
  
  var html = 
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body { font-family: Arial, sans-serif; margin: 30px; font-size: 11px; }' +
    '.header { text-align: center; border-bottom: 3px solid #1565c0; padding-bottom: 15px; margin-bottom: 20px; }' +
    '.header h1 { color: #1a237e; margin: 0; font-size: 20px; }' +
    '.filters { background: #f0f4ff; padding: 10px; border-radius: 6px; margin-bottom: 15px; font-size: 10px; }' +
    'table { width: 100%; border-collapse: collapse; margin: 10px 0; }' +
    'th { background: #1565c0; color: #fff; padding: 8px; text-align: left; font-size: 10px; }' +
    'td { padding: 7px 8px; border-bottom: 1px solid #eee; font-size: 10px; }' +
    '.summary { background: #f8fafc; font-weight: bold; }' +
    '.footer { text-align: center; margin-top: 20px; font-size: 9px; color: #999; border-top: 1px solid #eee; padding-top: 10px; }' +
    '.amount { text-align: right; font-weight: bold; color: #2e7d32; }' +
    '</style></head><body>' +
    '<div class="header">' +
      '<h1>SandalMist - Incentive Report</h1>' +
      '<p style="margin:5px 0 0;color:#666;">Generated: ' + Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy HH:mm') + '</p>' +
    '</div>' +
    '<div class="filters"><strong>Filters Applied:</strong> ' + filterStr + '</div>' +
    '<table>' +
      '<thead><tr>' +
        '<th>ID</th><th>Sales Rep</th><th>Period</th>' +
        '<th>Total Sales (RM)</th><th>Base Threshold (RM)</th><th>Eligible (RM)</th>' +
        '<th>Rate</th><th>Incentive (RM)</th><th>Status</th>' +
      '</tr></thead>' +
      '<tbody>';
  
  for (var k = 0; k < records.length; k++) {
    var r = records[k];
    html += '<tr>' +
      '<td>' + r.id + '</td>' +
      '<td>' + r.salesRep + '</td>' +
      '<td>' + r.month + ' ' + r.year + '</td>' +
      '<td class="amount">' + formatNumber(r.totalSales) + '</td>' +
      '<td class="amount">' + formatNumber(r.baseThreshold) + '</td>' +
      '<td class="amount">' + formatNumber(r.eligibleAmount) + '</td>' +
      '<td>' + ((r.incentiveRate || 0) * 100).toFixed(1) + '%</td>' +
      '<td class="amount">' + formatNumber(r.incentiveAmount) + '</td>' +
      '<td>' + r.status + '</td>' +
    '</tr>';
  }
  
  html += '<tr class="summary">' +
    '<td colspan="3">TOTAL</td>' +
    '<td class="amount">' + formatNumber(totalSales) + '</td>' +
    '<td colspan="3"></td>' +
    '<td class="amount">' + formatNumber(totalIncentives) + '</td>' +
    '<td></td>' +
  '</tr>';
  
  html += '</tbody></table>' +
    '<div class="footer">' +
      '<p>SandalMist Sales & Marketing Management System | Confidential</p>' +
      '<p>This report contains ' + records.length + ' incentive record(s)</p>' +
    '</div>' +
  '</body></html>';
  
  // Convert to PDF
  var blob = Utilities.newBlob(html, 'text/html', 'incentive_report.html');
  var pdfBlob = blob.getAs('application/pdf').setName('Incentive_Report_' + 
    Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf');
  
  var folder = getSubFolder('Reports');
  var file = folder.createFile(pdfBlob);
  
  logActivity(user.email, 'EXPORT_INCENTIVE_PDF', 'Filters: ' + filterStr);
  
  return {
    success: true,
    downloadUrl: file.getUrl(),
    filename: file.getName(),
    recordCount: records.length
  };
}

function formatNumber(n) {
  return Number(n || 0).toLocaleString('en-MY', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

// ─── MANAGEMENT REPORTS PDF GENERATION ────────────────────────

function generateManagementReportPDF(data, user) {
  requireHROrAdmin(user);
  
  var month = data.month;
  var year = data.year;
  var reportType = data.reportType || 'monthly'; // 'monthly', 'leads', 'bookings'
  
  if (!month || !year) {
    return { success: false, error: 'Month and year required' };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var html = '';
  var filename = '';
  
  if (reportType === 'monthly') {
    var report = getMonthlyReport({ month: month, year: year }, user);
    if (!report.success || !report.data || report.data.length === 0) {
      return { success: false, error: 'No data available for this period' };
    }
    
    html = buildMonthlyReportPDF(report.data, month, year);
    filename = 'Monthly_Report_' + month + '_' + year + '.pdf';
    
  } else if (reportType === 'leads') {
    var leads = getLeadsForPeriod(month, year);
    html = buildLeadsReportPDF(leads, month, year);
    filename = 'Leads_Report_' + month + '_' + year + '.pdf';
    
  } else if (reportType === 'bookings') {
    var bookings = getBookingsForPeriod(month, year);
    html = buildBookingsReportPDF(bookings, month, year);
    filename = 'Bookings_Report_' + month + '_' + year + '.pdf';
  }
  
  // Convert to PDF
  var blob = Utilities.newBlob(html, 'text/html', filename.replace('.pdf', '.html'));
  var pdfBlob = blob.getAs('application/pdf').setName(filename);
  
  var folder = getSubFolder('Reports');
  var file = folder.createFile(pdfBlob);
  
  logActivity(user.email, 'EXPORT_MGMT_PDF', reportType + ' - ' + month + '/' + year);
  
  return {
    success: true,
    downloadUrl: file.getUrl(),
    filename: filename
  };
}

function buildMonthlyReportPDF(reportData, month, year) {
  var totals = reportData.reduce(function(t, r) {
    t.revenue += r.revenue || 0;
    t.bookings += r.bookings || 0;
    t.newLeads += r.newLeads || 0;
    t.wonLeads += r.wonLeads || 0;
    return t;
  }, { revenue: 0, bookings: 0, newLeads: 0, wonLeads: 0 });
  
  var html = 
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body { font-family: Arial, sans-serif; margin: 25px; font-size: 11px; }' +
    '.header { text-align: center; border-bottom: 3px solid #1565c0; padding-bottom: 12px; margin-bottom: 18px; }' +
    'h1 { color: #1a237e; margin: 0; font-size: 18px; }' +
    'table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 10px; }' +
    'th { background: #1565c0; color: #fff; padding: 8px; text-align: left; }' +
    'td { padding: 7px 8px; border-bottom: 1px solid #eee; }' +
    '.summary { background: #f0f4ff; font-weight: bold; }' +
    '.amount { text-align: right; font-weight: bold; }' +
    '.rank { text-align: center; font-weight: bold; color: #f4b400; }' +
    '</style></head><body>' +
    '<div class="header">' +
      '<h1>SandalMist - Monthly Performance Report</h1>' +
      '<p style="margin:5px 0;">Period: ' + month + ' ' + year + '</p>' +
      '<p style="margin:0;font-size:10px;color:#666;">Generated: ' + 
        Utilities.formatDate(new Date(), 'UTC', 'dd MMM yyyy HH:mm') + '</p>' +
    '</div>' +
    '<h3 style="margin:15px 0 8px;">Summary Statistics</h3>' +
    '<table>' +
      '<tr><th>Metric</th><th>Value</th></tr>' +
      '<tr><td>Total Revenue</td><td class="amount">RM ' + formatNumber(totals.revenue) + '</td></tr>' +
      '<tr><td>Total Bookings</td><td class="amount">' + totals.bookings + '</td></tr>' +
      '<tr><td>New Leads</td><td class="amount">' + totals.newLeads + '</td></tr>' +
      '<tr><td>Won Leads</td><td class="amount">' + totals.wonLeads + '</td></tr>' +
      '<tr><td>Conversion Rate</td><td class="amount">' + 
        (totals.newLeads > 0 ? Math.round((totals.wonLeads / totals.newLeads) * 100) : 0) + '%</td></tr>' +
    '</table>' +
    '<h3 style="margin:20px 0 8px;">Team Performance Rankings</h3>' +
    '<table>' +
      '<thead><tr>' +
        '<th>Rank</th><th>Sales Rep</th><th>Revenue (RM)</th><th>Bookings</th>' +
        '<th>Leads</th><th>Won</th><th>Conv %</th><th>DSR</th>' +
      '</tr></thead><tbody>';
  
  for (var i = 0; i < reportData.length; i++) {
    var r = reportData[i];
    html += '<tr>' +
      '<td class="rank">' + (i + 1) + '</td>' +
      '<td>' + r.rep + '</td>' +
      '<td class="amount">' + formatNumber(r.revenue) + '</td>' +
      '<td class="amount">' + r.bookings + '</td>' +
      '<td class="amount">' + r.newLeads + '</td>' +
      '<td class="amount">' + r.wonLeads + '</td>' +
      '<td class="amount">' + r.conversionRate + '%</td>' +
      '<td class="amount">' + r.dsrCount + '</td>' +
    '</tr>';
  }
  
  html += '<tr class="summary">' +
    '<td colspan="2">TOTAL</td>' +
    '<td class="amount">' + formatNumber(totals.revenue) + '</td>' +
    '<td class="amount">' + totals.bookings + '</td>' +
    '<td class="amount">' + totals.newLeads + '</td>' +
    '<td class="amount">' + totals.wonLeads + '</td>' +
    '<td colspan="2"></td>' +
  '</tr></tbody></table>' +
  '<div style="margin-top:20px;font-size:9px;color:#999;text-align:center;border-top:1px solid #eee;padding-top:10px;">' +
    'SandalMist SM SalesBoard | Confidential | ' + reportData.length + ' sales representatives' +
  '</div></body></html>';
  
  return html;
}

function getLeadsForPeriod(month, year, status) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEETS.LEADS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data     = sheet.getRange(1, 1, lastRow, 17).getValues();
  var monthInt = month ? parseInt(month) : 0;
  var yearInt  = year  ? parseInt(year)  : 0;
  var statusFlt = (status || '').toLowerCase();

  var leads = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var r = data[i];
    if (monthInt && yearInt) {
      var d = new Date(r[1]);
      if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) continue;
    }
    if (statusFlt && String(r[10]).toLowerCase() !== statusFlt) continue;
    leads.push({
      id: r[0], date: r[1], source: r[2], client: r[3],
      contactPerson: r[4], assignedTo: r[9], status: r[10], value: r[13]
    });
  }
  leads.reverse();
  return leads;
}

// ─── SHARED PDF HELPERS ───────────────────────────────────────

function pdfCss() {
  return '<style>' +
    'body{font-family:Arial,sans-serif;margin:28px;font-size:11px;color:#222;}' +
    '.ph{text-align:center;border-bottom:3px solid #1565c0;padding-bottom:12px;margin-bottom:18px;}' +
    '.ph img{height:36px;vertical-align:middle;margin-right:10px;}' +
    '.ph h1{display:inline;color:#1a237e;font-size:18px;vertical-align:middle;}' +
    '.ph p{margin:4px 0 0;font-size:10px;color:#666;}' +
    '.meta{background:#f0f4ff;padding:8px 12px;border-radius:5px;margin-bottom:14px;font-size:10px;}' +
    'table{width:100%;border-collapse:collapse;margin:10px 0;font-size:10px;}' +
    'th{background:#1565c0;color:#fff;padding:7px 8px;text-align:left;}' +
    'td{padding:6px 8px;border-bottom:1px solid #eee;}' +
    'tr:nth-child(even) td{background:#f8fafc;}' +
    '.sum td{background:#e8f4fd;font-weight:700;}' +
    '.amt{text-align:right;font-weight:700;color:#2e7d32;}' +
    '.foot{text-align:center;margin-top:18px;font-size:9px;color:#999;' +
      'border-top:1px solid #eee;padding-top:8px;}' +
    '.badge{display:inline-block;padding:2px 7px;border-radius:10px;font-size:9px;font-weight:600;}' +
    '.b-pending{background:#fff3e0;color:#e65100;}' +
    '.b-approved,.b-confirmed,.b-completed,.b-won{background:#e8f5e9;color:#2e7d32;}' +
    '.b-rejected,.b-lost{background:#fce4ec;color:#c62828;}' +
    '.b-new,.b-qualified{background:#e3f2fd;color:#1565c0;}' +
    '</style>';
}

function pdfHeader(title, period) {
  var ts = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd MMM yyyy, hh:mm a');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' + pdfCss() + '</head><body>' +
    '<div class="ph">' +
      '<img src="https://sandalmistresort.com/wp-content/uploads/2024/09/logo-white.png"' +
      ' style="height:36px;background:#1565c0;border-radius:4px;padding:4px 8px;">' +
      '<h1>SandalMist — ' + title + '</h1>' +
      (period ? '<p>Period: ' + period + '</p>' : '') +
      '<p>Generated: ' + ts + '</p>' +
    '</div>';
}

function pdfFooter(count, label) {
  return '<div class="foot"><p>SandalMist Sales &amp; Marketing Management System | Confidential</p>' +
    '<p>This report contains ' + count + ' ' + (label || 'record') + (count !== 1 ? 's' : '') + '</p>' +
    '</div></body></html>';
}

function statusBadgePDF(status) {
  var cls = (status || '').toLowerCase().replace(/\s+/g, '');
  return '<span class="badge b-' + cls + '">' + (status || '') + '</span>';
}

function savePDF(html, filename) {
  var blob    = Utilities.newBlob(html, 'text/html', filename.replace('.pdf', '.html'));
  var pdfBlob = blob.getAs('application/pdf').setName(filename);
  return getSubFolder('Reports').createFile(pdfBlob);
}

// ─── DSR LIST PDF ─────────────────────────────────────────────

function generateDSRListPDF(filters, user) {
  requireHROrAdmin(user);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DAILY_REPORTS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No DSR records found' };

  var raw      = sheet.getRange(1, 1, lastRow, 14).getValues();
  var monthInt = filters.month ? parseInt(filters.month) : 0;
  var yearInt  = filters.year  ? parseInt(filters.year)  : 0;
  var repFilter = (filters.salesRep || '').toLowerCase();

  var rows = [];
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    var r = raw[i];
    if (monthInt && yearInt) {
      var d = new Date(r[1]);
      if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) continue;
    }
    if (repFilter && String(r[2]).toLowerCase().indexOf(repFilter) === -1) continue;
    rows.push(r);
  }
  rows.reverse();

  var period = monthInt ? (filters.month + '/' + filters.year) : 'All Periods';
  var html = pdfHeader('Daily Sales Reports', period) +
    '<div class="meta"><strong>Filter:</strong> Period: ' + period +
    (repFilter ? ' | Rep: ' + filters.salesRep : '') +
    ' | Total: ' + rows.length + ' records</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Date</th><th>Sales Rep</th><th>Location</th>' +
      '<th>Client</th><th>Purpose</th><th>Result</th><th>Status</th>' +
    '</tr></thead><tbody>';

  rows.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r[0] + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r[1]), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td><b>' + r[2] + '</b></td><td>' + r[4] + '</td>' +
      '<td>' + r[5] + '</td>' +
      '<td style="max-width:160px;">' + String(r[7]).substring(0, 80) + (r[7].length > 80 ? '...' : '') + '</td>' +
      '<td style="max-width:140px;">' + String(r[8]).substring(0, 70) + (r[8].length > 70 ? '...' : '') + '</td>' +
      '<td>' + statusBadgePDF(r[12]) + '</td>' +
    '</tr>';
  });

  html += '</tbody></table>' + pdfFooter(rows.length, 'daily report');
  var file = savePDF(html, 'DSR_Report_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf');
  logActivity(user.email, 'EXPORT_DSR_PDF', rows.length + ' records | ' + period);
  return { success: true, downloadUrl: file.getUrl(), filename: file.getName(), recordCount: rows.length };
}

// ─── TRAVEL PLANS PDF ────────────────────────────────────────

function generateTravelPDF(filters, user) {
  requireHROrAdmin(user);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TRAVEL_PLANS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No travel records found' };

  var raw       = sheet.getRange(1, 1, lastRow, 17).getValues();
  var statusFlt = (filters.status || '').toLowerCase();
  var repFilter = (filters.salesRep || '').toLowerCase();

  var dateFrom = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var dateTo   = filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  var rows = [], total = 0;
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    var r = raw[i];
    if (statusFlt && String(r[12]).toLowerCase() !== statusFlt) continue;
    if (repFilter && String(r[1]).toLowerCase().indexOf(repFilter) === -1) continue;
    if (dateFrom || dateTo) {
      var d = new Date(r[3]);
      if (!isNaN(d)) {
        if (dateFrom && d < dateFrom) continue;
        if (dateTo   && d > dateTo)   continue;
      }
    }
    rows.push(r);
    total += parseFloat(r[8]) || 0;
  }
  rows.reverse();

  var periodLabel = filters.dateFrom ? (filters.dateFrom + ' to ' + (filters.dateTo || 'now')) : 'All';
  var html = pdfHeader('Travel Plans Report', periodLabel) +
    '<div class="meta"><strong>Filter:</strong> Status: ' + (filters.status || 'All') +
    (repFilter ? ' | Rep: ' + filters.salesRep : '') +
    (filters.dateFrom ? ' | Period: ' + periodLabel : '') +
    ' | Total: ' + rows.length + ' records' +
    ' | Expected Revenue: RM ' + formatNumber(total) + '</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Sales Rep</th><th>Travel Date</th><th>City</th>' +
      '<th>Purpose</th><th>Exp. Revenue</th><th>Days</th><th>Status</th>' +
    '</tr></thead><tbody>';

  rows.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r[0] + '</td>' +
      '<td><b>' + r[1] + '</b></td>' +
      '<td>' + Utilities.formatDate(new Date(r[3]), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td>' + r[5] + '</td>' +
      '<td style="max-width:160px;">' + String(r[7]).substring(0, 80) + '</td>' +
      '<td class="amt">RM ' + formatNumber(r[8]) + '</td>' +
      '<td style="text-align:center;">' + r[9] + 'd</td>' +
      '<td>' + statusBadgePDF(r[12]) + '</td>' +
    '</tr>';
  });

  html += '<tr class="sum"><td colspan="5"><b>TOTAL</b></td>' +
    '<td class="amt">RM ' + formatNumber(total) + '</td><td colspan="2"></td></tr>';
  html += '</tbody></table>' + pdfFooter(rows.length, 'travel plan');

  var file = savePDF(html, 'Travel_Report_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf');
  logActivity(user.email, 'EXPORT_TRAVEL_PDF', rows.length + ' records');
  return { success: true, downloadUrl: file.getUrl(), filename: file.getName(), recordCount: rows.length };
}

// ─── LEADS PDF ────────────────────────────────────────────────

function buildLeadsReportPDF(leads, month, year) {
  var total = leads.reduce(function(s, r) { return s + (parseFloat(r.value) || 0); }, 0);
  var html  = pdfHeader('Leads Report', month + ' ' + year) +
    '<div class="meta"><strong>Period:</strong> ' + month + ' ' + year +
    ' | Total leads: ' + leads.length + ' | Total value: RM ' + formatNumber(total) + '</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Date</th><th>Client</th><th>Source</th>' +
      '<th>Assigned To</th><th>Budget (RM)</th><th>Status</th>' +
    '</tr></thead><tbody>';

  leads.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r.id + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r.date), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td><b>' + r.client + '</b></td>' +
      '<td>' + r.source + '</td>' +
      '<td>' + r.assignedTo + '</td>' +
      '<td class="amt">' + (r.value ? 'RM ' + formatNumber(r.value) : '-') + '</td>' +
      '<td>' + statusBadgePDF(r.status) + '</td>' +
    '</tr>';
  });

  html += '<tr class="sum"><td colspan="5"><b>TOTAL</b></td>' +
    '<td class="amt">RM ' + formatNumber(total) + '</td><td></td></tr>';
  html += '</tbody></table>' + pdfFooter(leads.length, 'lead');
  return html;
}

function generateLeadsPDF(filters, user) {
  requireHROrAdmin(user);

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEETS.LEADS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No lead records found' };

  var raw       = sheet.getRange(1, 1, lastRow, 17).getValues();
  var statusFlt = (filters.status   || '').toLowerCase();
  var repFilter = (filters.salesRep || '').toLowerCase();
  var dateFrom  = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var dateTo    = filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  var leads = [], totalVal = 0;
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    var r = raw[i];
    if (statusFlt && String(r[10]).toLowerCase() !== statusFlt) continue;
    if (repFilter && String(r[9]).toLowerCase().indexOf(repFilter) === -1) continue;
    if (dateFrom || dateTo) {
      var d = new Date(r[1]);
      if (!isNaN(d)) {
        if (dateFrom && d < dateFrom) continue;
        if (dateTo   && d > dateTo)   continue;
      }
    }
    leads.push(r);
    totalVal += parseFloat(r[8]) || 0;
  }
  leads.reverse();

  if (!leads.length) return { success: false, error: 'No leads for the selected filters' };

  var periodLabel = filters.dateFrom ? (filters.dateFrom + ' to ' + (filters.dateTo || 'now')) : 'All';
  var html = pdfHeader('Leads Report', periodLabel) +
    '<div class="meta"><strong>Filter:</strong> Status: ' + (filters.status || 'All') +
    (repFilter ? ' | Rep: ' + filters.salesRep : '') +
    (filters.dateFrom ? ' | Period: ' + periodLabel : '') +
    ' | Total: ' + leads.length + ' records' +
    ' | Total Value: RM ' + formatNumber(totalVal) + '</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Date</th><th>Client</th><th>Source</th>' +
      '<th>Assigned To</th><th>Budget (RM)</th><th>Status</th>' +
    '</tr></thead><tbody>';

  leads.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r[0] + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r[1]), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td><b>' + r[3] + '</b></td>' +
      '<td>' + r[2] + '</td>' +
      '<td>' + r[9] + '</td>' +
      '<td class="amt">' + (r[8] ? 'RM ' + formatNumber(r[8]) : '-') + '</td>' +
      '<td>' + statusBadgePDF(r[10]) + '</td>' +
    '</tr>';
  });

  html += '<tr class="sum"><td colspan="5"><b>TOTAL</b></td>' +
    '<td class="amt">RM ' + formatNumber(totalVal) + '</td><td></td></tr>';
  html += '</tbody></table>' + pdfFooter(leads.length, 'lead');

  var filename = 'Leads_Report_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf';
  var file = savePDF(html, filename);
  logActivity(user.email, 'EXPORT_LEADS_PDF', leads.length + ' records');
  return { success: true, downloadUrl: file.getUrl(), filename: file.getName(), recordCount: leads.length };
}

// ─── BOOKINGS PDF ─────────────────────────────────────────────

function buildBookingsReportPDF(bookings, month, year) {
  var total = bookings.reduce(function(s, r) { return s + (parseFloat(r.totalValue) || 0); }, 0);
  var html  = pdfHeader('Bookings Report', month + ' ' + year) +
    '<div class="meta"><strong>Period:</strong> ' + month + ' ' + year +
    ' | Total bookings: ' + bookings.length + ' | Total revenue: RM ' + formatNumber(total) + '</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Date</th><th>Client</th><th>Property</th>' +
      '<th>Check-In</th><th>Nights</th><th>Revenue (RM)</th>' +
      '<th>Sales Rep</th><th>Status</th>' +
    '</tr></thead><tbody>';

  bookings.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r.id + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r.date), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td><b>' + r.client + '</b></td>' +
      '<td>' + r.property + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r.checkIn || r.date), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td style="text-align:center;">' + (r.nights || '-') + '</td>' +
      '<td class="amt">RM ' + formatNumber(r.totalValue) + '</td>' +
      '<td>' + r.salesRep + '</td>' +
      '<td>' + statusBadgePDF(r.status) + '</td>' +
    '</tr>';
  });

  html += '<tr class="sum"><td colspan="6"><b>TOTAL REVENUE</b></td>' +
    '<td class="amt">RM ' + formatNumber(total) + '</td><td colspan="2"></td></tr>';
  html += '</tbody></table>' + pdfFooter(bookings.length, 'booking');
  return html;
}

function generateBookingsPDF(filters, user) {
  requireHROrAdmin(user);

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEETS.BOOKINGS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No booking records found' };

  var raw       = sheet.getRange(1, 1, lastRow, 20).getValues();
  var statusFlt = (filters.status   || '').toLowerCase();
  var repFilter = (filters.salesRep || '').toLowerCase();
  var dateFrom  = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var dateTo    = filters.dateTo   ? new Date(filters.dateTo)   : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  var bookings = [], totalRev = 0;
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i][0]) continue;
    var r = raw[i];
    if (statusFlt && String(r[16]).toLowerCase() !== statusFlt) continue;
    if (repFilter && String(r[15]).toLowerCase().indexOf(repFilter) === -1) continue;
    if (dateFrom || dateTo) {
      var d = new Date(r[1]);
      if (!isNaN(d)) {
        if (dateFrom && d < dateFrom) continue;
        if (dateTo   && d > dateTo)   continue;
      }
    }
    bookings.push(r);
    totalRev += parseFloat(r[12]) || 0;
  }
  bookings.reverse();

  if (!bookings.length) return { success: false, error: 'No bookings for the selected filters' };

  var periodLabel = filters.dateFrom ? (filters.dateFrom + ' to ' + (filters.dateTo || 'now')) : 'All';
  var html = pdfHeader('Bookings Report', periodLabel) +
    '<div class="meta"><strong>Filter:</strong> Status: ' + (filters.status || 'All') +
    (repFilter ? ' | Rep: ' + filters.salesRep : '') +
    (filters.dateFrom ? ' | Period: ' + periodLabel : '') +
    ' | Total: ' + bookings.length + ' records' +
    ' | Total Revenue: RM ' + formatNumber(totalRev) + '</div>' +
    '<table><thead><tr>' +
      '<th>ID</th><th>Date</th><th>Client</th><th>Property</th>' +
      '<th>Check-In</th><th>Nights</th><th>Revenue (RM)</th>' +
      '<th>Sales Rep</th><th>Status</th>' +
    '</tr></thead><tbody>';

  bookings.forEach(function(r) {
    html += '<tr>' +
      '<td style="font-family:monospace;color:#1565c0;">' + r[0] + '</td>' +
      '<td>' + Utilities.formatDate(new Date(r[1]), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') + '</td>' +
      '<td><b>' + r[2] + '</b></td>' +
      '<td>' + r[5] + '</td>' +
      '<td>' + (r[7] ? Utilities.formatDate(new Date(r[7]), 'Asia/Kuala_Lumpur', 'dd-MM-yyyy') : '-') + '</td>' +
      '<td style="text-align:center;">' + (r[9] || '-') + '</td>' +
      '<td class="amt">RM ' + formatNumber(r[12]) + '</td>' +
      '<td>' + r[15] + '</td>' +
      '<td>' + statusBadgePDF(r[16]) + '</td>' +
    '</tr>';
  });

  html += '<tr class="sum"><td colspan="6"><b>TOTAL REVENUE</b></td>' +
    '<td class="amt">RM ' + formatNumber(totalRev) + '</td><td colspan="2"></td></tr>';
  html += '</tbody></table>' + pdfFooter(bookings.length, 'booking');

  var filename = 'Bookings_Report_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss') + '.pdf';
  var file = savePDF(html, filename);
  logActivity(user.email, 'EXPORT_BOOKINGS_PDF', bookings.length + ' records');
  return { success: true, downloadUrl: file.getUrl(), filename: file.getName(), recordCount: bookings.length };
}

function getBookingsForPeriod(month, year, status) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEETS.BOOKINGS);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data      = sheet.getRange(1, 1, lastRow, 20).getValues();
  var monthInt  = month ? parseInt(month) : 0;
  var yearInt   = year  ? parseInt(year)  : 0;
  var statusFlt = (status || '').toLowerCase();

  var bookings = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var r = data[i];
    if (monthInt && yearInt) {
      var d = new Date(r[1]);
      if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) continue;
    }
    if (statusFlt && String(r[16]).toLowerCase() !== statusFlt) continue;
    bookings.push({
      id: r[0], date: r[1], client: r[2], property: r[5],
      checkIn: r[7], checkOut: r[8], nights: r[9],
      totalValue: r[12], salesRep: r[15], status: r[16]
    });
  }
  bookings.reverse();
  return bookings;
}

function buildBookingsReportPDF(bookings, month, year) {
  // Similar PDF structure for bookings
  return '<!DOCTYPE html><html><body><h1>Bookings Report ' + month + '/' + year + '</h1></body></html>';
}
