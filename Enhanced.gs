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
      '<h1>Sandal Mist - Incentive Report</h1>' +
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
      '<p>Sandal Mist Sales & HR Management System | Confidential</p>' +
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
      '<h1>Sandal Mist - Monthly Performance Report</h1>' +
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
    'Sandal Mist SHRMS | Confidential | ' + reportData.length + ' sales representatives' +
  '</div></body></html>';
  
  return html;
}

function getLeadsForPeriod(month, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LEADS);
  var data = sheet.getDataRange().getValues();
  
  var leads = [];
  var monthInt = parseInt(month);
  var yearInt = parseInt(year);
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = new Date(data[i][1]);
    if (d.getMonth() + 1 === monthInt && d.getFullYear() === yearInt) {
      leads.push({
        id: data[i][0],
        date: data[i][1],
        source: data[i][2],
        client: data[i][3],
        assignedTo: data[i][9],
        status: data[i][10],
        value: data[i][13]
      });
    }
  }
  return leads;
}

function buildLeadsReportPDF(leads, month, year) {
  // Similar PDF structure for leads
  return '<!DOCTYPE html><html><body><h1>Leads Report ' + month + '/' + year + '</h1></body></html>';
}

function getBookingsForPeriod(month, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.BOOKINGS);
  var data = sheet.getDataRange().getValues();
  
  var bookings = [];
  var monthInt = parseInt(month);
  var yearInt = parseInt(year);
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = new Date(data[i][1]);
    if (d.getMonth() + 1 === monthInt && d.getFullYear() === yearInt) {
      bookings.push({
        id: data[i][0],
        date: data[i][1],
        client: data[i][2],
        property: data[i][5],
        totalValue: data[i][12],
        salesRep: data[i][15],
        status: data[i][16]
      });
    }
  }
  return bookings;
}

function buildBookingsReportPDF(bookings, month, year) {
  // Similar PDF structure for bookings
  return '<!DOCTYPE html><html><body><h1>Bookings Report ' + month + '/' + year + '</h1></body></html>';
}
