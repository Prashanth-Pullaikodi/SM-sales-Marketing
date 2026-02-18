/**
 * ============================================================
 * Email.gs - Email Automation System
 * ============================================================
 */

const EMAIL_LOGO = "https://sandalmistresort.com/wp-content/uploads/2024/09/logo-white.png";
const COMPANY_NAME = "SandalMist";
const COMPANY_TAGLINE = "Hospitality Sales Excellence";
const COMPANY_FOOTER = "SandalMist Sales & Marketing Management System | Confidential";

// ─── EMAIL HEADER/FOOTER TEMPLATES ───────────────────────────

function getEmailHeader(title, subtitle) {
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #333; }
  .wrapper { max-width: 680px; margin: 20px auto; }
  .header { background: linear-gradient(135deg, #1a237e 0%, #0d47a1 50%, #1565c0 100%);
    padding: 30px 40px; border-radius: 12px 12px 0 0; text-align: center; }
  .header img { max-width: 200px; height: 50px; object-fit: contain; margin-bottom: 12px; display: block; margin-left: auto; margin-right: auto; }
  .header h1 { color: #fff; font-size: 24px; font-weight: 700; letter-spacing: 1px; }
  .header p { color: rgba(255,255,255,0.8); font-size: 13px; margin-top: 4px; }
  .badge { display: inline-block; background: rgba(255,255,255,0.2);
    color: #fff; padding: 4px 14px; border-radius: 20px; font-size: 12px;
    margin-top: 10px; letter-spacing: 0.5px; }
  .body { background: #fff; padding: 35px 40px; }
  .title-bar { border-left: 4px solid #1565c0; padding: 8px 0 8px 16px;
    margin-bottom: 24px; }
  .title-bar h2 { font-size: 20px; color: #1a237e; font-weight: 700; }
  .title-bar p { font-size: 13px; color: #666; margin-top: 3px; }
  table.data-table { width: 100%; border-collapse: collapse; margin: 20px 0;
    border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
  table.data-table th { background: #1565c0; color: #fff; padding: 12px 16px;
    text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
  table.data-table td { padding: 11px 16px; font-size: 13px; color: #333;
    border-bottom: 1px solid #f0f0f0; }
  table.data-table tr:nth-child(even) td { background: #f8fafc; }
  table.data-table tr:last-child td { border-bottom: none; }
  .highlight-box { background: linear-gradient(135deg, #e8f4fd, #f0f8ff);
    border: 1px solid #b3d9f7; border-radius: 8px; padding: 18px 22px; margin: 20px 0; }
  .highlight-box .label { font-size: 11px; color: #666; text-transform: uppercase;
    letter-spacing: 0.5px; margin-bottom: 4px; }
  .highlight-box .value { font-size: 20px; font-weight: 700; color: #1565c0; }
  .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 600; letter-spacing: 0.3px; }
  .status-pending { background: #fff3e0; color: #e65100; }
  .status-approved { background: #e8f5e9; color: #2e7d32; }
  .status-rejected { background: #fce4ec; color: #c62828; }
  .status-submitted { background: #e3f2fd; color: #1565c0; }
  .btn { display: inline-block; padding: 12px 28px; background: #1565c0;
    color: #fff; text-decoration: none; border-radius: 6px; font-size: 14px;
    font-weight: 600; margin: 10px 0; }
  .btn-success { background: #2e7d32; }
  .btn-danger { background: #c62828; }
  .divider { border: none; border-top: 1px solid #eee; margin: 25px 0; }
  .section-label { font-size: 11px; color: #999; text-transform: uppercase;
    letter-spacing: 1px; margin-bottom: 8px; font-weight: 600; }
  .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin: 16px 0; }
  .info-item { background: #f8fafc; border-radius: 6px; padding: 10px 14px; }
  .info-item .key { font-size: 11px; color: #888; margin-bottom: 3px; }
  .info-item .val { font-size: 14px; color: #333; font-weight: 500; }
  .footer { background: #1a237e; color: rgba(255,255,255,0.7); text-align: center;
    padding: 20px 30px; border-radius: 0 0 12px 12px; font-size: 12px; }
  .footer a { color: rgba(255,255,255,0.9); text-decoration: none; }
  .footer .logo-footer { font-size: 16px; color: #fff; font-weight: 700; margin-bottom: 6px; }
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <img src="${EMAIL_LOGO}" alt="SandalMist Logo">
    <h1>${COMPANY_NAME}</h1>
    <p>${COMPANY_TAGLINE}</p>
    <span class="badge">SM SalesBoard Notification</span>
  </div>
  <div class="body">
    <div class="title-bar">
      <h2>${title}</h2>
      ${subtitle ? `<p>${subtitle}</p>` : ""}
    </div>`;
}

function getEmailFooter(timestamp) {
  const ts = timestamp || Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMM yyyy, hh:mm a");
  return `
    <hr class="divider">
    <p style="font-size:12px;color:#999;text-align:center;">
      This is an automated notification from ${APP_NAME}.<br>
      Generated on ${ts} | ${COMPANY_FOOTER}
    </p>
  </div>
  <div class="footer">
    <div class="logo-footer">🏨 ${COMPANY_NAME}</div>
    <p>Sales &amp; HR Management System</p>
    <p style="margin-top:8px;">
      <a href="mailto:admin@sandalmist.com">admin@sandalmist.com</a>
    </p>
  </div>
</div>
</body>
</html>`;
}

// ─── DSR EMAIL ────────────────────────────────────────────────

function sendDSREmail(id, data, user, pdfLink) {
  var recipients = getHREmailList();
  if (recipients.length === 0) {
    Logger.log("No HR recipients found for DSR email");
    return;
  }

  var timestamp = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMM yyyy, hh:mm a");

  var html = getEmailHeader("Daily Sales Report Submitted", "A new DSR has been filed and requires your attention") + 
    '<div class="info-grid">' +
      '<div class="info-item">' +
        '<div class="key">Report ID</div>' +
        '<div class="val" style="color:#1565c0;font-weight:700;">' + id + '</div>' +
      '</div>' +
      '<div class="info-item">' +
        '<div class="key">Submitted By</div>' +
        '<div class="val">' + user.name + '</div>' +
      '</div>' +
      '<div class="info-item">' +
        '<div class="key">Date</div>' +
        '<div class="val">' + data.date + '</div>' +
      '</div>' +
      '<div class="info-item">' +
        '<div class="key">Location</div>' +
        '<div class="val">' + data.location + '</div>' +
      '</div>' +
    '</div>' +
    '<table class="data-table">' +
      '<tr><th colspan="2">Report Details</th></tr>' +
      '<tr><td><strong>Client Name</strong></td><td>' + data.clientName + '</td></tr>' +
      '<tr><td><strong>Contact</strong></td><td>' + (data.contact || '-') + '</td></tr>' +
      '<tr><td><strong>Visit Purpose</strong></td><td>' + data.purpose + '</td></tr>' +
      '<tr><td><strong>Result / Outcome</strong></td><td>' + data.result + '</td></tr>' +
      '<tr><td><strong>Next Action</strong></td><td>' + (data.nextAction || '-') + '</td></tr>' +
      '<tr><td><strong>Status</strong></td><td><span class="status-badge status-submitted">Submitted</span></td></tr>' +
    '</table>' +
    (pdfLink ? 
      '<div class="highlight-box">' +
        '<div class="label">&#128206; PDF Attachment</div>' +
        '<p style="font-size:13px;margin-top:6px;">The full DSR report has been saved as PDF.</p>' +
        '<a href="' + pdfLink + '" class="btn">&#128196; View PDF Report</a>' +
      '</div>' : '') +
    getEmailFooter(timestamp);

  // Use HTML entities for emojis in subject
  var subject = 'DSR Submitted: ' + id + ' - ' + user.name + ' | ' + data.date;

  recipients.forEach(function(email) {
    try {
      GmailApp.sendEmail(email, subject, "Please view in HTML client.", { htmlBody: html });
    } catch (e) {
      Logger.log("Failed to send to " + email + ": " + e.message);
    }
  });

  Logger.log("DSR email sent to: " + recipients.join(", "));
}

// ─── TRAVEL APPROVAL EMAIL ────────────────────────────────────

function sendTravelApprovalEmail(id, data, user, rowIndex) {
  const recipients = getHREmailList();
  if (recipients.length === 0) return;

  // Generate a one-time secure token (valid 48 hours)
  const token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('tv_token_' + id, JSON.stringify({
    token:    token,
    rowIndex: rowIndex,
    expires:  Date.now() + (48 * 60 * 60 * 1000),
    used:     false
  }));

  const webUrl    = ScriptApp.getService().getUrl();
  const approveUrl = webUrl + '?tvAction=approve&id=' + encodeURIComponent(id) + '&token=' + token;
  const rejectUrl  = webUrl + '?tvAction=reject&id='  + encodeURIComponent(id) + '&token=' + token;

  const html = getEmailHeader("Travel Plan Approval Required",
    "A new travel plan needs your approval") + `
    <div style="background:#fff3e0;border-left:4px solid #e65100;padding:14px 18px;
      border-radius:0 8px 8px 0;margin-bottom:20px;">
      <strong>⚠️ Action Required:</strong> Please review and approve or reject this travel request.
    </div>

    <div class="info-grid">
      <div class="info-item">
        <div class="key">Travel Plan ID</div>
        <div class="val" style="color:#e65100;font-weight:700;">${id}</div>
      </div>
      <div class="info-item">
        <div class="key">Submitted By</div>
        <div class="val">${user.name}</div>
      </div>
      <div class="info-item">
        <div class="key">Travel Date</div>
        <div class="val">${data.travelDate}</div>
      </div>
      <div class="info-item">
        <div class="key">Return Date</div>
        <div class="val">${data.returnDate || 'TBD'}</div>
      </div>
    </div>

    <table class="data-table">
      <tr><th colspan="2">Travel Details</th></tr>
      <tr><td><strong>Destination City</strong></td><td>${data.city}</td></tr>
      <tr><td><strong>Clients to Visit</strong></td><td>${data.clients || '-'}</td></tr>
      <tr><td><strong>Purpose</strong></td><td>${data.purpose}</td></tr>
      <tr><td><strong>Expected Revenue</strong></td>
          <td style="font-weight:700;color:#2e7d32;">RM ${Number(data.expectedRevenue || 0).toLocaleString()}</td></tr>
      <tr><td><strong>Estimated Duration</strong></td><td>${data.estimatedDays} Day(s)</td></tr>
      <tr><td><strong>Transport</strong></td><td>${data.transport || '-'}</td></tr>
      <tr><td><strong>Accommodation</strong></td><td>${data.accommodation || '-'}</td></tr>
      <tr><td><strong>Status</strong></td>
          <td><span class="status-badge status-pending">Pending Approval</span></td></tr>
    </table>

    <div style="text-align:center;margin:30px 0;">
      <p style="font-size:13px;color:#555;margin-bottom:20px;">
        Click a button below to take action directly from your email.<br>
        The sales representative will be notified automatically.
      </p>
      <a href="${approveUrl}" class="btn btn-success"
        style="display:inline-block;padding:14px 36px;background:#2e7d32;color:#fff;
          text-decoration:none;border-radius:8px;font-size:15px;font-weight:700;
          margin:0 10px;">✅ Approve</a>
      <a href="${rejectUrl}" class="btn btn-danger"
        style="display:inline-block;padding:14px 36px;background:#c62828;color:#fff;
          text-decoration:none;border-radius:8px;font-size:15px;font-weight:700;
          margin:0 10px;">❌ Reject</a>
    </div>
    <p style="font-size:11px;color:#aaa;text-align:center;">
      These links expire in 48 hours and can only be used once.
    </p>
  ` + getEmailFooter();

  const subject = `✈️ Approval Required: ${id} — ${user.name} to ${data.city}`;

  recipients.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, "Please view in HTML.", { htmlBody: html });
    } catch (e) {
      Logger.log("Failed travel email to " + email + ": " + e.message);
    }
  });
}

// ─── TRAVEL STATUS NOTIFICATION ───────────────────────────────

function sendTravelStatusEmail(rowData, status, approver, notes) {
  const repEmail = String(rowData[2]).trim();
  if (!isValidEmail(repEmail)) return;

  const statusIcon = status === "Approved" ? "✅" : "❌";
  const statusClass = status === "Approved" ? "status-approved" : "status-rejected";

  const html = getEmailHeader(
    `Travel Plan ${status}`,
    `Your travel request has been ${status.toLowerCase()} by ${approver.name}`
  ) + `
    <div class="highlight-box">
      <div class="label">Decision</div>
      <div class="value">${statusIcon} ${status}</div>
    </div>

    <table class="data-table">
      <tr><th colspan="2">Travel Plan Details</th></tr>
      <tr><td><strong>Plan ID</strong></td><td>${rowData[0]}</td></tr>
      <tr><td><strong>Destination</strong></td><td>${rowData[5]}</td></tr>
      <tr><td><strong>Travel Date</strong></td><td>${rowData[3]}</td></tr>
      <tr><td><strong>Return Date</strong></td><td>${rowData[4] || 'TBD'}</td></tr>
      <tr><td><strong>Purpose</strong></td><td>${rowData[7]}</td></tr>
      <tr><td><strong>Decision By</strong></td><td>${approver.name}</td></tr>
      <tr><td><strong>Decision Date</strong></td>
          <td>${Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMM yyyy")}</td></tr>
      <tr><td><strong>Status</strong></td>
          <td><span class="status-badge ${statusClass}">${status}</span></td></tr>
    </table>

    ${notes ? `
    <div class="highlight-box">
      <div class="label">📝 Notes from Approver</div>
      <p style="margin-top:8px;font-size:14px;">${notes}</p>
    </div>` : ""}

    ${status === "Rejected" ? `
    <p style="font-size:13px;color:#555;">
      If you have questions about this decision, please contact HR directly.
    </p>` : `
    <p style="font-size:13px;color:#2e7d32;">
      Your travel plan has been approved. Please ensure all travel documents are prepared
      and submit your expense report upon return.
    </p>`}
  ` + getEmailFooter();

  const subject = `${statusIcon} Travel Plan ${rowData[0]}: ${status} | ${rowData[5]}`;

  try {
    GmailApp.sendEmail(repEmail, subject, "Please view in HTML.", { htmlBody: html });
  } catch (e) {
    Logger.log("Travel status email error: " + e.message);
  }
}

// ─── MONTHLY REPORT EMAIL ─────────────────────────────────────

function buildMonthlyReportEmail(stats) {
  const monthName = [
    "", "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ][stats.thisMonth];

  return getEmailHeader(
    `Monthly Summary Report — ${monthName} ${stats.thisYear}`,
    "Management overview of sales performance"
  ) + `
    <div class="section-label">Key Performance Indicators</div>
    <table style="width:100%;border-collapse:collapse;margin:0 0 20px 0;">
      <tr>
        <td style="padding:8px;">
          <div class="highlight-box">
            <div class="label">Monthly Revenue</div>
            <div class="value">RM ${Number(stats.monthlyRevenue || 0).toLocaleString()}</div>
          </div>
        </td>
        <td style="padding:8px;">
          <div class="highlight-box">
            <div class="label">Total Bookings</div>
            <div class="value">${stats.confirmedBookings}</div>
          </div>
        </td>
      </tr>
      <tr>
        <td style="padding:8px;">
          <div class="highlight-box">
            <div class="label">New Leads</div>
            <div class="value">${stats.totalLeads}</div>
          </div>
        </td>
        <td style="padding:8px;">
          <div class="highlight-box">
            <div class="label">Conversion Rate</div>
            <div class="value">${stats.conversionRate}%</div>
          </div>
        </td>
      </tr>
    </table>

    <table class="data-table">
      <tr>
        <th>Metric</th>
        <th>This Month</th>
        <th>Total YTD</th>
      </tr>
      <tr>
        <td>DSR Reports</td>
        <td>${stats.monthlyDSR}</td>
        <td>${stats.totalDSR}</td>
      </tr>
      <tr>
        <td>Won Leads</td>
        <td>${stats.wonLeads}</td>
        <td>-</td>
      </tr>
      <tr>
        <td>Pending Travels</td>
        <td>${stats.pendingTravel}</td>
        <td>-</td>
      </tr>
      <tr>
        <td>Incentives Paid</td>
        <td>-</td>
        <td>RM ${Number(stats.totalIncentives || 0).toLocaleString()}</td>
      </tr>
    </table>

    <p style="font-size:13px;color:#555;">
      Login to the system for the complete detailed report with individual breakdowns.
    </p>
  ` + getEmailFooter();
}

// ─── PDF GENERATION ──────────────────────────────────────────

function generateDSRPdf(id, data, user) {
  var htmlContent = 
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<style>' +
      'body { font-family: Arial, sans-serif; font-size: 12px; color: #333; margin: 40px; }' +
      '.header { text-align: center; border-bottom: 3px solid #1565c0; padding-bottom: 20px; margin-bottom: 25px; }' +
      '.header h1 { font-size: 22px; color: #1a237e; margin: 0; }' +
      '.header p { color: #666; margin: 4px 0 0; }' +
      '.report-id { font-size: 13px; color: #1565c0; font-weight: bold; }' +
      'table { width: 100%; border-collapse: collapse; margin: 15px 0; }' +
      'th { background: #1565c0; color: #fff; padding: 10px; text-align: left; font-size: 11px; }' +
      'td { padding: 9px 10px; border-bottom: 1px solid #eee; }' +
      'td:first-child { font-weight: bold; width: 35%; background: #f8fafc; }' +
      '.footer { text-align: center; margin-top: 30px; font-size: 10px; color: #999;' +
        'border-top: 1px solid #eee; padding-top: 15px; }' +
      '.stamp { display: inline-block; border: 2px solid #2e7d32; color: #2e7d32;' +
        'padding: 6px 20px; border-radius: 4px; font-size: 14px; font-weight: bold;' +
        'letter-spacing: 2px; margin: 15px 0; }' +
    '</style>' +
    '</head>' +
    '<body>' +
      '<div class="header">' +
        '<h1>SandalMist</h1>' +
        '<p>' + COMPANY_TAGLINE + '</p>' +
        '<p style="margin-top:10px;"><strong>DAILY SALES REPORT</strong></p>' +
        '<p class="report-id">' + id + '</p>' +
      '</div>' +

      '<table>' +
        '<tr><th colspan="2">Report Information</th></tr>' +
        '<tr><td>Report ID</td><td>' + id + '</td></tr>' +
        '<tr><td>Submitted By</td><td>' + user.name + '</td></tr>' +
        '<tr><td>Email</td><td>' + user.email + '</td></tr>' +
        '<tr><td>Submission Date</td><td>' + Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMMM yyyy, hh:mm a") + '</td></tr>' +
      '</table>' +

      '<table>' +
        '<tr><th colspan="2">Visit Details</th></tr>' +
        '<tr><td>Visit Date</td><td>' + data.date + '</td></tr>' +
        '<tr><td>Location</td><td>' + data.location + '</td></tr>' +
        '<tr><td>Client Name</td><td>' + data.clientName + '</td></tr>' +
        '<tr><td>Contact Number</td><td>' + (data.contact || '-') + '</td></tr>' +
        '<tr><td>Purpose of Visit</td><td>' + data.purpose + '</td></tr>' +
        '<tr><td>Result / Outcome</td><td>' + data.result + '</td></tr>' +
        '<tr><td>Next Action Required</td><td>' + (data.nextAction || '-') + '</td></tr>' +
      '</table>' +

      '<div style="text-align:center;margin:20px 0;">' +
        '<span class="stamp">SUBMITTED</span>' +
      '</div>' +

      '<div class="footer">' +
        '<p>This document is confidential and intended for internal use only.</p>' +
        '<p>' + COMPANY_FOOTER + ' | Generated: ' + Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMM yyyy HH:mm") + '</p>' +
      '</div>' +
    '</body>' +
    '</html>';

  // Convert HTML to PDF using Google's conversion service
  var blob = Utilities.newBlob(htmlContent, 'text/html', id + '.html');
  var pdfBlob = blob.getAs('application/pdf').setName(id + '_DSR.pdf');
  
  var folder = getSubFolder("DSR_PDFs");
  var file = folder.createFile(pdfBlob);

  return file.getUrl();
}

function saveFileToDrive(base64Data, fileName, folderName) {
  try {
    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, "application/octet-stream", fileName);
    const folder = getSubFolder(folderName);
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    Logger.log("File upload error: " + e.message);
    return "";
  }
}

