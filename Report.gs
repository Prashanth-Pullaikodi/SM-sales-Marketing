/**
 * ============================================================
 * Reports.gs - Advanced Reporting & Analytics
 * ============================================================
 */

/**
 * Generate a full team performance report as an HTML string
 * (used by export and monthly email)
 */
function generateTeamPerformanceReport(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Collect data ──────────────────────────────────────────
  const bookings = ss.getSheetByName(SHEETS.BOOKINGS).getDataRange().getValues();
  const leads    = ss.getSheetByName(SHEETS.LEADS).getDataRange().getValues();
  const dsr      = ss.getSheetByName(SHEETS.DAILY_REPORTS).getDataRange().getValues();
  const travel   = ss.getSheetByName(SHEETS.TRAVEL_PLANS).getDataRange().getValues();

  const monthInt = parseInt(month);
  const yearInt  = parseInt(year);

  const performance = {}; // key = salesRep name

  function getOrCreate(rep) {
    if (!performance[rep]) {
      performance[rep] = {
        rep, bookings: 0, revenue: 0, commission: 0,
        newLeads: 0, wonLeads: 0, lostLeads: 0,
        dsrCount: 0, travelCount: 0, roi: 0
      };
    }
    return performance[rep];
  }

  // Bookings
  bookings.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[1]);
    if (isNaN(d)) return;
    if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) return;
    const p = getOrCreate(String(row[15]).trim());
    p.bookings++;
    p.revenue    += parseFloat(row[12]) || 0;
    p.commission += parseFloat(row[13]) || 0;
  });

  // Leads
  leads.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[1]);
    if (isNaN(d)) return;
    if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) return;
    const p = getOrCreate(String(row[9]).trim());
    p.newLeads++;
    if (row[10] === "Won")  p.wonLeads++;
    if (row[10] === "Lost") p.lostLeads++;
  });

  // DSR
  dsr.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[1]);
    if (isNaN(d)) return;
    if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) return;
    getOrCreate(String(row[2]).trim()).dsrCount++;
  });

  // Travel
  travel.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[3]);
    if (isNaN(d)) return;
    if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) return;
    if (row[12] === "Approved") getOrCreate(String(row[1]).trim()).travelCount++;
  });

  // Compute conversion & totals
  const rows = Object.values(performance);
  let grandRevenue = 0, grandBookings = 0, grandLeads = 0;

  rows.forEach(p => {
    p.conversionRate = p.newLeads > 0
      ? Math.round((p.wonLeads / p.newLeads) * 100) : 0;
    grandRevenue   += p.revenue;
    grandBookings  += p.bookings;
    grandLeads     += p.newLeads;
  });

  rows.sort((a, b) => b.revenue - a.revenue);

  const monthName = [
    "", "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ][monthInt];

  // ── Build HTML report ─────────────────────────────────────
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body { font-family:'Segoe UI',Arial,sans-serif; background:#f5f7fa;
         margin:0; padding:0; color:#333; }
  .container { max-width:1000px; margin:0 auto; padding:20px; }
  .report-header { background:linear-gradient(135deg,#1a237e,#1565c0);
    color:#fff; padding:30px 35px; border-radius:12px; margin-bottom:24px; }
  .report-header h1 { font-size:24px; margin:0 0 6px; }
  .report-header p  { font-size:14px; opacity:0.85; margin:0; }
  .kpi-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin-bottom:24px; }
  .kpi-card { background:#fff; border-radius:10px; padding:20px 22px;
    box-shadow:0 2px 8px rgba(0,0,0,0.08); }
  .kpi-card .label { font-size:12px; color:#888; text-transform:uppercase;
    letter-spacing:0.5px; margin-bottom:6px; }
  .kpi-card .value { font-size:26px; font-weight:700; color:#1565c0; }
  .kpi-card .sub { font-size:12px; color:#aaa; margin-top:4px; }
  .section { background:#fff; border-radius:10px; padding:22px 25px;
    box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; }
  .section h3 { font-size:16px; color:#1a237e; margin:0 0 16px;
    padding-bottom:10px; border-bottom:2px solid #e8f0fe; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th { background:#f0f4ff; color:#1a237e; padding:11px 14px;
    text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:0.4px; }
  td { padding:11px 14px; border-bottom:1px solid #f0f0f0; color:#444; }
  tr:hover td { background:#f8faff; }
  tr:last-child td { border-bottom:none; }
  .rank { display:inline-flex; align-items:center; justify-content:center;
    width:24px; height:24px; border-radius:50%; font-size:12px;
    font-weight:700; color:#fff; }
  .rank-1 { background:#f4b400; }
  .rank-2 { background:#9aa0a6; }
  .rank-3 { background:#c6893a; }
  .rank-n { background:#e8eaed; color:#666; }
  .revenue-bar { background:#e8f0fe; border-radius:4px; height:8px;
    position:relative; overflow:hidden; margin-top:4px; }
  .revenue-fill { height:100%; background:linear-gradient(90deg,#1565c0,#42a5f5);
    border-radius:4px; }
  .badge { display:inline-block; padding:3px 9px; border-radius:12px;
    font-size:11px; font-weight:600; }
  .badge-high { background:#e8f5e9; color:#2e7d32; }
  .badge-mid  { background:#fff3e0; color:#e65100; }
  .badge-low  { background:#fce4ec; color:#c62828; }
  .footer-note { text-align:center; font-size:11px; color:#999; margin-top:20px; }
</style>
</head>
<body>
<div class="container">
  <div class="report-header">
    <h1>🏨 ${COMPANY_NAME} — Team Performance Report</h1>
    <p>Period: ${monthName} ${yearInt} &nbsp;|&nbsp; Generated: ${Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd MMM yyyy, hh:mm a")}</p>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="label">Total Revenue</div>
      <div class="value">RM ${formatNumber(grandRevenue)}</div>
      <div class="sub">All confirmed bookings</div>
    </div>
    <div class="kpi-card">
      <div class="label">Total Bookings</div>
      <div class="value">${grandBookings}</div>
      <div class="sub">Confirmed this month</div>
    </div>
    <div class="kpi-card">
      <div class="label">Total Leads</div>
      <div class="value">${grandLeads}</div>
      <div class="sub">New leads this month</div>
    </div>
  </div>

  <div class="section">
    <h3>📊 Sales Team Performance Rankings</h3>
    <table>
      <tr>
        <th>Rank</th>
        <th>Sales Rep</th>
        <th>Revenue (RM)</th>
        <th>Bookings</th>
        <th>New Leads</th>
        <th>Won Leads</th>
        <th>Conversion</th>
        <th>DSR Count</th>
      </tr>
      ${rows.map((p, i) => {
        const rankClass = i === 0 ? "rank-1" : i === 1 ? "rank-2" : i === 2 ? "rank-3" : "rank-n";
        const maxRev = rows[0].revenue || 1;
        const barWidth = Math.round((p.revenue / maxRev) * 100);
        const convBadge = p.conversionRate >= 50 ? "badge-high" :
                          p.conversionRate >= 25 ? "badge-mid" : "badge-low";
        return `
      <tr>
        <td><span class="rank ${rankClass}">${i+1}</span></td>
        <td><strong>${p.rep}</strong></td>
        <td>
          <div>RM ${formatNumber(p.revenue)}</div>
          <div class="revenue-bar">
            <div class="revenue-fill" style="width:${barWidth}%"></div>
          </div>
        </td>
        <td>${p.bookings}</td>
        <td>${p.newLeads}</td>
        <td>${p.wonLeads}</td>
        <td><span class="badge ${convBadge}">${p.conversionRate}%</span></td>
        <td>${p.dsrCount}</td>
      </tr>`;
      }).join("")}
    </table>
  </div>

  <div class="section">
    <h3>💰 Commission Summary</h3>
    <table>
      <tr>
        <th>Sales Rep</th>
        <th>Revenue (RM)</th>
        <th>Commission (10%)</th>
        <th>Incentive Eligible</th>
      </tr>
      ${rows.map(p => {
        const incentive = Math.max(0, p.revenue - 500000) * 0.01;
        return `
      <tr>
        <td>${p.rep}</td>
        <td>RM ${formatNumber(p.revenue)}</td>
        <td>RM ${formatNumber(p.commission)}</td>
        <td>RM ${formatNumber(incentive)}</td>
      </tr>`;
      }).join("")}
      <tr style="font-weight:700;background:#f0f4ff;">
        <td>TOTAL</td>
        <td>RM ${formatNumber(grandRevenue)}</td>
        <td>RM ${formatNumber(rows.reduce((s,p)=>s+p.commission,0))}</td>
        <td>RM ${formatNumber(rows.reduce((s,p)=>s+Math.max(0,p.revenue-500000)*0.01,0))}</td>
      </tr>
    </table>
  </div>

  <p class="footer-note">
    ${COMPANY_FOOTER} &nbsp;|&nbsp; ${monthName} ${yearInt} Performance Report
  </p>
</div>
</body>
</html>`;

  return html;
}

/**
 * Format number with commas
 */
function formatNumber(n) {
  return Number(n || 0).toLocaleString("en-MY", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
}

/**
 * Save performance report to Drive and return URL
 */
function saveReportToDrive(htmlContent, filename) {
  const folder = getSubFolder("Reports");
  const blob = Utilities.newBlob(htmlContent, "text/html", filename);
  const file = folder.createFile(blob);
  return file.getUrl();
}

/**
 * Scheduled monthly incentive calculation
 */
function calculateMonthlyIncentives(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookings = ss.getSheetByName(SHEETS.BOOKINGS).getDataRange().getValues();
  const incentives = ss.getSheetByName(SHEETS.INCENTIVES);

  const monthInt = parseInt(typeof month === "number" ? month :
    ["January","February","March","April","May","June",
     "July","August","September","October","November","December"
    ].indexOf(month) + 1);

  const yearInt = parseInt(year);
  const BASE = 500000;
  const RATE = 0.01;

  const salesTotals = {};

  bookings.slice(1).forEach(row => {
    if (!row[0]) return;
    if (row[16] !== "Confirmed" && row[16] !== "Completed") return;
    const d = new Date(row[1]);
    if (d.getMonth() + 1 !== monthInt || d.getFullYear() !== yearInt) return;
    const rep = String(row[15]).trim();
    const email = String(row[4]).trim(); // this is client email, we need rep email
    salesTotals[rep] = (salesTotals[rep] || 0) + (parseFloat(row[12]) || 0);
  });

  const monthName = [
    "","January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ][monthInt];

  Object.keys(salesTotals).forEach(rep => {
    const total = salesTotals[rep];
    const eligible = Math.max(0, total - BASE);
    const amount = eligible * RATE;

    const repUser = getUserByName(rep);
    incentives.appendRow([
      "INC-" + yearInt + monthInt + rep.replace(/\s+/g,""),
      rep,
      repUser ? repUser.email : "",
      monthName,
      yearInt,
      total, BASE, eligible, RATE, amount,
      amount > 0 ? "Pending Payment" : "Not Eligible",
      "",
      Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss")
    ]);
  });
}

