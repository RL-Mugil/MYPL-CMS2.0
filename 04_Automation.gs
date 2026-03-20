/**
 * ============================================================
 * 04_Automation.gs — Deadline Alerts, Analytics, Triggers
 * ============================================================
 */

/**
 * Send deadline alerts for cases due within 10 days
 * Set up a daily time-driven trigger for this function.
 */
function sendDeadlineAlerts() {
  var deadlines = getUpcomingDeadlines(10);

  if (deadlines.length === 0) {
    Logger.log("No upcoming deadlines.");
    return { sent: 0 };
  }

  var sentCount = 0;

  deadlines.forEach(function(caseObj) {
    var daysLeft = Math.ceil(
      (new Date(caseObj.NEXT_DEADLINE).getTime() - new Date().getTime()) / 86400000
    );

    var subject = "⏰ Deadline Alert: " + caseObj.PATENT_TITLE + " (" + daysLeft + " days left)";
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: ${daysLeft <= 3 ? '#d32f2f' : '#ff8f00'}; color: white; padding: 16px; text-align: center; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0;">⏰ Deadline in ${daysLeft} Day${daysLeft > 1 ? 's' : ''}</h2>
        </div>
        <div style="padding: 20px; background: #fff; border: 1px solid #eee; border-radius: 0 0 8px 8px;">
          <table style="width: 100%;">
            <tr><td style="padding: 6px 0; font-weight: bold; width: 140px;">Case ID:</td><td>${caseObj.CASE_ID}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Patent Title:</td><td>${caseObj.PATENT_TITLE}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Client:</td><td>${caseObj.CLIENT_NAME}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Country:</td><td>${caseObj.COUNTRY}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Status:</td><td>${caseObj.CURRENT_STATUS}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Deadline:</td><td style="color: #d32f2f; font-weight: bold;">${Utilities.formatDate(new Date(caseObj.NEXT_DEADLINE), Session.getScriptTimeZone(), "dd-MMM-yyyy")}</td></tr>
            <tr><td style="padding: 6px 0; font-weight: bold;">Attorney:</td><td>${caseObj.ATTORNEY}</td></tr>
          </table>
          <p style="margin-top: 16px; color: #666; font-size: 12px;">This is an automated reminder from the IP Patent Firm Case Management System.</p>
        </div>
      </div>
    `;

    // Send to attorney
    if (caseObj.ATTORNEY) {
      try {
        MailApp.sendEmail({ to: caseObj.ATTORNEY, subject: subject, htmlBody: htmlBody });
        sentCount++;
      } catch (e) { Logger.log("Failed to email attorney: " + e.message); }
    }

    // Send to client
    var client = getClientById(caseObj.CLIENT_ID);
    if (client && client.EMAIL) {
      try {
        MailApp.sendEmail({ to: client.EMAIL, subject: subject, htmlBody: htmlBody });
        sentCount++;
      } catch (e) { Logger.log("Failed to email client: " + e.message); }
    }
  });

  logActivity_("DEADLINE_ALERTS", "SYSTEM", "", sentCount + " alerts sent for " + deadlines.length + " cases");

  Logger.log("✅ Sent " + sentCount + " deadline alerts for " + deadlines.length + " cases");
  return { sent: sentCount, cases: deadlines.length };
}

/**
 * Get analytics data for admin dashboard
 */
function getAnalytics() {
  var clients = getAllClients();
  var cases = getAllCases();
  var invoiceSheet = getSheet_("INVOICE");
  var invoiceData = invoiceSheet.getDataRange().getValues();

  var now = new Date();
  var thisYear = now.getFullYear();
  var thisMonth = now.getMonth();

  // Client metrics
  var totalClients = clients.length;
  var activeClients = clients.filter(function(c) { return c.STATUS === "Active"; }).length;

  // Case metrics
  var totalCases = cases.length;
  var activeCases = cases.filter(function(c) {
    return ["Filed", "Under Examination", "Published", "Drafted"].indexOf(c.CURRENT_STATUS) > -1;
  }).length;

  var casesThisYear = cases.filter(function(c) {
    return c.CREATED_DATE && new Date(c.CREATED_DATE).getFullYear() === thisYear;
  }).length;

  var pendingDeadlines = getUpcomingDeadlines(30).length;

  // Case status distribution
  var statusDist = {};
  cases.forEach(function(c) {
    var s = c.CURRENT_STATUS || "Unknown";
    statusDist[s] = (statusDist[s] || 0) + 1;
  });

  // Cases by country
  var countryDist = {};
  cases.forEach(function(c) {
    var co = c.COUNTRY || "Unknown";
    countryDist[co] = (countryDist[co] || 0) + 1;
  });

  // Invoice/Revenue metrics
  var totalRevenue = 0;
  var revenueThisMonth = 0;
  var pendingPayments = 0;
  var pendingAmount = 0;
  var revenueByClient = {};

  for (var i = 1; i < invoiceData.length; i++) {
    if (!invoiceData[i][0]) continue;
    var total = parseFloat(invoiceData[i][11]) || 0; // TOTAL
    var status = invoiceData[i][12]; // PAYMENT_STATUS
    var invDate = invoiceData[i][4]; // INVOICE_DATE
    var clientName = invoiceData[i][2]; // CLIENT_NAME

    if (status === "Paid") {
      totalRevenue += total;
      if (invDate instanceof Date && invDate.getFullYear() === thisYear && invDate.getMonth() === thisMonth) {
        revenueThisMonth += total;
      }
    } else {
      pendingPayments++;
      pendingAmount += total;
    }

    revenueByClient[clientName] = (revenueByClient[clientName] || 0) + total;
  }

  return {
    totalClients: totalClients,
    activeClients: activeClients,
    totalCases: totalCases,
    activeCases: activeCases,
    casesThisYear: casesThisYear,
    pendingDeadlines: pendingDeadlines,
    totalRevenue: totalRevenue,
    revenueThisMonth: revenueThisMonth,
    pendingPayments: pendingPayments,
    pendingAmount: pendingAmount,
    statusDistribution: statusDist,
    countryDistribution: countryDist,
    revenueByClient: revenueByClient
  };
}

/**
 * Show analytics dashboard dialog
 */
function showAnalyticsDashboard() {
  var analytics = getAnalytics();
  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; background: #f5f5f5; }
      h2 { color: #1a237e; text-align: center; }
      .grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin: 16px 0; }
      .card { background: white; border-radius: 10px; padding: 16px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
      .card .value { font-size: 28px; font-weight: bold; color: #1a237e; }
      .card .label { font-size: 12px; color: #666; margin-top: 4px; }
      .card.alert { border-left: 4px solid #d32f2f; }
      .card.alert .value { color: #d32f2f; }
      .card.success { border-left: 4px solid #2e7d32; }
      .card.success .value { color: #2e7d32; }
      .section { background: white; border-radius: 10px; padding: 16px; margin: 12px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
      .section h3 { color: #1a237e; margin: 0 0 12px 0; font-size: 14px; }
      .bar { display: flex; align-items: center; margin: 6px 0; }
      .bar-label { width: 120px; font-size: 12px; color: #333; }
      .bar-fill { height: 20px; background: #3f51b5; border-radius: 4px; min-width: 4px; transition: width 0.3s; }
      .bar-value { margin-left: 8px; font-size: 12px; font-weight: bold; color: #333; }
    </style>
    <h2>📊 Analytics Dashboard</h2>
    <div class="grid">
      <div class="card"><div class="value">${analytics.totalClients}</div><div class="label">Total Clients</div></div>
      <div class="card success"><div class="value">${analytics.activeCases}</div><div class="label">Active Cases</div></div>
      <div class="card"><div class="value">${analytics.casesThisYear}</div><div class="label">Filed This Year</div></div>
      <div class="card success"><div class="value">₹${(analytics.revenueThisMonth/1000).toFixed(1)}K</div><div class="label">Revenue This Month</div></div>
      <div class="card alert"><div class="value">${analytics.pendingDeadlines}</div><div class="label">Upcoming Deadlines</div></div>
      <div class="card alert"><div class="value">${analytics.pendingPayments}</div><div class="label">Unpaid Invoices (₹${(analytics.pendingAmount/1000).toFixed(1)}K)</div></div>
    </div>
    <div class="section">
      <h3>📋 Cases by Status</h3>
      ${Object.keys(analytics.statusDistribution).map(function(s) {
        var count = analytics.statusDistribution[s];
        var pct = Math.round((count / Math.max(analytics.totalCases, 1)) * 100);
        return '<div class="bar"><span class="bar-label">'+s+'</span><div class="bar-fill" style="width:'+pct+'%"></div><span class="bar-value">'+count+'</span></div>';
      }).join("")}
    </div>
    <div class="section">
      <h3>🌍 Cases by Country</h3>
      ${Object.keys(analytics.countryDistribution).map(function(c) {
        var count = analytics.countryDistribution[c];
        var pct = Math.round((count / Math.max(analytics.totalCases, 1)) * 100);
        return '<div class="bar"><span class="bar-label">'+c+'</span><div class="bar-fill" style="width:'+pct+'%; background: #00897b;"></div><span class="bar-value">'+count+'</span></div>';
      }).join("")}
    </div>
  `).setWidth(560).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Analytics Dashboard");
}

/**
 * Set up daily deadline alert trigger
 */
function setupDailyTrigger() {
  // Delete existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "sendDeadlineAlerts") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new daily trigger at 9 AM
  ScriptApp.newTrigger("sendDeadlineAlerts")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log("✅ Daily deadline alert trigger created (9 AM)");
}

/**
 * Deploy portal info
 */
function deployPortalInfo() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "🌐 Client Portal Deployment",
    "To deploy the client portal:\n\n" +
    "1. Open Apps Script editor (Extensions > Apps Script)\n" +
    "2. Make sure 05_ClientPortal.gs is included\n" +
    "3. Go to Deploy > New deployment\n" +
    "4. Select 'Web app'\n" +
    "5. Execute as: 'User accessing the web app'\n" +
    "6. Who has access: 'Anyone with Google account'\n" +
    "7. Click Deploy\n" +
    "8. Share the URL with your clients\n\n" +
    "Each client will only see their own data based on their Google login email.",
    ui.ButtonSet.OK
  );
}
