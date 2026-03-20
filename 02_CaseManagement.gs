/**
 * ============================================================
 * 02_CaseManagement.gs — Case CRUD Operations
 * ============================================================
 */

/**
 * Create a new case for a client
 */
function createCase(caseData) {
  return withScriptLock_(function() {
    var sheet = getSheet_("CASE");
    var client = getClientById(caseData.clientId);
    if (!client) {
      return { success: false, message: "Client not found: " + caseData.clientId };
    }
    var caseId = generateCaseIdForClient_(client);

    var caseFolderId = "";
    if (client.CLIENT_FOLDER_ID) {
      try {
        var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
        var title = (caseData.patentTitle || "Untitled").substring(0, 30);
        var caseFolder = clientFolder.createFolder(caseId + "_" + title);
        caseFolderId = caseFolder.getId();
        var subfolderNames = ["DRAFTS", "FILINGS", "OFFICE_ACTIONS", "RESPONSES", "MISC"];
        for (var s = 0; s < subfolderNames.length; s++) {
          caseFolder.createFolder(subfolderNames[s]);
        }
      } catch (e) {
        Logger.log("Could not create case folder: " + e.message);
      }
    }

    sheet.appendRow([
      caseId,
      caseData.clientId,
      client.CLIENT_NAME,
      caseData.patentTitle || "",
      caseData.applicationNumber || "",
      caseData.country || "India",
      caseData.filingDate || "",
      caseData.status || "Drafted",
      caseData.nextDeadline || "",
      caseData.attorney || "",
      caseFolderId,
      caseData.priority || "Normal",
      caseData.patentType || "Utility",
      client.ORG_ID || caseData.orgId || "",
      caseData.assignedStaffEmail || client.ASSIGNED_STAFF_EMAIL || "",
      caseData.galvanizerEmail || "",
      caseData.workflowStage || "Drafting",
      caseData.notes || "",
      new Date(),
      new Date()
    ]);

    try {
      logActivity_("CREATE_CASE", "CASE", caseId, "Patent: " + (caseData.patentTitle || ""));
    } catch (e) {
      Logger.log("Logging failed: " + e.message);
    }

    return {
      success: true,
      caseId: caseId,
      message: "Case " + caseId + " created for client " + caseData.clientId
    };
  });
}

/**
 * Get cases by CLIENT_ID (for client portal)
 */
function getClientCases(clientId) {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var cases = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === clientId) {
      var caseObj = {};
      for (var j = 0; j < headers.length; j++) {
        caseObj[headers[j]] = data[i][j];
      }
      cases.push(caseObj);
    }
  }
  return cases;
}

/**
 * Get cases by attorney email
 */
function getAttorneyCases(attorneyEmail) {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var cases = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === attorneyEmail) {
      var caseObj = {};
      for (var j = 0; j < headers.length; j++) {
        caseObj[headers[j]] = data[i][j];
      }
      cases.push(caseObj);
    }
  }
  return cases;
}

/**
 * Get all cases (admin view)
 */
function getAllCases() {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var cases = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var caseObj = {};
      for (var j = 0; j < headers.length; j++) {
        caseObj[headers[j]] = data[i][j];
      }
      cases.push(caseObj);
    }
  }
  return cases;
}

/**
 * Update case status/details
 */
function updateCase(caseId, updates) {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === caseId) {
      var keys = Object.keys(updates);
      for (var k = 0; k < keys.length; k++) {
        var colIdx = headers.indexOf(keys[k]);
        if (colIdx > -1) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updates[keys[k]]);
        }
      }
      // Update LAST_UPDATED
      var lastUpdatedIdx = headers.indexOf("LAST_UPDATED");
      if (lastUpdatedIdx > -1) {
        sheet.getRange(i + 1, lastUpdatedIdx + 1).setValue(new Date());
      }
      try { logActivity_("UPDATE_CASE", "CASE", caseId, JSON.stringify(updates)); } catch(e) {}
      return { success: true };
    }
  }
  return { success: false, message: "Case not found" };
}

/**
 * Get upcoming deadlines (within N days)
 */
function getUpcomingDeadlines(daysAhead) {
  daysAhead = daysAhead || 10;
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var deadlines = [];
  var now = new Date();
  var cutoff = new Date(now.getTime() + daysAhead * 86400000);

  for (var i = 1; i < data.length; i++) {
    var deadline = data[i][8]; // NEXT_DEADLINE
    if (deadline && deadline instanceof Date && deadline >= now && deadline <= cutoff) {
      var caseObj = {};
      for (var j = 0; j < headers.length; j++) {
        caseObj[headers[j]] = data[i][j];
      }
      deadlines.push(caseObj);
    }
  }

  deadlines.sort(function(a, b) {
    return new Date(a.NEXT_DEADLINE) - new Date(b.NEXT_DEADLINE);
  });

  return deadlines;
}

// ── UI Dialog for Creating Case ────────────────────────────
function showCreateCaseDialog() {
  var clients = getAllClients();
  var clientOptions = clients.map(function(c) {
    return '<option value="' + c.CLIENT_ID + '">' + c.CLIENT_ID + ' - ' + c.CLIENT_NAME + '</option>';
  }).join("");

  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; }
      h2 { color: #1a237e; }
      label { display: block; margin-top: 10px; font-weight: 500; }
      input, select, textarea { width: 100%; padding: 8px 12px; margin-top: 4px; border: 1px solid #ddd;
        border-radius: 6px; font-size: 14px; box-sizing: border-box; }
      .btn { background: #1a237e; color: white; padding: 10px 24px; border: none;
        border-radius: 6px; cursor: pointer; margin-top: 16px; }
      .btn:hover { background: #283593; }
      .btn-cancel { background: #757575; margin-left: 8px; }
      .row { display: flex; gap: 12px; }
      .row > div { flex: 1; }
      .required { color: #d32f2f; }
      #status { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }
    </style>
    <h2>📋 Create New Case</h2>
    <label>Client <span class="required">*</span></label>
    <select id="clientId">${clientOptions}</select>
    <label>Patent Title <span class="required">*</span></label>
    <input id="patentTitle" placeholder="e.g. Method for AI-Based Drug Discovery" />
    <div class="row">
      <div>
        <label>Application Number</label>
        <input id="applicationNumber" placeholder="e.g. IN202411001234" />
      </div>
      <div>
        <label>Country</label>
        <select id="country">
          <option>India</option><option>USA</option><option>EPO</option>
          <option>China</option><option>Japan</option><option>PCT</option>
          <option>Other</option>
        </select>
      </div>
    </div>
    <div class="row">
      <div>
        <label>Filing Date</label>
        <input id="filingDate" type="date" />
      </div>
      <div>
        <label>Next Deadline</label>
        <input id="nextDeadline" type="date" />
      </div>
    </div>
    <div class="row">
      <div>
        <label>Status</label>
        <select id="status_sel">
          <option>Drafted</option><option>Filed</option><option>Published</option>
          <option>Under Examination</option><option>Granted</option>
          <option>Refused</option><option>Abandoned</option><option>Lapsed</option>
        </select>
      </div>
      <div>
        <label>Patent Type</label>
        <select id="patentType">
          <option>Utility</option><option>Design</option><option>Plant</option>
          <option>Provisional</option><option>PCT</option>
        </select>
      </div>
    </div>
    <label>Attorney Email</label>
    <input id="attorney" type="email" placeholder="attorney@firm.com" />
    <label>Priority</label>
    <select id="priority">
      <option>Normal</option><option>High</option><option>Urgent</option>
    </select>
    <label>Notes</label>
    <textarea id="notes" rows="2"></textarea>
    <div>
      <button class="btn" id="submitBtn" onclick="submitCase()">Create Case</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    <div id="status"></div>
    <script>
      function submitCase() {
        var data = {
          clientId: document.getElementById('clientId').value,
          patentTitle: document.getElementById('patentTitle').value,
          applicationNumber: document.getElementById('applicationNumber').value,
          country: document.getElementById('country').value,
          filingDate: document.getElementById('filingDate').value,
          nextDeadline: document.getElementById('nextDeadline').value,
          status: document.getElementById('status_sel').value,
          patentType: document.getElementById('patentType').value,
          attorney: document.getElementById('attorney').value,
          priority: document.getElementById('priority').value,
          notes: document.getElementById('notes').value
        };
        if (!data.clientId || !data.patentTitle) {
          showMsg('Please fill required fields.', 'error'); return;
        }
        // Disable button to prevent double-click
        document.getElementById('submitBtn').disabled = true;
        document.getElementById('submitBtn').textContent = 'Creating...';
        showMsg('Creating case... This may take a moment.', 'info');
        google.script.run
          .withSuccessHandler(function(r) {
            showMsg(r.message, r.success ? 'success' : 'error');
            if (r.success) {
              setTimeout(function(){ google.script.host.close(); }, 2000);
            } else {
              document.getElementById('submitBtn').disabled = false;
              document.getElementById('submitBtn').textContent = 'Create Case';
            }
          })
          .withFailureHandler(function(e) {
            showMsg('Error: ' + e.message, 'error');
            document.getElementById('submitBtn').disabled = false;
            document.getElementById('submitBtn').textContent = 'Create Case';
          })
          .createCase(data);
      }
      function showMsg(msg, type) {
        var el = document.getElementById('status');
        el.style.display = 'block';
        el.style.background = type==='error'?'#ffebee':type==='success'?'#e8f5e9':'#e3f2fd';
        el.style.color = type==='error'?'#c62828':type==='success'?'#2e7d32':'#1565c0';
        el.textContent = msg;
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "Create New Case");
}
