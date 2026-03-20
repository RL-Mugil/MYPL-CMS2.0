/**
 * ============================================================
 * 01_ClientManagement.gs — Client CRUD Operations
 * ============================================================
 */

/**
 * Create a new client with folder structure and portal access
 */
function createClient(clientData) {
  return withScriptLock_(function() {
    var sheet = getSheet_("MASTER_CLIENT");
    var clientRegion = clientData.clientRegion || (String(clientData.isAbroad || "").toLowerCase() === "true" ? "Abroad" : "India");
    var clientCode = normalizeClientCode_(clientData.clientCode || clientData.clientId || "", clientRegion);
    if (!looksLikeModernClientCode_(clientCode)) {
      return { success: false, message: "Valid client code is required, for example A61M or 870Y." };
    }
    var clientId = clientCode;
    if (getClientById(clientId)) {
      return { success: false, message: "Client code already exists: " + clientId };
    }
    var clientFolders = createClientFolderStructure_(clientId, clientData.clientName);
    var clientType = clientData.clientType || (clientData.orgId ? "Organization" : "Individual");

    sheet.appendRow([
      clientId,
      clientData.clientName,
      clientData.contactPerson,
      clientData.email,
      clientData.phone || "",
      clientData.address || "",
      new Date(),
      "Active",
      clientFolders.rootId,
      "Enabled",
      clientType,
      clientCode,
      clientRegion,
      clientData.orgId || "",
      clientData.clientAdminUserId || "",
      clientData.assignedStaffEmail || "",
      clientData.notes || ""
    ]);

    if (clientData.email) {
      try {
        DriveApp.getFolderById(clientFolders.rootId)
          .addEditor(clientData.email);
      } catch (e) {
        Logger.log("Could not share folder with " + clientData.email + ": " + e.message);
      }
    }

    addClientUser_(clientId, clientData.email, clientData.contactPerson, {
      orgId: clientData.orgId || "",
      role: clientType === "Organization" ? "Client Admin" : "Individual Client"
    });
    logActivity_("CREATE_CLIENT", "CLIENT", clientId, "Created client: " + clientData.clientName);

    return {
      success: true,
      clientId: clientId,
      folderId: clientFolders.rootId,
      message: "Client " + clientId + " created successfully!"
    };
  });
}

/**
 * Create client folder structure in Drive
 */
function createClientFolderStructure_(clientId, clientName) {
  var parentFolder = getFolder_("02_CLIENT_FOLDERS");
  var folderName = clientId + "_" + clientName.replace(/[^a-zA-Z0-9 ]/g, "");
  var existingFolder = parentFolder.getFoldersByName(folderName);
  var clientFolder = existingFolder.hasNext() ? existingFolder.next() : parentFolder.createFolder(folderName);
  var subIds = {};

  CONFIG.CLIENT_SUBFOLDERS.forEach(function(name) {
    var subIter = clientFolder.getFoldersByName(name);
    var sub = subIter.hasNext() ? subIter.next() : clientFolder.createFolder(name);
    subIds[name] = sub.getId();
  });

  return {
    rootId: clientFolder.getId(),
    subfolders: subIds
  };
}

/**
 * Add client to USER_ROLES sheet with default password
 */
function addClientUser_(clientId, email, name, options) {
  var sheet = getSheet_("USERS");
  var userId = generateId_("USR", sheet, 0);
  options = options || {};
  
  sheet.appendRow([
    userId,
    email,
    name,
    options.role || "Individual Client",
    clientId,
    options.orgId || "",
    "External",
    "Active",
    "",
    "No",
    "",
    "",
    new Date()
  ]);
}

/**
 * Get all clients (for admin view)
 */
function getAllClients() {
  var sheet = getSheet_("MASTER_CLIENT");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var clients = [];

  for (var i = 1; i < data.length; i++) {
    var client = {};
    for (var j = 0; j < headers.length; j++) {
      client[headers[j]] = data[i][j];
    }
    clients.push(client);
  }
  return clients;
}

/**
 * Get client by email (for portal login)
 */
function getClientByEmail(email) {
  var sheet = getSheet_("MASTER_CLIENT");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === email) {
      var client = {};
      for (var j = 0; j < headers.length; j++) {
        client[headers[j]] = data[i][j];
      }
      return client;
    }
  }
  return null;
}

/**
 * Get client by CLIENT_ID
 */
function getClientById(clientId) {
  var sheet = getSheet_("MASTER_CLIENT");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) {
      var client = {};
      for (var j = 0; j < headers.length; j++) {
        client[headers[j]] = data[i][j];
      }
      return client;
    }
  }
  return null;
}

/**
 * Update client details
 */
function updateClient(clientId, updates) {
  var sheet = getSheet_("MASTER_CLIENT");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) {
      var keys = Object.keys(updates);
      for (var k = 0; k < keys.length; k++) {
        var colIdx = headers.indexOf(keys[k]);
        if (colIdx > -1) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updates[keys[k]]);
        }
      }
      logActivity_("UPDATE_CLIENT", "CLIENT", clientId, JSON.stringify(updates));
      return { success: true };
    }
  }
  return { success: false, message: "Client not found" };
}

// ── UI Dialog for Creating Client ──────────────────────────
function showCreateClientDialog() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; }
      h2 { color: #1a237e; margin-bottom: 16px; }
      label { display: block; margin-top: 12px; font-weight: 500; color: #333; }
      input, textarea { width: 100%; padding: 8px 12px; margin-top: 4px; border: 1px solid #ddd;
        border-radius: 6px; font-size: 14px; box-sizing: border-box; }
      input:focus, textarea:focus { border-color: #1a237e; outline: none; box-shadow: 0 0 0 2px rgba(26,35,126,0.1); }
      .btn { background: #1a237e; color: white; padding: 10px 24px; border: none;
        border-radius: 6px; cursor: pointer; font-size: 14px; margin-top: 20px; }
      .btn:hover { background: #283593; }
      .btn-cancel { background: #757575; margin-left: 8px; }
      .required { color: #d32f2f; }
      #status { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }
    </style>
    <h2>➕ Create New Client</h2>
    <label>Client Name <span class="required">*</span></label>
    <input id="clientName" placeholder="e.g. ABC Technologies Pvt Ltd" />
    <label>Contact Person <span class="required">*</span></label>
    <input id="contactPerson" placeholder="e.g. John Doe" />
    <label>Email <span class="required">*</span></label>
    <input id="email" type="email" placeholder="e.g. john@abctech.com" />
    <label>Phone</label>
    <input id="phone" placeholder="e.g. +91-9876543210" />
    <label>Address</label>
    <textarea id="address" rows="2" placeholder="Full address"></textarea>
    <label>Notes</label>
    <textarea id="notes" rows="2" placeholder="Any additional notes"></textarea>
    <div>
      <button class="btn" onclick="submitClient()">Create Client</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    <div id="status"></div>
    <script>
      function submitClient() {
        var data = {
          clientName: document.getElementById('clientName').value,
          contactPerson: document.getElementById('contactPerson').value,
          email: document.getElementById('email').value,
          phone: document.getElementById('phone').value,
          address: document.getElementById('address').value,
          notes: document.getElementById('notes').value
        };
        if (!data.clientName || !data.contactPerson || !data.email) {
          showStatus('Please fill all required fields.', 'error');
          return;
        }
        showStatus('Creating client...', 'info');
        google.script.run
          .withSuccessHandler(function(result) {
            showStatus(result.message, 'success');
            setTimeout(function() { google.script.host.close(); }, 3000);
          })
          .withFailureHandler(function(err) {
            showStatus('Error: ' + err.message, 'error');
          })
          .createClient(data);
      }
      function showStatus(msg, type) {
        var el = document.getElementById('status');
        el.style.display = 'block';
        el.style.background = type === 'error' ? '#ffebee' : type === 'success' ? '#e8f5e9' : '#e3f2fd';
        el.style.color = type === 'error' ? '#c62828' : type === 'success' ? '#2e7d32' : '#1565c0';
        el.textContent = msg;
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, "Create New Client");
}
