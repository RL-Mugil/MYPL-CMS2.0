/**
 * ============================================================
 * IP PATENT FIRM - CLIENT PORTFOLIO & CASE MANAGEMENT SYSTEM
 * 00_Setup.gs — Master Setup Script
 * ============================================================
 * Run setupEntireSystem() ONCE to create all folders, sheets,
 * and initial structure automatically.
 * ============================================================
 */

// ── Global Config ──────────────────────────────────────────
var CONFIG = {
  DEFAULT_ROOT_FOLDER_NAME: "IP_PATENT_FIRM_SYSTEM",
  SUBFOLDERS: [
    "01_DATABASE",
    "02_CLIENT_FOLDERS",
    "03_CLIENT_PORTAL",
    "04_INVOICE_SYSTEM",
    "05_CASE_DATABASE",
    "06_AUTOMATION"
  ],
  SHEETS: {
    MASTER_CLIENT: {
      name: "MASTER_CLIENT_DATABASE",
      folder: "01_DATABASE",
      headers: [
        "CLIENT_ID", "CLIENT_NAME", "CONTACT_PERSON", "EMAIL", "PHONE",
        "ADDRESS", "DATE_ONBOARDED", "STATUS", "CLIENT_FOLDER_ID",
        "PORTAL_ACCESS", "CLIENT_TYPE", "CLIENT_CODE", "CLIENT_REGION",
        "ORG_ID", "CLIENT_ADMIN_USER_ID", "ASSIGNED_STAFF_EMAIL", "NOTES"
      ]
    },
    ORGANIZATIONS: {
      name: "ORGANIZATIONS",
      folder: "01_DATABASE",
      headers: [
        "ORG_ID", "ORG_NAME", "PRIMARY_EMAIL", "PRIMARY_PHONE", "ADDRESS",
        "ORG_CODE", "CLIENT_ADMIN_USER_ID", "ASSIGNED_STAFF_EMAIL", "STATUS",
        "DATE_CREATED", "NOTES"
      ]
    },
    CASE: {
      name: "CASE_DATABASE",
      folder: "05_CASE_DATABASE",
      headers: [
        "CASE_ID", "CLIENT_ID", "CLIENT_NAME", "PATENT_TITLE",
        "APPLICATION_NUMBER", "COUNTRY", "FILING_DATE", "CURRENT_STATUS",
        "NEXT_DEADLINE", "ATTORNEY", "CASE_FOLDER_ID", "PRIORITY",
        "PATENT_TYPE", "ORG_ID", "ASSIGNED_STAFF_EMAIL", "GALVANIZER_EMAIL",
        "WORKFLOW_STAGE", "NOTES", "CREATED_DATE", "LAST_UPDATED"
      ]
    },
    INVOICE: {
      name: "INVOICE_DATABASE",
      folder: "04_INVOICE_SYSTEM",
      headers: [
        "Client Manager", "Docket Number", "Docket# [Invoice UIN]", "Invoice Date",
        "Tax Invoice Date", "Tax Serial Number", "Tax Invoice Number",
        "Client Reference", "ClientCode", "Invention#", "Patent Office",
        "First Inventor", "Title Identfication (Only key words, Enter in lower case}",
        "Entity Status", "PO Application # e.g., US 15339876", "Main ServiceCode",
        "Service Code 2", "Service Code 3",
        "Additional information on service (Do not repeat what is in code description)",
        "State of Supply", "Name of Client", "POFeeAck", "POFee", "ServFee",
        "TAX_TYPE", "IGST", "SGST", "CGST", "Expenses", "InvAmount",
        "Amount due", "Attorney Fee", "Consultant fee", "Referral fee",
        "Net Revenue", "PAYMENT_STATUS", "PAYMENT_DATE", "PAYMENT_MODE",
        "PAYMENT_AMOUNT", "INVOICE_PDF_LINK", "CLIENT_ID", "CLIENT_NAME",
        "CASE_ID", "ORG_ID", "NOTES", "INVOICE_ID"
      ]
    },
    USERS: {
      name: "USER_ROLES",
      folder: "01_DATABASE",
      headers: [
        "USER_ID", "EMAIL", "FULL_NAME", "ROLE", "CLIENT_ID",
        "ORG_ID", "DEPARTMENT", "STATUS", "PASSWORD_HASH",
        "CAN_VIEW_FINANCE", "REPORTS_TO", "ADDITIONAL_ROLES", "CREATED_DATE"
      ]
    },
    CIRCLES: {
      name: "CIRCLES",
      folder: "01_DATABASE",
      headers: [
        "CIRCLE_ID", "CIRCLE_NAME", "DESCRIPTION", "STATUS",
        "CREATED_BY", "CREATED_AT", "UPDATED_AT"
      ]
    },
    CIRCLE_MEMBERS: {
      name: "CIRCLE_MEMBERS",
      folder: "01_DATABASE",
      headers: [
        "MEMBERSHIP_ID", "CIRCLE_ID", "USER_ID", "USER_EMAIL",
        "ROLE_IN_CIRCLE", "STATUS", "ADDED_BY", "CREATED_AT", "UPDATED_AT"
      ]
    },
    DAILY_PRIORITIES: {
      name: "DAILY_PRIORITIES",
      folder: "06_AUTOMATION",
      headers: [
        "ENTRY_ID", "USER_EMAIL", "USER_NAME", "ROLE", "ENTRY_DATE",
        "PRIORITY_1", "PRIORITY_2", "PRIORITY_3", "NOTES",
        "STATUS", "CREATED_AT", "UPDATED_AT"
      ]
    },
    DAILY_WRAPUPS: {
      name: "DAILY_WRAPUPS",
      folder: "06_AUTOMATION",
      headers: [
        "WRAPUP_ID", "USER_EMAIL", "USER_NAME", "ROLE", "ENTRY_DATE",
        "HIGH_POINTS", "LOW_POINTS", "HELP_NEEDED", "ADMIN_REVIEW_STATUS",
        "ADMIN_REVIEW_NOTES", "CREATED_AT", "UPDATED_AT"
      ]
    },
    EXPENSE_CLAIMS: {
      name: "EXPENSE_CLAIMS",
      folder: "04_INVOICE_SYSTEM",
      headers: [
        "CLAIM_ID", "USER_EMAIL", "USER_NAME", "ROLE", "CLAIM_DATE",
        "CATEGORY", "AMOUNT", "DESCRIPTION", "BILL_LINK", "STATUS",
        "SUBMITTED_TO", "ADMIN_REMARKS", "CREATED_AT", "UPDATED_AT"
      ]
    },
    MESSAGE_THREADS: {
      name: "MESSAGE_THREADS",
      folder: "03_CLIENT_PORTAL",
      headers: [
        "THREAD_ID", "THREAD_TYPE", "TITLE", "RELATED_ENTITY_TYPE",
        "RELATED_ENTITY_ID", "ORG_ID", "CLIENT_ID", "VISIBLE_TO_CLIENT",
        "CREATED_BY", "LAST_MESSAGE_AT", "STATUS", "CREATED_AT"
      ]
    },
    MESSAGES: {
      name: "MESSAGES",
      folder: "03_CLIENT_PORTAL",
      headers: [
        "MESSAGE_ID", "THREAD_ID", "SENDER_EMAIL", "SENDER_NAME", "SENDER_ROLE",
        "MESSAGE_TEXT", "IS_INTERNAL", "CREATED_AT", "READ_BY"
      ]
    },
    NOTIFICATIONS: {
      name: "NOTIFICATIONS",
      folder: "03_CLIENT_PORTAL",
      headers: [
        "NOTIFICATION_ID", "USER_EMAIL", "TITLE", "BODY", "TYPE",
        "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "IS_READ",
        "CREATED_AT", "READ_AT"
      ]
    },
    MESSAGE_PARTICIPANTS: {
      name: "MESSAGE_PARTICIPANTS",
      folder: "03_CLIENT_PORTAL",
      headers: [
        "PARTICIPANT_ID", "THREAD_ID", "USER_EMAIL", "USER_NAME", "USER_ROLE",
        "PARTICIPANT_TYPE", "STATUS", "CREATED_AT", "UPDATED_AT"
      ]
    },
    TASKS: {
      name: "TASKS",
      folder: "06_AUTOMATION",
      headers: [
        "TASK_ID", "TITLE", "DESCRIPTION", "STATUS", "PRIORITY",
        "ASSIGNED_TO_EMAIL", "ASSIGNED_TO_NAME", "ASSIGNED_BY_EMAIL", "ASSIGNED_BY_NAME",
        "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "CLIENT_ID", "ORG_ID",
        "DUE_DATE", "START_DATE", "COMPLETED_AT", "TAGS", "NOTES",
        "CREATED_AT", "UPDATED_AT"
      ]
    },
    ACTIVITY_TIMELINE: {
      name: "ACTIVITY_TIMELINE",
      folder: "06_AUTOMATION",
      headers: [
        "EVENT_ID", "EVENT_TYPE", "TITLE", "DESCRIPTION", "ENTITY_TYPE",
        "ENTITY_ID", "CLIENT_ID", "ORG_ID", "CASE_ID", "USER_EMAIL",
        "USER_NAME", "VISIBILITY", "CREATED_AT"
      ]
    },
    APPROVALS: {
      name: "APPROVALS",
      folder: "06_AUTOMATION",
      headers: [
        "APPROVAL_ID", "APPROVAL_TYPE", "TITLE", "DESCRIPTION", "STATUS",
        "REQUESTED_BY_EMAIL", "REQUESTED_BY_NAME", "APPROVER_EMAIL", "APPROVER_ROLE",
        "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "CLIENT_ID", "ORG_ID",
        "REQUEST_DATE", "DECISION_DATE", "DECISION_NOTES", "CREATED_AT", "UPDATED_AT"
      ]
    },
    DOCUMENT_REQUESTS: {
      name: "DOCUMENT_REQUESTS",
      folder: "03_CLIENT_PORTAL",
      headers: [
        "REQUEST_ID", "TITLE", "DESCRIPTION", "REQUEST_TYPE", "STATUS",
        "CLIENT_ID", "ORG_ID", "CASE_ID", "REQUESTED_BY_EMAIL", "ASSIGNED_TO_EMAIL",
        "DUE_DATE", "DRIVE_LINK", "CLIENT_VISIBLE", "APPROVAL_STATUS",
        "CREATED_AT", "UPDATED_AT"
      ]
    },
    ACTIVITY_LOG: {
      name: "ACTIVITY_LOG",
      folder: "06_AUTOMATION",
      headers: [
        "TIMESTAMP", "USER_EMAIL", "ACTION", "ENTITY_TYPE",
        "ENTITY_ID", "DETAILS"
      ]
    }
  },
  CLIENT_SUBFOLDERS: [
    "APPLICATIONS", "OFFICE_ACTIONS", "RESPONSES",
    "INVOICES", "CERTIFICATES", "COMMUNICATION"
  ]
};

/**
 * ════════════════════════════════════════════════════════════
 * MASTER SETUP — Run this function ONCE
 * ════════════════════════════════════════════════════════════
 */
function setupEntireSystem() {
  var ui = SpreadsheetApp.getUi();
  var folderPrompt = ui.prompt(
    "IP Patent Firm System Setup",
    "Enter the root folder name to use in Google Drive:",
    ui.ButtonSet.OK_CANCEL
  );

  if (folderPrompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Setup cancelled.");
    return;
  }

  var folderName = folderPrompt.getResponseText().trim() || CONFIG.DEFAULT_ROOT_FOLDER_NAME;
  var response = ui.alert(
    "⚙️ IP Patent Firm System Setup",
    "This will create the entire folder architecture, spreadsheets, " +
    "and initial configuration in your Google Drive.\n\n" +
    "Proceed?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("Setup cancelled.");
    return;
  }

  try {
    Logger.log("🚀 Starting IP Patent Firm System Setup...");

    // Step 1: Create root folder
    var rootFolder = createRootFolder_(folderName);
    Logger.log("✅ Root folder created: " + rootFolder.getId());

    // Step 2: Create subfolders
    var subFolders = createSubFolders_(rootFolder);
    Logger.log("✅ Subfolders created");

    // Step 3: Create all spreadsheets
    var spreadsheetIds = createAllSpreadsheets_(subFolders);
    Logger.log("✅ Spreadsheets created");

    // Step 4: Store system config in Script Properties
    storeSystemConfig_(rootFolder, subFolders, spreadsheetIds);
    Logger.log("✅ System config stored");

    // Step 5: Set up initial admin user
    setupInitialAdmin_();
    Logger.log("✅ Initial admin configured");

    // Step 6: Format sheets
    formatAllSheets_(spreadsheetIds);
    Logger.log("✅ Sheets formatted");

    // Step 7: Create custom menus
    createCustomMenus_();
    Logger.log("✅ Custom menus created");

    ui.alert(
      "✅ Setup Complete!",
      "Your IP Patent Firm System is ready.\n\n" +
      "📁 Root Folder: " + rootFolder.getName() + "\n" +
      "📊 All databases created and formatted.\n" +
      "👤 You are set as Admin.\n\n" +
      "Check your Google Drive for the folder structure.",
      ui.ButtonSet.OK
    );

    Logger.log("🎉 Setup completed successfully!");

  } catch (e) {
    Logger.log("❌ Setup error: " + e.toString());
    ui.alert("❌ Error", "Setup failed: " + e.message, ui.ButtonSet.OK);
  }
}

// ── Helper: Create Root Folder ─────────────────────────────
function createRootFolder_(folderName) {
  var existing = DriveApp.getFoldersByName(folderName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(folderName);
}

// ── Helper: Create Subfolders ──────────────────────────────
function createSubFolders_(rootFolder) {
  var folders = {};
  CONFIG.SUBFOLDERS.forEach(function(name) {
    var iter = rootFolder.getFoldersByName(name);
    if (iter.hasNext()) {
      folders[name] = iter.next();
    } else {
      folders[name] = rootFolder.createFolder(name);
    }
  });
  return folders;
}

// ── Helper: Create All Spreadsheets ────────────────────────
function createAllSpreadsheets_(subFolders) {
  var ids = {};

  Object.keys(CONFIG.SHEETS).forEach(function(key) {
    var sheetConfig = CONFIG.SHEETS[key];
    var targetFolder = subFolders[sheetConfig.folder];

    // Check if spreadsheet already exists
    var existing = targetFolder.getFilesByName(sheetConfig.name);
    var ss;
    if (existing.hasNext()) {
      ss = SpreadsheetApp.open(existing.next());
    } else {
      ss = SpreadsheetApp.create(sheetConfig.name);
      // Move to correct folder
      var file = DriveApp.getFileById(ss.getId());
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }

    // Set up headers on first sheet without duplicating existing ones
    var sheet = ss.getSheets()[0];
    sheet.setName(sheetConfig.name);
    ensureHeaders_(sheet, sheetConfig.headers);

    ids[key] = ss.getId();
  });

  return ids;
}

// ── Helper: Store Config in Script Properties ──────────────
function storeSystemConfig_(rootFolder, subFolders, spreadsheetIds) {
  var props = PropertiesService.getScriptProperties();

  props.setProperty("ROOT_FOLDER_ID", rootFolder.getId());

  Object.keys(subFolders).forEach(function(name) {
    props.setProperty("FOLDER_" + name, subFolders[name].getId());
  });

  Object.keys(spreadsheetIds).forEach(function(key) {
    props.setProperty("SHEET_" + key, spreadsheetIds[key]);
  });

  props.setProperty("SYSTEM_INITIALIZED", "true");
  props.setProperty("SETUP_DATE", new Date().toISOString());
  props.setProperty("ROOT_FOLDER_NAME", rootFolder.getName());
}

function ensureHeaders_(sheet, headers) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var finalHeaders = existing.slice();

  headers.forEach(function(header, idx) {
    if (!finalHeaders[idx]) {
      finalHeaders[idx] = header;
    }
    if (finalHeaders.indexOf(header) === -1) {
      finalHeaders.push(header);
    }
  });

  finalHeaders = finalHeaders.filter(function(value) { return value; });
  sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
}

// ── Helper: Setup Initial Admin ────────────────────────────
function setupInitialAdmin_() {
  var email = Session.getActiveUser().getEmail();
  var props = PropertiesService.getScriptProperties();
  var usersSheetId = props.getProperty("SHEET_USERS");
  var ss = SpreadsheetApp.openById(usersSheetId);
  var sheet = ss.getSheetByName("USER_ROLES");

  // Check if admin already exists
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) return;
  }

  sheet.appendRow([
    "USR001", email, "System Admin", "Super Admin", "",
    "", "Management", "Active", "", "Yes", "", "",
    new Date()
  ]);
}

// ── Helper: Format All Sheets ──────────────────────────────
function formatAllSheets_(spreadsheetIds) {
  Object.keys(spreadsheetIds).forEach(function(key) {
    var ss = SpreadsheetApp.openById(spreadsheetIds[key]);
    var sheet = ss.getSheets()[0];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn());

    // Style headers
    headers.setBackground("#1a237e")
           .setFontColor("#ffffff")
           .setFontWeight("bold")
           .setFontSize(10)
           .setHorizontalAlignment("center");

    // Freeze header row
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (var c = 1; c <= sheet.getLastColumn(); c++) {
      sheet.autoResizeColumn(c);
    }

    // Set minimum column width
    for (var c = 1; c <= sheet.getLastColumn(); c++) {
      if (sheet.getColumnWidth(c) < 120) {
        sheet.setColumnWidth(c, 120);
      }
    }
  });
}

// ── Custom Menus ───────────────────────────────────────────
function onOpen() {
  createCustomMenus_();
}

function createCustomMenus_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏢 IP Firm System")
    .addItem("➕ New Client", "showCreateClientDialog")
    .addItem("📋 New Case", "showCreateCaseDialog")
    .addSeparator()
    .addItem("🧾 Generate Invoice", "showGenerateInvoiceDialog")
    .addItem("⏰ Send Deadline Alerts", "sendDeadlineAlerts")
    .addSeparator()
    .addItem("📊 Analytics Dashboard", "showAnalyticsDashboard")
    .addItem("🌐 Deploy Client Portal", "deployPortalInfo")
    .addSeparator()
    .addItem("👤 Register User", "showRegisterUserDialog")
    .addItem("⚙️ System Setup", "setupEntireSystem")
    .addToUi();
}

// ── Utility: Get Sheet by Config Key ───────────────────────
function getSheet_(configKey) {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty("SHEET_" + configKey);
  if (!sheetId) throw new Error("System not initialized. Run setupEntireSystem() first.");
  var ss = SpreadsheetApp.openById(sheetId);
  return ss.getSheets()[0];
}

// ── Utility: Get Folder by Name ────────────────────────────
function getFolder_(folderName) {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty("FOLDER_" + folderName);
  if (!folderId) throw new Error("Folder not found: " + folderName);
  return DriveApp.getFolderById(folderId);
}

// ── Utility: Generate Sequential ID ───────────────────────
function generateId_(prefix, sheet, idColumn) {
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;

  for (var i = 1; i < data.length; i++) {
    var id = data[i][idColumn];
    if (id && id.toString().startsWith(prefix)) {
      var num = parseInt(id.toString().replace(prefix, ""), 10);
      if (num > maxNum) maxNum = num;
    }
  }

  var nextNum = maxNum + 1;
  var padLength = prefix === "CL" ? 3 : 4;
  return prefix + ("0".repeat(padLength) + nextNum).slice(-padLength);
}

function withScriptLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ── Utility: Log Activity ─────────────────────────────────
function logActivity_(action, entityType, entityId, details) {
  try {
    var sheet = getSheet_("ACTIVITY_LOG");
    sheet.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || "system",
      action,
      entityType,
      entityId,
      details || ""
    ]);
  } catch (e) {
    Logger.log("Activity log error: " + e.message);
  }
}

// ── Utility: Hash Password (SHA-256) ──────────────────────
function hashPassword_(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return rawHash.map(function(byte) {
    var v = (byte < 0) ? byte + 256 : byte;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function updateInvoiceHeadersOnly() {
  var sheet = getSheet_("INVOICE");
  ensureHeaders_(sheet, CONFIG.SHEETS.INVOICE.headers);
  SpreadsheetApp.getUi().alert("INVOICE headers updated.");
}
