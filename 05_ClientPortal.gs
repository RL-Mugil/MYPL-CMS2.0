/**
 * ============================================================
 * 05_ClientPortal.gs — Web App Client Portal with Login
 * ============================================================
 * Deploy as Web App:
 *   Execute as: Me (the script owner)
 *   Who has access: Anyone
 * ============================================================
 */

/**
 * Web App entry point
 */
function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  var result;
  try {
    var payload;

    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    else if (e.parameter && e.parameter.action) {
      payload = {
        action: e.parameter.action,
        params: e.parameter.params ? JSON.parse(e.parameter.params) : {}
      };
    }
    else {
      result = { error: 'No action specified' };
      output.setContent(JSON.stringify(result));
      return output;
    }

    result = getPortalData(payload.action, payload.params || {});
  } catch (err) {
    result = { error: 'Request error: ' + err.message };
  }

  output.setContent(JSON.stringify(result));
  return output;
}

function getPortalCache_() {
  return CacheService.getScriptCache();
}

function getPortalCacheVersion_(scope) {
  var key = "PORTAL_CACHE_VER_" + String(scope || "global").toUpperCase();
  return PropertiesService.getScriptProperties().getProperty(key) || "1";
}

function bumpPortalCacheVersions_(scopes) {
  var props = PropertiesService.getScriptProperties();
  (scopes || []).forEach(function(scope) {
    var key = "PORTAL_CACHE_VER_" + String(scope || "global").toUpperCase();
    props.setProperty(key, String(new Date().getTime()));
  });
}

function buildPortalCacheKey_(scope, session, extra) {
  var email = session && session.email ? String(session.email).toLowerCase() : "public";
  return ["portal", scope, getPortalCacheVersion_(scope), email, extra || ""].join("|");
}

function readPortalCache_(key) {
  var cached = getPortalCache_().get(key);
  if (!cached) return null;
  try {
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

function writePortalCache_(key, value, ttlSeconds) {
  try {
    getPortalCache_().put(key, JSON.stringify(value), ttlSeconds || 60);
  } catch (e) {}
  return value;
}

function removePortalCacheKeys_(keys) {
  if (!keys || !keys.length) return;
  try {
    getPortalCache_().removeAll(keys);
  } catch (e) {
    keys.forEach(function(key) {
      try { getPortalCache_().remove(key); } catch (inner) {}
    });
  }
}

function clearPortalCaches_(scopesOrSession) {
  var scopes = Object.prototype.toString.call(scopesOrSession) === "[object Array]"
    ? scopesOrSession
    : ["dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "users", "clients", "organizations", "circles", "circleMembers", "dailyOps", "dailyAudit", "expenses", "messages", "directInbox", "notifications", "galvanizerQueue", "tasks", "timeline", "workflowBoard", "attorneyWorkspace", "approvals", "documentRequests", "smartSearch"];
  bumpPortalCacheVersions_(scopes);
}

// ── Authentication Functions ────────────────────────────────

/**
 * Login with email and password
 */
function portalLogin(email, password) {
  if (!email || !password) return { success: false, message: "Email and password required." };
  var user = getUserRecordByEmail_(email);
  var passwordHash = hashPassword_(password);

  if (!user || String(user.STATUS || "") !== "Active") {
    return { success: false, message: "Email not found or account inactive." };
  }

  if (!user.PASSWORD_HASH) {
    return { success: false, message: "Local password login is not enabled for this account." };
  }

  if (user.PASSWORD_HASH !== passwordHash) {
    return { success: false, message: "Incorrect password." };
  }

  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put("session_" + token, JSON.stringify(buildSessionFromUser_(user)), 21600);
  logActivity_("LOGIN", "USER", user.USER_ID, "Login from portal");

  return {
    success: true,
    token: token,
    role: normalizeRole_(user.ROLE),
    name: user.FULL_NAME || user.EMAIL
  };
}

/**
 * Validate session token and return user info
 */
function validateSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var sessionData = cache.get("session_" + token);
  if (!sessionData) return null;
  return JSON.parse(sessionData);
}

/**
 * Logout — remove session
 */
function portalLogout(token) {
  if (token) {
    var cache = CacheService.getScriptCache();
    cache.remove("session_" + token);
  }
  return { success: true };
}

/**
 * Change password
 */
function changePassword(token, oldPassword, newPassword) {
  var session = validateSession(token);
  if (!session) return { success: false, message: "Session expired. Please login again." };
  
  if (!newPassword || newPassword.length < 6) {
    return { success: false, message: "Password must be at least 6 characters." };
  }
  
  var user = getUserRecordByEmail_(session.email);
  if (!user) return { success: false, message: "User not found." };
  if (!user.PASSWORD_HASH) return { success: false, message: "Local password login is not enabled for this account." };

  var oldHash = hashPassword_(oldPassword);
  var newHash = hashPassword_(newPassword);

  if (user.PASSWORD_HASH !== oldHash) {
    return { success: false, message: "Current password is incorrect." };
  }

  updateRecordById_("USERS", "USER_ID", user.USER_ID, { PASSWORD_HASH: newHash });
  return { success: true, message: "Password changed successfully." };
}

/**
 * API endpoint for portal AJAX calls — now token-based
 */
/**
 * Health Check function for monitoring system status
 */
function health_check() {
  try {
    var initialized = PropertiesService.getScriptProperties().getProperty("SYSTEM_INITIALIZED");
    return { 
      status: initialized ? "OK" : "NOT_INITIALIZED", 
      timestamp: new Date(),
      user: Session.getActiveUser().getEmail()
    };
  } catch (e) {
    return { status: "ERROR", message: e.message };
  }
}

/**
 * API endpoint for portal AJAX calls — now token-based with full error handling
 */
function getPortalData(action, params) {
  Logger.log("API_CALL [" + action + "] by " + Session.getActiveUser().getEmail());
  try {
    // Login doesn't need a token
    if (action === "login") {
      return portalLogin(params.email, params.password);
    }
    if (action ==="clerklogin") {
      return clerkLogin_(params.email);
    }
    if (action === "logout") {
      return portalLogout(params.token);
    }
    if (action === "health") {
      return health_check();
    }
    
    // All other actions need a valid session
    var session = validateSession(params.token);
    if (!session) return { error: "Session expired. Please login again.", sessionExpired: true };
    
    var userRole = session;
  
    switch (action) {
      case "getDashboard":
        return getPortalDashboard_(session.email, userRole, params.filters || {});
      case "getDashboardSummary":
        return getPortalDashboardSummary_(session.email, userRole, params.filters || {});
      case "getDashboardDetails":
        return getPortalDashboardDetails_(session.email, userRole, params.filters || {});
      case "getCases":
        return getPortalCases_(session.email, userRole, params.filters || {});
      case "getCasesPage":
        return getPortalCasesPage_(session.email, userRole, params.filters || {}, params.limit, params.offset);
      case "getInvoices":
        return getPortalInvoices_(session.email, userRole);
      case "getDocuments":
        return getPortalDocuments_(session.email, userRole);
      case "submitContact":
        return submitContactForm_(session.email, params);
      case "getUserInfo":
        return {
          email: session.email,
          role: userRole.role,
          name: userRole.name,
          clientId: userRole.clientId || "",
          orgId: userRole.orgId || "",
          canViewFinance: canViewFinance_(userRole)
        };
      case "changePassword":
        return changePassword(params.token, params.oldPassword, params.newPassword);
      case "markInvoicePaid":
        if (!canViewFinance_(userRole) || !hasRoleAtLeast_(getEffectiveRoles_(userRole), "Admin")) return { error: "Access denied. Admin only." };
        return updatePaymentStatus(params.invoiceId, "Paid", new Date(), "Portal Admin");
      case "updateInvoicePayment":
        if (!canViewFinance_(userRole) || !hasRoleAtLeast_(getEffectiveRoles_(userRole), "Admin")) return { error: "Access denied. Admin only." };
        return updateInvoicePayment_(params.invoiceId, params.paymentData || {});
      case "sendInvoice":
        if (!canViewFinance_(userRole) || !hasRoleAtLeast_(getEffectiveRoles_(userRole), "Admin")) return { error: "Access denied. Admin only." };
        return sendInvoiceById_(params.invoiceId, params.email || "");
      
      // --- Management API (Admin Only) ---
      case "getUsers":
        if (!canAccessManagement_(userRole)) return { error: "Access denied." };
        return getPortalUsers_(userRole, params.filters || {});
      case "saveUser":
        if (!canAccessManagement_(userRole)) return { error: "Access denied." };
        return savePortalUser_(params.userData, userRole);
      case "deleteUser":
        if (!canAccessManagement_(userRole)) return { error: "Access denied." };
        return deletePortalUser_(params.userId, userRole);
        
      case "getClients":
        if (!canAccessManagement_(userRole)) return { error: "Access denied." };
        return getPortalClients_(userRole);
      case "getAccessibleClients":
        return getPortalClients_(userRole);
      case "saveClient":
        if (!canAccessManagement_(userRole)) return { error: "Access denied." };
        return savePortalClient_(params.clientData);
      case "deleteClient":
        if (!canManageAllData_(userRole)) return { error: "Access denied." };
        return deletePortalClient_(params.clientId);
        
      case "saveCase":
        if (!hasRoleAtLeast_(getEffectiveRoles_(userRole), "Staff")) return { error: "Access denied." };
        return savePortalCase_(params.caseData);
      case "bulkUpdateCases":
        return bulkUpdateCases_(userRole, params.bulkData || {});
      case "bulkImportDocketTrakRows":
        return bulkImportDocketTrakRows_(userRole, params.importData || {});
      case "deleteCase":
        if (!canManageAllData_(userRole)) return { error: "Access denied." };
        return deletePortalCase_(params.caseId);
        
      case "saveInvoice":
        if (!canViewFinance_(userRole) || !hasRoleAtLeast_(getEffectiveRoles_(userRole), "Admin")) return { error: "Access denied." };
        return savePortalInvoice_(params.invoiceData);
      case "deleteInvoice":
        if (!canViewFinance_(userRole) || !hasRoleAtLeast_(getEffectiveRoles_(userRole), "Admin")) return { error: "Access denied." };
        return deletePortalInvoice_(params.invoiceId);
      case "getOrganizations":
        return listOrganizationsForSession_(userRole);
      case "saveOrganization":
        return saveOrganization_(userRole, params.orgData || {});
      case "getOrganizationUsers":
        return getOrganizationUsers_(userRole, params.orgId);
      case "saveDailyPriority":
        return saveDailyPriority_(userRole, params.priorityData || {});
      case "saveDailyWrapup":
        return saveDailyWrapup_(userRole, params.wrapupData || {});
      case "getDailyOpsOverview":
        return getDailyOpsOverview_(userRole);
      case "getDailyAudit":
        return getDailyAuditView_(userRole, params.filters || {});
      case "submitExpenseClaim":
        return submitExpenseClaim_(userRole, params.claimData || {});
      case "getExpenseClaims":
        return getExpenseClaims_(userRole, params.filters || {});
      case "reviewExpenseClaim":
        return saveExpenseClaimReview_(userRole, params.claimId, params.reviewData || {});
      case "getNotifications":
        return getNotifications_(userRole);
      case "markNotificationRead":
        return markNotificationRead_(userRole, params.notificationId);
      case "deleteNotification":
        return deleteNotification_(userRole, params.notificationId);
      case "clearNotifications":
        return clearNotifications_(userRole);
      case "getMessageThreads":
        return getMessageThreads_(userRole);
      case "getDirectInbox":
        return getDirectInbox_(userRole);
      case "getThreadMessages":
        return getMessagesForThread_(userRole, params.threadId);
      case "markThreadRead":
        return markThreadRead_(userRole, params.threadId);
      case "saveMessageThread":
        return saveMessageThread_(userRole, params.threadData || {});
      case "createDirectThread":
        return createDirectThread_(userRole, params.threadData || {});
      case "deleteMessageThread":
        return deleteMessageThread_(userRole, params.threadId);
      case "sendThreadMessage":
        return sendThreadMessage_(userRole, params.messageData || {});
      case "getTasks":
        return getTasks_(userRole, params.filters || {});
      case "saveTask":
        return saveTask_(userRole, params.taskData || {});
      case "updateTaskStatus":
        return updateTaskStatus_(userRole, params.taskId, params.status, params.notes || "");
      case "getActivityTimeline":
        return getActivityTimeline_(userRole, params.filters || {});
      case "getWorkflowBoard":
        return getWorkflowBoard_(userRole, params.filters || {});
      case "getAttorneyWorkspace":
        return getAttorneyWorkspace_(userRole, params.filters || {});
      case "getSmartSearch":
        return getSmartSearch_(userRole, params.query || "", params.scope || "all");
      case "getApprovalRequests":
        return getApprovalRequests_(userRole, params.filters || {});
      case "saveApprovalRequest":
        return saveApprovalRequest_(userRole, params.approvalData || {});
      case "reviewApprovalRequest":
        return reviewApprovalRequest_(userRole, params.approvalId, params.reviewData || {});
      case "getDocumentRequests":
        return getDocumentRequests_(userRole, params.filters || {});
      case "saveDocumentRequest":
        return saveDocumentRequest_(userRole, params.requestData || {});
      case "reviewDocumentRequest":
        return reviewDocumentRequest_(userRole, params.requestId, params.reviewData || {});
      case "uploadPortalDocument":
        return uploadPortalDocument_(userRole, params.documentData || {});
      case "uploadExpenseBill":
        return uploadExpenseBill_(userRole, params.fileData || {});
      case "getGalvanizerCommandCenter":
        return getGalvanizerCommandCenter_(userRole, params.filters || {});
      case "getCircles":
        return getCircles_(userRole);
      case "getGalvanizerQueue":
        return getGalvanizerQueue_(userRole, params.filters || {});
      case "getCircleMembers":
        return getCircleMembers_(userRole, params.circleId);
      case "saveCircle":
        return saveCircle_(userRole, params.circleData || {});
      case "deleteCircle":
        return deleteCircle_(userRole, params.circleId);
      case "saveCircleMember":
        return saveCircleMember_(userRole, params.memberData || {});
      case "removeCircleMember":
        return removeCircleMember_(userRole, params.membershipId);
        
      default:
        return { error: "Unknown action: " + action };
    }
  } catch (e) {
    Logger.log("SERVER_ERROR [" + action + "]: " + e.message);
    // Return specific error to client so they know WHAT failed (e.g. "System not initialized")
    return { error: "Server Error: " + e.message };
  }
}


// ── Portal Management API Handlers (Admin Only) ───────────────

function getPortalUsers_(session, filters) {
  var cacheKey = buildPortalCacheKey_("users", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var sheet = getSheet_("USERS");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (session && !canManageAllData_(session) && session.orgId && data[i][headers.indexOf("ORG_ID")] !== session.orgId) {
      continue;
    }
    var u = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === "PASSWORD_HASH") continue;
      u[headers[j]] = data[i][j];
    }
    users.push(u);
  }
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(filterUsers_(users, filters)), 300);
}

function savePortalUser_(userData, session) {
  var sheet = getSheet_("USERS");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  if (userData.ORG_ID) {
    var normalizedRole = normalizeRole_(userData.ROLE || userData.role);
    if (normalizedRole !== "Client Admin") {
      userData.ROLE = "Client Employee";
    }
    userData.DEPARTMENT = "External";
    userData.CAN_VIEW_FINANCE = "No";
  }
  if (session && !canManageAllData_(session) && normalizeRole_(session.role) === "Client Admin") {
    userData.ORG_ID = session.orgId;
    if (["Client Admin", "Client Employee"].indexOf(normalizeRole_(userData.ROLE || userData.role)) === -1) {
      return { error: "Client Admin can manage only company users." };
    }
  }
  
  if (userData.USER_ID) {
    // Update existing
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userData.USER_ID) {
        var keys = Object.keys(userData);
        for (var k = 0; k < keys.length; k++) {
          var colIdx = headers.indexOf(keys[k]);
          // Don't update ID, and don't update password if it's empty
          if (colIdx > 0 && !(keys[k] === 'PASSWORD_HASH' && !userData[keys[k]]) && keys[k] !== 'PASSWORD') {
            var val = userData[keys[k]];
            if (keys[k] === 'PASSWORD_HASH') { val = hashPassword_(val); }
            sheet.getRange(i + 1, colIdx + 1).setValue(val);
          }
        }
        if (userData.PASSWORD) {
          var pwdColIdx = headers.indexOf('PASSWORD_HASH');
          if (pwdColIdx > -1) {
            sheet.getRange(i + 1, pwdColIdx + 1).setValue(hashPassword_(userData.PASSWORD));
          }
        }
        clearPortalCaches_(["users", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations", "circles", "circleMembers"]);
        return { success: true, message: "User updated successfully" };
      }
    }
    return { error: "User not found" };
  } else {
    // Creating from portal now handled via registerUser (already exists), but we can route here if needed.
    // However, the prompt says forms will submit here. Let's just use registerUser logic.
    if (!userData.EMAIL) return { error: "Missing email" };
    var createResult = registerUser({
      fullName: userData.FULL_NAME,
      email: userData.EMAIL,
      password: userData.PASSWORD || userData.PASSWORD_HASH || "",
      role: userData.ROLE,
      clientId: userData.CLIENT_ID,
      orgId: userData.ORG_ID,
      department: userData.DEPARTMENT,
      canViewFinance: userData.CAN_VIEW_FINANCE || "No",
      reportsTo: userData.REPORTS_TO || "",
      additionalRoles: userData.ADDITIONAL_ROLES || "",
      createdBy: session ? session.email : ""
    });
    if (createResult && createResult.success) clearPortalCaches_(["users", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations", "circles", "circleMembers"]);
    return createResult;
  }
}

function deletePortalUser_(userId, session) {
  var sheet = getSheet_("USERS");
  var data = sheet.getDataRange().getValues();
  var statusCol = data[0].indexOf("STATUS");
  if (statusCol === -1) statusCol = 6; // Fallback based on registerUser
  var orgCol = data[0].indexOf("ORG_ID");
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      if (session && !canManageAllData_(session) && session.orgId && orgCol > -1 && data[i][orgCol] !== session.orgId) {
        return { error: "Access denied." };
      }
      sheet.getRange(i + 1, statusCol + 1).setValue("Inactive");
      clearPortalCaches_(["users", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations", "circles", "circleMembers"]);
      return { success: true, message: "User deleted (Inactive)" };
    }
  }
  return { error: "User not found" };
}

function getPortalClients_(session) {
  var cacheKey = buildPortalCacheKey_("clients", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var result = !session ? sanitizeDataForFrontend_(getAllClients()) : sanitizeDataForFrontend_(getAccessibleClientsForUser_(session));
  return writePortalCache_(cacheKey, result, 300);
}

function savePortalClient_(clientData) {
  if (String(clientData.CLIENT_TYPE || "Individual") === "Individual") {
    clientData.ORG_ID = "";
    clientData.CLIENT_ADMIN_USER_ID = "";
  }
  if (clientData.CLIENT_ID) {
    // Map to the format updateClient expects logic
    var updates = Object.assign({}, clientData);
    delete updates.CLIENT_ID; // Remove so we don't accidentally update ID
    var res = updateClient(clientData.CLIENT_ID, updates);
    if(res.success) {
      clearPortalCaches_(["clients", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations"]);
      return { success: true, message: "Client updated successfully" };
    }
    return res;
  } else {
    // Use existing createClient
    var res = createClient({
      clientName: clientData.CLIENT_NAME,
      contactPerson: clientData.CONTACT_PERSON,
      email: clientData.EMAIL,
      phone: clientData.PHONE,
      address: clientData.ADDRESS,
      notes: clientData.NOTES,
      clientCode: clientData.CLIENT_CODE || "",
      clientRegion: clientData.CLIENT_REGION || "India",
      clientType: clientData.CLIENT_TYPE || "Individual",
      orgId: clientData.ORG_ID || "",
      clientAdminUserId: clientData.CLIENT_ADMIN_USER_ID || "",
      assignedStaffEmail: clientData.ASSIGNED_STAFF_EMAIL || ""
    });
    if(res.success) {
      clearPortalCaches_(["clients", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations"]);
      return { success: true, message: res.message };
    }
    return res;
  }
}

function deletePortalClient_(clientId) {
  var sheet = getSheet_("MASTER_CLIENT");
  var data = sheet.getDataRange().getValues();
  var statusCol = data[0].indexOf("STATUS");
  if (statusCol === -1) statusCol = 7;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) {
      sheet.getRange(i + 1, statusCol + 1).setValue("Deleted");
      clearPortalCaches_(["clients", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage", "organizations"]);
      return { success: true, message: "Client deleted" };
    }
  }
  return { error: "Client not found" };
}

function savePortalCase_(caseData) {
  function notifyAssignee_(email, label, caseId, title) {
    if (!email) return;
    createNotification_(email, "Case assigned", label + " assignment for " + caseId, "CASE", caseId);
    if (title) {
      try {
        createTimelineEvent_({
          EVENT_TYPE: "CASE_ASSIGNED",
          TITLE: "Case assigned",
          DESCRIPTION: label + " assigned for " + caseId + " - " + title,
          ENTITY_TYPE: "CASE",
          ENTITY_ID: caseId,
          CASE_ID: caseId,
          USER_EMAIL: email,
          USER_NAME: email,
          VISIBILITY: "Internal"
        });
      } catch (e) {}
    }
  }

  if (caseData.CASE_ID) {
    var existingCase = (getAllCases() || []).find(function(item) { return item.CASE_ID === caseData.CASE_ID; }) || null;
    var updates = Object.assign({}, caseData);
    delete updates.CASE_ID;
    // Handle Date conversions
    if (updates.FILING_DATE) { updates.FILING_DATE = new Date(updates.FILING_DATE); }
    if (updates.NEXT_DEADLINE) { updates.NEXT_DEADLINE = new Date(updates.NEXT_DEADLINE); }
    var res = updateCase(caseData.CASE_ID, updates);
    if(res.success) {
      var changedAssignees = [
        { field: "ASSIGNED_STAFF_EMAIL", label: "Staff" },
        { field: "GALVANIZER_EMAIL", label: "Galvanizer" },
        { field: "ATTORNEY", label: "Attorney" }
      ];
      changedAssignees.forEach(function(item) {
        var nextValue = String(caseData[item.field] || "").trim().toLowerCase();
        var previousValue = String((existingCase && existingCase[item.field]) || "").trim().toLowerCase();
        if (nextValue && nextValue !== previousValue) {
          notifyAssignee_(caseData[item.field], item.label, caseData.CASE_ID, caseData.PATENT_TITLE || (existingCase && existingCase.PATENT_TITLE) || "");
        }
      });
      clearPortalCaches_(["cases", "casesPage", "dashboard", "dashboardSummary", "dashboardDetails", "documents", "galvanizerQueue"]);
      return { success: true, message: "Case updated successfully" };
    }
    return res;
  } else {
    var res = createCase({
      clientId: caseData.CLIENT_ID,
      patentTitle: caseData.PATENT_TITLE,
      applicationNumber: caseData.APPLICATION_NUMBER,
      country: caseData.COUNTRY,
      filingDate: caseData.FILING_DATE ? new Date(caseData.FILING_DATE) : "",
      nextDeadline: caseData.NEXT_DEADLINE ? new Date(caseData.NEXT_DEADLINE) : "",
      status: caseData.CURRENT_STATUS,
      patentType: caseData.PATENT_TYPE,
      attorney: caseData.ATTORNEY,
      priority: caseData.PRIORITY,
      orgId: caseData.ORG_ID || "",
      assignedStaffEmail: caseData.ASSIGNED_STAFF_EMAIL || "",
      galvanizerEmail: caseData.GALVANIZER_EMAIL || "",
      workflowStage: caseData.WORKFLOW_STAGE || "",
      notes: caseData.NOTES
    });
    if(res.success) {
      notifyAssignee_(caseData.ASSIGNED_STAFF_EMAIL || "", "Staff", res.caseId, caseData.PATENT_TITLE || "");
      notifyAssignee_(caseData.GALVANIZER_EMAIL || "", "Galvanizer", res.caseId, caseData.PATENT_TITLE || "");
      notifyAssignee_(caseData.ATTORNEY || "", "Attorney", res.caseId, caseData.PATENT_TITLE || "");
      clearPortalCaches_(["cases", "casesPage", "dashboard", "dashboardSummary", "dashboardDetails", "documents", "galvanizerQueue"]);
      return { success: true, message: "Case created successfully" };
    }
    return res;
  }
}

function deletePortalCase_(caseId) {
  var sheet = getSheet_("CASE");
  var data = sheet.getDataRange().getValues();
  var statusCol = data[0].indexOf("CURRENT_STATUS");
  if (statusCol === -1) statusCol = 7;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === caseId) {
      sheet.getRange(i + 1, statusCol + 1).setValue("Deleted");
      clearPortalCaches_(["cases", "casesPage", "dashboard", "dashboardSummary", "dashboardDetails", "invoices"]);
      return { success: true, message: "Case deleted safely" };
    }
  }
  return { error: "Case not found" };
}

function savePortalInvoice_(invoiceData) {
  var result = upsertCompanyInvoice_(invoiceData || {});
  if (result && result.success) {
    clearPortalCaches_(["invoices", "dashboard", "dashboardSummary", "dashboardDetails"]);
  }
  return result;
}

function deletePortalInvoice_(invoiceId) {
  var existing = getInvoiceById_(invoiceId);
  if (!existing) return { error: "Invoice not found" };
  var ok = updateRecordById_("INVOICE", "INVOICE_ID", existing.INVOICE_ID, {
    PAYMENT_STATUS: "Deleted"
  });
  if (ok) clearPortalCaches_(["invoices", "dashboard", "dashboardSummary", "dashboardDetails"]);
  return ok ? { success: true, message: "Invoice deleted" } : { error: "Invoice not found" };
}

// ── Register User Dialog (Admin) ────────────────────────────
function showRegisterUserDialog() {

  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; }
      h2 { color: #1a237e; }
      label { display: block; margin-top: 10px; font-weight: 500; }
      input, select { width: 100%; padding: 8px 12px; margin-top: 4px; border: 1px solid #ddd;
        border-radius: 6px; font-size: 14px; box-sizing: border-box; }
      .btn { background: #1a237e; color: white; padding: 10px 24px; border: none;
        border-radius: 6px; cursor: pointer; margin-top: 16px; }
      .btn:hover { background: #283593; }
      .btn-cancel { background: #757575; margin-left: 8px; }
      .required { color: #d32f2f; }
      #status { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }
    </style>
    <h2>👤 Register New User</h2>
    <label>Full Name <span class="required">*</span></label>
    <input id="fullName" placeholder="e.g. John Doe" />
    <label>Email <span class="required">*</span></label>
    <input id="email" type="email" placeholder="e.g. john@example.com" />
    <label>Password <span class="required">*</span></label>
    <input id="password" type="password" placeholder="Min 6 characters" />
    <label>Role <span class="required">*</span></label>
    <select id="role">
      <option value="Client">Client</option>
      <option value="Attorney">Attorney</option>
      <option value="Admin">Admin</option>
    </select>
    <label>Client ID (for Client role only)</label>
    <input id="clientId" placeholder="e.g. CL001" />
    <label>Department</label>
    <input id="department" placeholder="e.g. Patents, Legal" />
    <div>
      <button class="btn" onclick="submitUser()">Register User</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    <div id="status"></div>
    <script>
      function submitUser() {
        var data = {
          fullName: document.getElementById('fullName').value,
          email: document.getElementById('email').value,
          password: document.getElementById('password').value,
          role: document.getElementById('role').value,
          clientId: document.getElementById('clientId').value,
          department: document.getElementById('department').value
        };
        if (!data.fullName || !data.email || !data.password) {
          showMsg('Please fill all required fields.', 'error'); return;
        }
        if (data.password.length < 6) {
          showMsg('Password must be at least 6 characters.', 'error'); return;
        }
        showMsg('Registering user...', 'info');
        google.script.run
          .withSuccessHandler(function(r) { showMsg(r.message, r.success?'success':'error'); })
          .withFailureHandler(function(e) { showMsg('Error: ' + e.message, 'error'); })
          .registerUser(data);
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
  .setWidth(450)
  .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, "Register User");
}

/**
 * Register a new user (called from admin dialog)
 */
function registerUser(userData) {
  return withScriptLock_(function() {
  var sheet = getSheet_("USERS");
  var data = sheet.getDataRange().getValues();
  
  // Check duplicate email
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === userData.email) {
      return { success: false, message: "User with this email already exists." };
    }
  }
  
  var userId = generateId_("USR", sheet, 0);
  var passwordHash = userData.password ? hashPassword_(userData.password) : "";
  
  sheet.appendRow([
    userId,
    userData.email,
    userData.fullName,
    normalizeRole_(userData.role),
    userData.clientId || "",
    userData.orgId || "",
    userData.department || "",
    "Active",
    passwordHash,
    userData.canViewFinance || "No",
    userData.reportsTo || "",
    userData.additionalRoles || "",
    new Date()
  ]);
  
  logActivity_("REGISTER_USER", "USER", userId, "Registered: " + userData.email + " as " + userData.role);
  
  return { success: true, message: "User " + userId + " registered successfully." };
  });
}

// ── Dashboard data (filtered by role) ───────────────────────
function getPortalDashboard_(email, userRole, filters) {
  if (!userRole || userRole.role === "Unknown") {
    return { error: "Access denied. Your email is not registered in the system." };
  }
  var cacheKey = buildPortalCacheKey_("dashboard", userRole, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;

  try {
    var cases = filterCases_(getAccessibleCasesForUser_(userRole) || [], filters) || [];
    var invoices = filterInvoicesByCases_(getAccessibleInvoicesForUser_(userRole) || [], cases) || [];
  } catch (e) {
    Logger.log("Dashboard data fetch error: " + e.message);
    return { error: "Failed to fetch dashboard data: " + e.message };
  }

  // Aggregate Data for Charts
  var grantedByCountry = {};
  var pendingByCountry = {};
  var pendingByStatus = {};
  var totalGranted = 0;
  var totalPending = 0;

  var upcomingDeadlines = [];
  var recentActiveCases = [];

  var now = new Date();

  cases.forEach(function(c) {
    if (!c) return;
    var status = c.CURRENT_STATUS;
    var country = c.COUNTRY || "Unknown";
    
    // Categorize Granted vs Pending
    if (status === "Granted") {
      totalGranted++;
      grantedByCountry[country] = (grantedByCountry[country] || 0) + 1;
    } else if (["Abandoned", "Refused", "Lapsed"].indexOf(status) === -1) {
      // It is Pending/Active
      totalPending++;
      pendingByCountry[country] = (pendingByCountry[country] || 0) + 1;
      pendingByStatus[status] = (pendingByStatus[status] || 0) + 1;
      
      // Store for Recent Active Cases table
      var rCase = {
        docket: c.APPLICATION_NUMBER || c.CASE_ID,
        date: c.FILING_DATE ? Utilities.formatDate(new Date(c.FILING_DATE), Session.getScriptTimeZone(), "dd MMM yy") : "-",
        status: status,
        rawDate: c.FILING_DATE ? new Date(c.FILING_DATE).getTime() : 0,
        title: c.PATENT_TITLE
      };
      recentActiveCases.push(rCase);

      // Check upcoming deadlines
      if (c.NEXT_DEADLINE) {
        var dl = new Date(c.NEXT_DEADLINE);
        if (!isNaN(dl.getTime())) {
          upcomingDeadlines.push({
            docket: c.APPLICATION_NUMBER || c.CASE_ID,
            country: country,
            dateStr: Utilities.formatDate(dl, Session.getScriptTimeZone(), "dd MMM yy"),
            rawDate: dl.getTime(),
            id: c.CASE_ID,
            title: c.PATENT_TITLE
          });
        }
      }
    }
  });

  // Sort upcoming deadlines
  upcomingDeadlines.sort(function(a, b) { return a.rawDate - b.rawDate; });
  upcomingDeadlines = upcomingDeadlines.slice(0, 3); // Top 3

  // Sort recent active cases by newest filing
  recentActiveCases.sort(function(a, b) { return b.rawDate - a.rawDate; });
  recentActiveCases = recentActiveCases.slice(0, 5); // Top 5

  // Collect Pending Invoices
  var pendingInvoicesList = [];
  if (Array.isArray(invoices) && invoices.length > 0) {
    invoices.forEach(function(inv) {
      if (inv && isInvoiceOutstanding_(inv)) {
        var invDate = getInvoiceDateValue_(inv) ? new Date(getInvoiceDateValue_(inv)) : new Date(0);
        pendingInvoicesList.push({
          docket: getInvoiceDisplayNumber_(inv),
          dateStr: isNaN(invDate.getTime()) ? "-" : Utilities.formatDate(invDate, Session.getScriptTimeZone(), "dd MMM yy"),
          amount: getInvoiceAmountDueValue_(inv),
          rawDate: invDate.getTime()
        });
      }
    });
  }

  // Sort invoices by oldest first
  pendingInvoicesList.sort(function(a, b) { return a.rawDate - b.rawDate; });
  pendingInvoicesList = pendingInvoicesList.slice(0, 2); // Top 2 like screenshot

  return writePortalCache_(cacheKey, {
    role: userRole.role,
    name: userRole.name,
    totalGranted: totalGranted,
    totalPending: totalPending,
    grantedByCountry: grantedByCountry,
    pendingByCountry: pendingByCountry,
    pendingByStatus: pendingByStatus,
    upcomingDeadlines: upcomingDeadlines,
    recentActiveCases: recentActiveCases,
    pendingInvoicesList: pendingInvoicesList,
    unreadNotifications: getNotifications_(userRole).filter(function(item) { return item.IS_READ !== "Yes"; }).length,
    openThreads: getMessageThreads_(userRole).filter(function(item) { return item.STATUS !== "Closed"; }).length,
    myClientCount: getAccessibleClientsForUser_(userRole).length
  }, 60);
}

function buildDashboardBaseData_(userRole, filters) {
  var cases = filterCases_(getAccessibleCasesForUser_(userRole) || [], filters) || [];
  var invoices = filterInvoicesByCases_(getAccessibleInvoicesForUser_(userRole) || [], cases) || [];
  return { cases: cases, invoices: invoices };
}

function buildDashboardSummaryPayload_(userRole, data) {
  var cases = data.cases || [];
  var invoices = data.invoices || [];
  var totalGranted = 0;
  var totalPending = 0;
  var upcomingDeadlineCount = 0;
  var now = new Date();
  var inThirtyDays = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000));

  cases.forEach(function(c) {
    if (!c) return;
    var status = c.CURRENT_STATUS;
    if (status === "Granted") {
      totalGranted++;
    } else if (["Abandoned", "Refused", "Lapsed"].indexOf(status) === -1) {
      totalPending++;
      if (c.NEXT_DEADLINE) {
        var dl = new Date(c.NEXT_DEADLINE);
        if (!isNaN(dl.getTime()) && dl >= now && dl <= inThirtyDays) upcomingDeadlineCount++;
      }
    }
  });

  var pendingInvoicesCount = invoices.filter(function(inv) {
    return inv && isInvoiceOutstanding_(inv);
  }).length;

  return {
    role: userRole.role,
    name: userRole.name,
    totalGranted: totalGranted,
    totalPending: totalPending,
    upcomingDeadlineCount: upcomingDeadlineCount,
    pendingInvoicesCount: pendingInvoicesCount,
    unreadNotifications: getNotifications_(userRole).filter(function(item) { return item.IS_READ !== "Yes"; }).length,
    openThreads: getMessageThreads_(userRole).filter(function(item) { return item.STATUS !== "Closed"; }).length,
    myClientCount: getAccessibleClientsForUser_(userRole).length
  };
}

function canUsePrecomputedDashboardSummary_(userRole, filters) {
  var hasFilters = filters && Object.keys(filters).some(function(key) {
    return String(filters[key] || "").trim() !== "";
  });
  return !hasFilters && canManageAllData_(userRole);
}

function getPrecomputedDashboardSummarySnapshot_(userRole) {
  if (!canUsePrecomputedDashboardSummary_(userRole, {})) return null;
  var snapshotKey = "DASHBOARD_SUMMARY_SNAPSHOT_DEFAULT";
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(snapshotKey);
  if (!raw) return null;
  try {
    var parsed = JSON.parse(raw);
    if (parsed.version !== getPortalCacheVersion_("dashboardSummary")) return null;
    return parsed.payload || null;
  } catch (e) {
    return null;
  }
}

function savePrecomputedDashboardSummarySnapshot_(payload) {
  try {
    PropertiesService.getScriptProperties().setProperty("DASHBOARD_SUMMARY_SNAPSHOT_DEFAULT", JSON.stringify({
      version: getPortalCacheVersion_("dashboardSummary"),
      savedAt: new Date().getTime(),
      payload: payload
    }));
  } catch (e) {}
}

function buildDashboardDetailsPayload_(data) {
  var cases = data.cases || [];
  var invoices = data.invoices || [];
  var grantedByCountry = {};
  var pendingByCountry = {};
  var pendingByStatus = {};
  var totalGranted = 0;
  var totalPending = 0;
  var upcomingDeadlines = [];
  var recentActiveCases = [];
  var pendingInvoicesList = [];

  cases.forEach(function(c) {
    if (!c) return;
    var status = c.CURRENT_STATUS;
    var country = c.COUNTRY || "Unknown";

    if (status === "Granted") {
      totalGranted++;
      grantedByCountry[country] = (grantedByCountry[country] || 0) + 1;
    } else if (["Abandoned", "Refused", "Lapsed"].indexOf(status) === -1) {
      totalPending++;
      pendingByCountry[country] = (pendingByCountry[country] || 0) + 1;
      pendingByStatus[status] = (pendingByStatus[status] || 0) + 1;

      recentActiveCases.push({
        docket: c.APPLICATION_NUMBER || c.CASE_ID,
        date: c.FILING_DATE ? Utilities.formatDate(new Date(c.FILING_DATE), Session.getScriptTimeZone(), "dd MMM yy") : "-",
        status: status,
        rawDate: c.FILING_DATE ? new Date(c.FILING_DATE).getTime() : 0,
        title: c.PATENT_TITLE,
        CASE_ID: c.CASE_ID
      });

      if (c.NEXT_DEADLINE) {
        var dl = new Date(c.NEXT_DEADLINE);
        if (!isNaN(dl.getTime())) {
          upcomingDeadlines.push({
            docket: c.APPLICATION_NUMBER || c.CASE_ID,
            country: country,
            dateStr: Utilities.formatDate(dl, Session.getScriptTimeZone(), "dd MMM yy"),
            rawDate: dl.getTime(),
            id: c.CASE_ID,
            title: c.PATENT_TITLE
          });
        }
      }
    }
  });

  upcomingDeadlines.sort(function(a, b) { return a.rawDate - b.rawDate; });
  recentActiveCases.sort(function(a, b) { return b.rawDate - a.rawDate; });

  invoices.forEach(function(inv) {
    if (inv && isInvoiceOutstanding_(inv)) {
      var invDateValue = getInvoiceDateValue_(inv);
      var invDate = invDateValue ? new Date(invDateValue) : new Date(0);
      pendingInvoicesList.push({
        docket: getInvoiceDisplayNumber_(inv),
        dateStr: isNaN(invDate.getTime()) ? "-" : Utilities.formatDate(invDate, Session.getScriptTimeZone(), "dd MMM yy"),
        amount: getInvoiceAmountDueValue_(inv),
        rawDate: invDate.getTime()
      });
    }
  });

  pendingInvoicesList.sort(function(a, b) { return a.rawDate - b.rawDate; });

  return {
    totalGranted: totalGranted,
    totalPending: totalPending,
    grantedByCountry: grantedByCountry,
    pendingByCountry: pendingByCountry,
    pendingByStatus: pendingByStatus,
    upcomingDeadlines: upcomingDeadlines.slice(0, 3),
    recentActiveCases: recentActiveCases.slice(0, 5),
    pendingInvoicesList: pendingInvoicesList.slice(0, 2)
  };
}

function getPortalDashboardSummary_(email, userRole, filters) {
  if (!userRole || userRole.role === "Unknown") {
    return { error: "Access denied. Your email is not registered in the system." };
  }
  var cacheKey = buildPortalCacheKey_("dashboardSummary", userRole, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var precomputed = canUsePrecomputedDashboardSummary_(userRole, filters) ? getPrecomputedDashboardSummarySnapshot_(userRole) : null;
  if (precomputed) return writePortalCache_(cacheKey, precomputed, 60);
  try {
    var data = buildDashboardBaseData_(userRole, filters);
    var payload = buildDashboardSummaryPayload_(userRole, data);
    if (canUsePrecomputedDashboardSummary_(userRole, filters)) {
      savePrecomputedDashboardSummarySnapshot_(payload);
    }
    return writePortalCache_(cacheKey, payload, 60);
  } catch (e) {
    Logger.log("Dashboard summary fetch error: " + e.message);
    return { error: "Failed to fetch dashboard summary: " + e.message };
  }
}

function getPortalDashboardDetails_(email, userRole, filters) {
  if (!userRole || userRole.role === "Unknown") {
    return { error: "Access denied. Your email is not registered in the system." };
  }
  var cacheKey = buildPortalCacheKey_("dashboardDetails", userRole, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  try {
    var data = buildDashboardBaseData_(userRole, filters);
    return writePortalCache_(cacheKey, buildDashboardDetailsPayload_(data), 60);
  } catch (e) {
    Logger.log("Dashboard details fetch error: " + e.message);
    return { error: "Failed to fetch dashboard details: " + e.message };
  }
}

function sanitizeDataForFrontend_(dataArray) {
  if (dataArray == null) return dataArray;
  if (Array.isArray(dataArray)) {
    return dataArray.map(function(item) {
      return sanitizeDataForFrontend_(item);
    });
  }
  if (Object.prototype.toString.call(dataArray) === "[object Object]") {
    var sanitized = {};
    Object.keys(dataArray).forEach(function(key) {
      sanitized[key] = sanitizeDataForFrontend_(dataArray[key]);
    });
    return sanitized;
  }
  if (dataArray instanceof Date) {
    return Utilities.formatDate(dataArray, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  }
  return dataArray;
}

function getPortalCases_(email, userRole, filters) {
  var cacheKey = buildPortalCacheKey_("cases", userRole, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(filterCases_(getAccessibleCasesForUser_(userRole), filters)), 60);
}

function getPortalCasesPage_(email, userRole, filters, limit, offset) {
  var safeLimit = Math.max(1, Math.min(parseInt(limit, 10) || 50, 100));
  var safeOffset = Math.max(0, parseInt(offset, 10) || 0);
  var cacheKey = buildPortalCacheKey_("casesPage", userRole, JSON.stringify({
    filters: filters || {},
    limit: safeLimit,
    offset: safeOffset
  }));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;

  var allCases = filterCases_(getAccessibleCasesForUser_(userRole), filters);
  var total = allCases.length;
  var items = allCases.slice(safeOffset, safeOffset + safeLimit);
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_({
    items: items,
    total: total,
    limit: safeLimit,
    offset: safeOffset,
    nextOffset: safeOffset + items.length,
    hasMore: (safeOffset + items.length) < total
  }), 60);
}

function getInvoiceDateValue_(invoice) {
  return invoice["Invoice Date"] || invoice["Tax Invoice Date"] || invoice.INVOICE_DATE || "";
}

function getInvoiceDisplayNumber_(invoice) {
  return invoice["Docket# [Invoice UIN]"] || invoice["Tax Invoice Number"] || invoice.INVOICE_ID || "";
}

function getInvoiceClientCode_(invoice) {
  return invoice.ClientCode || invoice.CLIENT_ID || "";
}

function getInvoiceServiceSummary_(invoice) {
  return [
    invoice["Main ServiceCode"],
    invoice["Service Code 2"],
    invoice["Service Code 3"]
  ].filter(function(item) { return !!String(item || "").trim(); }).join(" / ");
}

function getInvoiceTotalValue_(invoice) {
  return parseFloat(invoice.InvAmount || invoice.TOTAL || 0) || 0;
}

function getInvoiceAmountDueValue_(invoice) {
  if (invoice.hasOwnProperty("Amount due")) return parseFloat(invoice["Amount due"] || 0) || 0;
  return String(invoice.PAYMENT_STATUS || "") === "Paid" ? 0 : getInvoiceTotalValue_(invoice);
}

function isInvoiceOutstanding_(invoice) {
  return String(invoice.PAYMENT_STATUS || "") !== "Paid" && String(invoice.PAYMENT_STATUS || "") !== "Deleted" && getInvoiceAmountDueValue_(invoice) > 0;
}

function getPortalInvoices_(email, userRole) {
  if (!canViewFinance_(userRole)) return { error: "Access denied." };
  return sanitizeDataForFrontend_(getAccessibleInvoicesForUser_(userRole));
}

function getPortalDocuments_(email, userRole) {
  return getAccessibleDocumentsForUser_(userRole);
}

function submitContactForm_(email, params) {
  var subject = "Client Portal Message: " + (params.subject || "No Subject");
  var body = "From: " + email + "\n" +
    "Case ID: " + (params.caseId || "N/A") + "\n\n" +
    (params.message || "No message");

  var adminEmail = Session.getEffectiveUser().getEmail();

  try {
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body,
      replyTo: email
    });
    logActivity_("CONTACT_FORM", "CLIENT", "", "Subject: " + params.subject);
    return { success: true, message: "Message sent successfully!" };
  } catch (e) {
    return { success: false, message: "Failed to send message: " + e.message };
  }
}

/**
 * Generate the full portal HTML with login page
 */
function getPortalHTML_() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>IP Patent Firm - Client Portal</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Material+Symbols+Outlined" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    :root {
      --primary: #2196f3;
      --primary-dark: #1976d2;
      --primary-light: #e3f2fd;
      --success: #4caf50;
      --warning: #ff9800;
      --danger: #f44336;
      --bg: #f4f7f6;
      --card: #ffffff;
      --text: #37474f;
      --text-muted: #78909c;
      --border: #eceff1;
      --radius: 8px;
      --shadow: 0 4px 15px rgba(0,0,0,0.03);
    }
    body { font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text); }

    /* ── Login Page ────────────────────────────── */
    .login-container {
      display: flex; align-items: center; justify-content: center;
      min-height: 100vh; background: linear-gradient(135deg, #1976d2 0%, #2196f3 50%, #64b5f6 100%);
    }
    .login-card {
      background: white; border-radius: 12px; padding: 48px 40px; width: 100%; max-width: 420px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.2); text-align: center;
    }
    .login-card .logo-icon { font-size: 48px; margin-bottom: 8px; }
    .login-card h1 { font-size: 24px; color: var(--primary-dark); margin-bottom: 4px; }
    .login-card p { font-size: 14px; color: var(--text-muted); margin-bottom: 28px; }
    .login-card .form-group { text-align: left; margin-bottom: 16px; }
    .login-card label { display: block; font-size: 13px; font-weight: 500; margin-bottom: 4px; color: #555; }
    .login-card input {
      width: 100%; padding: 12px 14px; border: 1.5px solid var(--border); border-radius: 8px;
      font-size: 14px; font-family: inherit; transition: border-color 0.2s;
    }
    .login-card input:focus { border-color: var(--primary); outline: none; box-shadow: 0 0 0 3px rgba(33,150,243,0.1); }
    .login-btn {
      width: 100%; padding: 12px; background: var(--primary-dark); color: white; border: none;
      border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer;
      margin-top: 8px; transition: background 0.2s;
    }
    .login-btn:hover { background: var(--primary); }
    .login-btn:disabled { opacity: 0.6; cursor: not-allowed; }
    .login-error { background: #ffebee; color: #c62828; padding: 10px; border-radius: 8px; font-size: 13px; margin-top: 12px; display: none; }
    .login-footer { margin-top: 24px; font-size: 12px; color: var(--text-muted); }

    /* ── App Layout ────────────────────────────── */
    .app-container { display: none; }
    .sidebar {
      position: fixed; left: 0; top: 0; bottom: 0; width: 240px;
      background: white; color: var(--text); padding: 24px 0; z-index: 100;
      border-right: 1px solid var(--border);
      transition: transform 0.3s;
    }
    .sidebar .logo { padding: 0 24px 24px; font-size: 18px; font-weight: 700; color: var(--primary-dark); display: flex; align-items: center; gap: 8px;}
    .sidebar .logo img { width: 24px; height: 24px; }
    .sidebar nav { padding: 16px 0; }
    .sidebar nav a {
      display: flex; align-items: center; gap: 12px; padding: 12px 24px; color: var(--text-muted);
      text-decoration: none; font-size: 14px; font-weight: 500; transition: all 0.2s; cursor: pointer;
    }
    .sidebar nav a:hover { background: var(--bg); color: var(--primary-dark); }
    .sidebar nav a.active { background: var(--primary-light); color: var(--primary-dark); border-right: 3px solid var(--primary); }
    .sidebar nav a .material-symbols-outlined { font-size: 20px; }
    .sidebar .logout-btn {
      position: absolute; bottom: 16px; left: 16px; right: 16px;
      display: flex; align-items: center; gap: 8px; padding: 10px 16px;
      background: var(--bg); border: none; color: var(--text-muted);
      border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: 500; font-family: inherit;
    }
    .sidebar .logout-btn:hover { background: var(--border); color: var(--primary-dark); }

    .main { margin-left: 240px; padding: 24px; min-height: 100vh; }
    .topbar {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 24px; padding-bottom: 16px; 
    }
    .topbar h1 { font-size: 22px; font-weight: 600; color: transparent; /* Title hidden visually as per mock */ }
    .topbar .user-info { font-size: 13px; color: var(--text-muted); display: flex; align-items: center; gap: 16px; background: white; padding: 8px 16px; border-radius: 20px; box-shadow: var(--shadow); }

    /* ── Dashboard Grid ────────────────────────────── */
    .dash-row { display: flex; gap: 20px; margin-bottom: 20px; flex-wrap: wrap; }
    .col-7 { flex: 7; min-width: 500px; }
    .col-5 { flex: 5; min-width: 300px; }
    .col-6 { flex: 6; min-width: 400px; }
    .col-half { flex: 1; min-width: 400px; }

    .card { background: var(--card); border-radius: var(--radius); padding: 24px; box-shadow: var(--shadow); height: 100%; border: 1px solid var(--border); }
    
    .chart-header { display: flex; justify-content: space-around; width: 100%; margin-bottom: 24px; }
    .chart-header h3 { font-size: 14px; font-weight: 600; color: var(--primary-dark); }
    
    .donut-container { position: relative; width: 160px; height: 160px; margin: 0 auto; }
    .donut-center-text { 
      position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); 
      font-size: 28px; font-weight: 400; color: var(--primary-dark); 
    }
    
    .legend-row { display: flex; justify-content: center; gap: 16px; margin-top: 24px; font-size: 12px; color: var(--text-muted); }
    .legend-item { display: flex; align-items: center; gap: 6px; }
    .legend-dot { width: 10px; height: 10px; border-radius: 50%; }

    /* Lists (Renewals & Invoices) */
    .list-card h3 { font-size: 14px; font-weight: 600; color: var(--primary-dark); margin-bottom: 16px; }
    .list-item { 
      background: #fdfdfd; border: 1px solid var(--border); border-radius: 6px; 
      padding: 12px; margin-bottom: 12px; border-left: 4px solid var(--primary-light);
      display: flex; flex-direction: column; gap: 8px;
    }
    .list-item.accent-blue { border-left-color: var(--primary); }
    .list-item.accent-dark { border-left-color: var(--primary-dark); }
    .list-item-row { display: flex; justify-content: space-between; align-items: center; font-size: 12px; }
    .list-item-title { font-weight: 500; color: var(--text); display: flex; align-items: center; gap: 4px; }
    .list-item-meta { color: var(--text-muted); display: flex; align-items: center; gap: 4px;}
    .list-item-highlight { background: var(--primary-light); color: var(--primary-dark); padding: 6px 12px; border-radius: 4px; font-weight: 600; display:flex; justify-content: space-between;}
    .view-all-link { text-align: right; font-size: 12px; font-weight: 500; color: var(--text-muted); cursor: pointer; margin-top: 8px; display: block; }
    
    /* Tables */
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th { background: #fdfdfd; color: var(--text-muted); padding: 12px; text-align: left; font-weight: 500; border-bottom: 1px solid var(--border); }
    td { padding: 14px 12px; border-bottom: 1px solid var(--border); color: var(--text); }
    tr:last-child td { border-bottom: none; }
    
    .status-text { color: var(--text); font-weight: 500; }
    .action-text { color: var(--text-muted); }



    .badge {
      display: inline-block; padding: 3px 10px; border-radius: 20px;
      font-size: 11px; font-weight: 600; text-transform: uppercase;
    }
    .badge-filed { background: #e3f2fd; color: var(--primary-dark); }
    .badge-granted { background: #e8f5e9; color: var(--success); }
    .badge-examination { background: #fff3e0; color: var(--warning); }
    .badge-drafted { background: #f3e5f5; color: #7b1fa2; }
    .badge-paid { background: #e8f5e9; color: var(--success); }
    .badge-unpaid { background: #ffebee; color: var(--danger); }

    .doc-section { margin-bottom: 20px; }
    .doc-section h3 { font-size: 14px; color: var(--primary-dark); margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
    .doc-item {
      display: flex; align-items: center; gap: 12px; padding: 10px 12px;
      border: 1px solid var(--border); border-radius: 8px; margin-bottom: 8px; transition: all 0.2s;
    }
    .doc-item:hover { border-color: var(--primary); background: var(--primary-light); }
    .doc-item a { color: var(--primary-dark); text-decoration: none; font-weight: 500; font-size: 13px; }
    .doc-item .meta { font-size: 11px; color: var(--text-muted); }

    .form-group { margin-bottom: 16px; }
    .form-group label { display: block; font-size: 13px; font-weight: 500; margin-bottom: 4px; }
    .form-group input, .form-group textarea, .form-group select {
      width: 100%; padding: 10px 14px; border: 1px solid var(--border); border-radius: 8px;
      font-size: 14px; font-family: inherit;
    }
    .form-group input:focus, .form-group textarea:focus { border-color: var(--primary); outline: none; box-shadow: 0 0 0 3px rgba(33,150,243,0.1); }
    .btn-primary {
      background: var(--primary-dark); color: white; border: none; padding: 10px 28px;
      border-radius: 8px; font-size: 14px; font-weight: 500; cursor: pointer;
    }
    .btn-primary:hover { background: var(--primary); }

    .loading { text-align: center; padding: 40px; color: var(--text-muted); width: 100%; }
    .spinner { display: inline-block; width: 32px; height: 32px; border: 3px solid var(--border);
      border-top-color: var(--primary-dark); border-radius: 50%; animation: spin 0.8s linear infinite; }
    @keyframes spin { to { transform: rotate(360deg); } }

    .alert-msg { padding: 12px 16px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; }
    .alert-error { background: #ffebee; color: #c62828; }
    .alert-success { background: #e8f5e9; color: #2e7d32; }

    .menu-toggle { display: none; position: fixed; top: 16px; left: 16px; z-index: 200;
      background: white; color: var(--text); border: 1px solid var(--border); padding: 8px 12px; border-radius: 8px; cursor: pointer; }
    @media (max-width: 768px) {
      .sidebar { transform: translateX(-100%); }
      .sidebar.open { transform: translateX(0); box-shadow: var(--shadow); }
      .main { margin-left: 0; }
      .menu-toggle { display: block; }
      .dash-row { flex-direction: column; }
      .col-7, .col-5, .col-6, .col-half { width: 100%; min-width: unset; }
      .login-card { margin: 16px; padding: 32px 24px; }
    }

    .page { display: none; }
    .page.active { display: block; }
  </style>
</head>
<body>

  <!-- ═══════ LOGIN PAGE ═══════ -->
  <div class="login-container" id="loginPage">
    <div class="login-card">
      <div class="logo-icon">⚖️</div>
      <h1>IP Patent Firm</h1>
      <p>Client Portal — Secure Login</p>
      <div class="form-group">
        <label>Email Address</label>
        <input id="loginEmail" type="email" placeholder="your@email.com" onkeydown="if(event.key==='Enter')doLogin()" />
      </div>
      <div class="form-group">
        <label>Password</label>
        <input id="loginPassword" type="password" placeholder="Enter your password" onkeydown="if(event.key==='Enter')doLogin()" />
      </div>
      <button class="login-btn" id="loginBtn" onclick="doLogin()">Sign In</button>
      <div class="login-error" id="loginError"></div>
      <div class="login-footer">🔒 Secured access for registered clients & attorneys</div>
    </div>
  </div>

  <!-- ═══════ APP (hidden until login) ═══════ -->
  <div class="app-container" id="appContainer">
    <button class="menu-toggle" onclick="document.querySelector('.sidebar').classList.toggle('open')">☰</button>

    <aside class="sidebar">
      <div class="logo">
        Metayage
      </div>
      <nav>
        <a onclick="showPage('dashboard')" class="active" data-page="dashboard">
          <span class="material-symbols-outlined">home</span> Home
        </a>
        <a onclick="showPage('cases')" data-page="cases">
          <span class="material-symbols-outlined">workspace_premium</span> Cases
        </a>
        <a onclick="showPage('documents')" data-page="documents">
          <span class="material-symbols-outlined">description</span> Documents
        </a>
        <a onclick="showPage('invoices')" data-page="invoices">
          <span class="material-symbols-outlined">attach_money</span> Finance
        </a>
        <a id="nav-management" onclick="showPage('management')" data-page="management" style="display:none;">
          <span class="material-symbols-outlined">admin_panel_settings</span> Management
        </a>
        <a onclick="showPage('settings')" data-page="settings">
          <span class="material-symbols-outlined">vpn_key</span> Change Password
        </a>
        <a onclick="showPage('contact')" data-page="contact">
          <span class="material-symbols-outlined">mail</span> Contact
        </a>
      </nav>
      <button class="logout-btn" onclick="doLogout()">
        <span class="material-symbols-outlined">logout</span> Sign Out
      </button>
    </aside>

    <main class="main">
      <div class="topbar">
        <h1 id="pageTitle">Dashboard</h1>
        <div class="user-info" id="userInfo">Loading...</div>
      </div>

      <div class="page active" id="page-dashboard">
        <div class="loading" id="dashboardLoading"><div class="spinner"></div><p>Loading dashboard...</p></div>
        <div id="dashboardContent" style="display:none;"></div>
      </div>
      <div class="page" id="page-cases">
        <div class="loading" id="casesLoading"><div class="spinner"></div><p>Loading cases...</p></div>
        <div id="casesContent" style="display:none;"></div>
      </div>
      <div class="page" id="page-documents">
        <div class="loading" id="documentsLoading"><div class="spinner"></div><p>Loading documents...</p></div>
        <div id="documentsContent" style="display:none;"></div>
      </div>
      <div class="page" id="page-invoices">
        <div class="loading" id="invoicesLoading"><div class="spinner"></div><p>Loading invoices...</p></div>
        <div id="invoicesContent" style="display:none;"></div>
      </div>
      <div class="page" id="page-management">
        <div class="topbar" style="margin-bottom:16px;">
          <h2 style="font-size:18px;">Admin Management Control</h2>
        </div>
        <div style="display:flex; gap:16px; margin-bottom: 24px;">
          <button class="btn-primary" onclick="loadManageUsers()">Manage Users</button>
          <button class="btn-primary" onclick="loadManageClients()">Manage Clients</button>
          <button class="btn-primary" onclick="loadManageCases()">Manage Cases</button>
        </div>
        <div class="loading" id="managementLoading" style="display:none;"><div class="spinner"></div><p>Loading data...</p></div>
        <div id="managementContent" style="display:none;"></div>
      </div>
      <div class="page" id="page-contact">
        <div class="card">
          <h2>📧 Contact Your Firm</h2>
          <div class="form-group"><label>Subject</label><input id="contactSubject" placeholder="What is this regarding?" /></div>
          <div class="form-group"><label>Related Case ID (optional)</label><input id="contactCaseId" placeholder="e.g. CASE0001" /></div>
          <div class="form-group"><label>Message</label><textarea id="contactMessage" rows="5" placeholder="Type your message here..."></textarea></div>
          <button class="btn-primary" onclick="submitContact()">Send Message</button>
          <div id="contactStatus"></div>
        </div>
      </div>
      <div class="page" id="page-settings">
        <div class="card">
          <h2>🔐 Change Password</h2>
          <div class="form-group"><label>Current Password</label><input id="oldPassword" type="password" /></div>
          <div class="form-group"><label>New Password (min 6 chars)</label><input id="newPassword" type="password" /></div>
          <div class="form-group"><label>Confirm New Password</label><input id="confirmPassword" type="password" /></div>
          <button class="btn-primary" onclick="doChangePassword()">Update Password</button>
          <div id="passwordStatus" style="margin-top:12px;"></div>
        </div>
      </div>
    </main>
  </div>

  <script>
    var SESSION_TOKEN = null;

    // ── Login ──────────────────────────────────
    function doLogin() {
      var email = document.getElementById('loginEmail').value.trim();
      var password = document.getElementById('loginPassword').value;
      if (!email || !password) { showLoginError('Please enter email and password.'); return; }
      
      document.getElementById('loginBtn').disabled = true;
      document.getElementById('loginBtn').textContent = 'Signing in...';
      
      google.script.run
        .withSuccessHandler(function(r) {
          if (r.success) {
            SESSION_TOKEN = r.token;
            document.getElementById('loginPage').style.display = 'none';
            document.getElementById('appContainer').style.display = 'block';
            document.getElementById('userInfo').textContent = r.name + ' (' + r.role + ')';
            if (r.role === 'Admin') {
              document.getElementById('nav-management').style.display = 'flex';
            } else {
              document.getElementById('nav-management').style.display = 'none';
            }
            loadDashboard();
          } else {
            showLoginError(r.message);
            document.getElementById('loginBtn').disabled = false;
            document.getElementById('loginBtn').textContent = 'Sign In';
          }
        })
        .withFailureHandler(function(e) {
          showLoginError('Connection error. Please try again.');
          document.getElementById('loginBtn').disabled = false;
          document.getElementById('loginBtn').textContent = 'Sign In';
        })
        .getPortalData('login', { email: email, password: password });
    }

    function showLoginError(msg) {
      var el = document.getElementById('loginError');
      el.textContent = msg;
      el.style.display = 'block';
    }

    function doLogout() {
      google.script.run.getPortalData('logout', { token: SESSION_TOKEN });
      SESSION_TOKEN = null;
      document.getElementById('appContainer').style.display = 'none';
      document.getElementById('loginPage').style.display = 'flex';
      document.getElementById('loginPassword').value = '';
      document.getElementById('loginError').style.display = 'none';
      document.getElementById('loginBtn').disabled = false;
      document.getElementById('loginBtn').textContent = 'Sign In';
    }

    // ── Handle session expiry ────────────────
    function handleResponse(data, containerId, callback) {
      if (!data) {
        var errContainer = typeof containerId === 'string' ? containerId : 'dashboardContent';
        showError(errContainer, 'Server returned no data. There may be a data loading error.');
        return;
      }
      if (data.sessionExpired) {
        alert('Session expired. Please login again.');
        doLogout();
        return;
      }
      var cb = callback || containerId;
      if (typeof cb === 'function') {
        cb(data);
      }
    }

    // ── Page Navigation ──────────────────────
    function showPage(page) {
      document.querySelectorAll('.page').forEach(function(p) { p.classList.remove('active'); });
      document.querySelectorAll('nav a').forEach(function(a) { a.classList.remove('active'); });
      document.getElementById('page-' + page).classList.add('active');
      var navLink = document.querySelector('[data-page="' + page + '"]');
      if (navLink) navLink.classList.add('active');

      var titles = { dashboard:'Dashboard', cases:'My Cases', documents:'Documents', invoices:'Invoices', management:'System Management', contact:'Contact Firm', settings:'Settings' };
      document.getElementById('pageTitle').textContent = titles[page] || page;

      document.querySelector('.sidebar').classList.remove('open');

      if (page === 'dashboard') loadDashboard();
      if (page === 'cases') loadCases();
      if (page === 'documents') loadDocuments();
      if (page === 'invoices') loadInvoices();
      if (page === 'management') loadManageUsers(); // default sub-tab
    }

    // ── Dashboard Layout & Charts ────────────────────────
    var chartInstances = {};

    function loadDashboard() {
      google.script.run
        .withSuccessHandler(function(data) {
          handleResponse(data, 'dashboardContent', function(d) {
            if (d.error) { showError('dashboardContent', d.error); return; }

            // Construct new complex HTML
            var html = '';
            
            // Top Row
            html += '<div class="dash-row">';
            // 1. Double Donut Charts
            html += '<div class="col-7"><div class="card">';
            html += '<div class="chart-header"><h3>Granted Patents</h3><h3>Pending Patents</h3></div>';
            html += '<div style="display:flex; justify-content:space-around; align-items:center;">';
            html += '<div class="donut-container"><canvas id="grantedChart"></canvas><div class="donut-center-text">'+d.totalGranted+'</div></div>';
            html += '<div class="donut-container"><canvas id="pendingChart"></canvas><div class="donut-center-text">'+d.totalPending+'</div></div>';
            html += '</div>';
            // Custom Legend
            html += '<div class="legend-row">';
            var countryColors = {'India': '#1a237e', 'USA': '#4dd0e1', 'Japan': '#80cbc4', 'EPO': '#f4511e', 'Korea': '#9e9e9e', 'Other': '#e0e0e0'};
            var allCountries = Object.keys(Object.assign({}, d.grantedByCountry, d.pendingByCountry));
            allCountries.filter(function(v, i, a) { return a.indexOf(v) === i; }).forEach(function(c) {
              html += '<div class="legend-item"><div class="legend-dot" style="background:'+(countryColors[c]||'#bdbdbd')+'"></div>'+c+'</div>';
            });
            html += '</div>';
            html += '</div></div>'; // end col-7 -> card

            // 2. Lists Stack
            html += '<div class="col-5" style="display:flex; flex-direction:column; gap:20px;">';
            
            // Renewals
            html += '<div class="card list-card"><h3>Upcoming Renewals</h3>';
            if (d.upcomingDeadlines.length === 0) { html += '<p style="font-size:12px;color:#999;">No upcoming deadlines.</p>'; }
            else {
              d.upcomingDeadlines.forEach(function(u) {
                html += '<div class="list-item accent-blue">';
                html += '<div class="list-item-row"><div class="list-item-title"><span class="material-symbols-outlined" style="font-size:16px;">bookmark</span>'+u.docket+'</div><div class="list-item-meta"><span class="material-symbols-outlined" style="font-size:14px;">location_on</span>'+u.country+'</div></div>';
                html += '<div class="list-item-row" style="color:var(--text-muted); margin-top:8px;">Due Date <span>'+u.dateStr+'</span></div>';
                html += '</div>';
              });
            }
            html += '</div>'; // End Renewals
            
            // Invoices
            html += '<div class="card list-card"><h3>Pending payments</h3>';
            if (d.pendingInvoicesList.length === 0) { html += '<p style="font-size:12px;color:#999;">No pending payments.</p>'; }
            else {
              d.pendingInvoicesList.forEach(function(inv) {
                html += '<div class="list-item accent-dark">';
                html += '<div class="list-item-row"><div class="list-item-meta">Docket No.<br><span style="color:var(--text);font-weight:500;">'+inv.docket+'</span></div><div class="list-item-meta" style="text-align:right;">Invoice Date<br><span style="color:var(--text);font-weight:500;">'+inv.dateStr+'</span></div></div>';
                html += '<div class="list-item-highlight"><span style="display:flex;align-items:center;gap:4px;"><span class="material-symbols-outlined" style="font-size:16px;">payments</span> Amount Due</span><span>INR '+(inv.amount).toLocaleString('en-IN')+'</span></div>';
                html += '</div>';
              });
            }
            html += '<div class="view-all-link" onclick="showPage(\\'invoices\\')">View All →</div>';
            html += '</div>'; // End Invoices
            
            html += '</div>'; // End col-5 stack
            html += '</div>'; // End Top Row
            
            // Bottom Row
            html += '<div class="dash-row">';
            // Bar Chart
            html += '<div class="col-half"><div class="card">';
            html += '<h3 style="font-size:14px; color:var(--primary-dark); text-align:center; margin-bottom: 24px;">Pending Patents by status</h3>';
            html += '<div style="position:relative; height: 260px; width: 100%;"><canvas id="barChart"></canvas></div>';
            html += '</div></div>';
            
            // Action Table
            html += '<div class="col-half"><div class="card list-card">';
            html += '<h3 style="text-align:center; margin-bottom:16px;">Action Required</h3>';
            html += '<table><thead><tr><th>Docket No.</th><th>Filing Date</th><th>Status</th><th>Update</th></tr></thead><tbody>';
            if (d.recentActiveCases.length === 0) {
              html += '<tr><td colspan="4" style="text-align:center;color:#999;">No active cases pending action.</td></tr>';
            } else {
              d.recentActiveCases.forEach(function(rc) {
                html += '<tr><td style="font-weight:600;">'+rc.docket+'</td><td>'+rc.date+'</td><td>'+rc.status+'</td><td class="action-text">Review case page</td></tr>';
              });
            }
            html += '</tbody></table>';
            html += '</div></div>';
            html += '</div>'; // End Bottom Row
            
            document.getElementById('dashboardContent').innerHTML = html;
            document.getElementById('dashboardContent').style.display = 'block';
            document.getElementById('dashboardLoading').style.display = 'none';

            // Draw Charts
            drawCharts(d.grantedByCountry, d.pendingByCountry, d.pendingByStatus, countryColors);
          });
        })
        .withFailureHandler(function(e) {
          showError('dashboardContent', 'Failed to load dashboard: ' + e.message);
        })
        .getPortalData('getDashboard', { token: SESSION_TOKEN });
    }

    function drawCharts(grantObj, pendObj, statObj, colors) {
      // Destroy old charts to prevent overlapping issues from reloading
      ['grantedChart', 'pendingChart', 'barChart'].forEach(function(cId) {
        if (chartInstances[cId]) { chartInstances[cId].destroy(); }
      });
      
      Chart.defaults.font.family = "'Inter', sans-serif";
      
      // Donut config generator
      var makeDonut = function(canvasId, dataObj) {
        var labels = Object.keys(dataObj);
        var data = labels.map(function(k){return dataObj[k];});
        var bgColors = labels.map(function(k){return colors[k] || '#bdbdbd';});
        
        var ctx = document.getElementById(canvasId).getContext('2d');
        chartInstances[canvasId] = new Chart(ctx, {
          type: 'doughnut',
          data: { labels: labels, datasets: [{ data: data, backgroundColor: bgColors, borderWidth: 0, cutout: '80%' }] },
          options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c) { return ' ' + c.label + ': ' + c.raw; } } } }
          }
        });
      };
      
      makeDonut('grantedChart', grantObj);
      makeDonut('pendingChart', pendObj);
      
      // Bar Chart
      var statLabels = Object.keys(statObj);
      var statData = statLabels.map(function(k){return statObj[k];});
      var btx = document.getElementById('barChart').getContext('2d');
      chartInstances['barChart'] = new Chart(btx, {
        type: 'bar',
        data: {
          labels: statLabels,
          datasets: [{ data: statData, backgroundColor: '#64b5f6', hoverBackgroundColor: '#2196f3', borderRadius: 4, barThickness: 16 }]
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { display: false, border: {display: false}, grid: {display: false} },
            y: {
              border: {display: false}, grid: {display: false},
              ticks: { color: '#78909c', padding: 12, autoSkip: false, font: {size: 11, weight: '500'} }
            }
          },
          animation: { duration: 1000, easing: 'easeOutQuart' }
        },
        plugins: [{
          id: 'custom_labels',
          afterDraw: function(chart) {
            var ctx = chart.ctx;
            chart.data.datasets.forEach(function(dataset, i) {
              var meta = chart.getDatasetMeta(i);
              meta.data.forEach(function(bar, index) {
                var data = dataset.data[index];
                if (data > 0) {
                  ctx.fillStyle = '#78909c';
                  ctx.font = '11px Inter';
                  ctx.textAlign = 'left';
                  ctx.textBaseline = 'middle';
                  ctx.fillText(data, bar.x + 8, bar.y);
                }
              });
            });
          }
        }]
      });
    }

    // ── Cases ────────────────────────────────
    function loadCases() {
      google.script.run
        .withSuccessHandler(function(cases) {
          handleResponse(cases, 'casesContent', function(cases) {
            if (cases.error) { showError('casesContent', cases.error); return; }
            var html = '<div class="card"><h2>📋 Case Portfolio</h2>';
            if (!cases || cases.length === 0) {
              html += '<p style="color:var(--text-muted);">No cases found.</p>';
            } else {
              var isAdmin = document.getElementById('userInfo').textContent.indexOf('(Admin)') > -1;
              var thead = '<tr><th>Case ID</th><th>Patent Title</th><th>App. No.</th><th>Country</th>' +
                '<th>Filing Date</th><th>Status</th><th>Deadline</th><th>Attorney</th>';
              if (isAdmin) thead += '<th>Actions</th>';
              thead += '</tr>';
              html += '<table><thead>' + thead + '</thead><tbody>';

              cases.forEach(function(c) {
                var safeCaseData = encodeURIComponent(JSON.stringify(c));
                var tr = '<tr><td>'+c.CASE_ID+'</td><td><strong>'+c.PATENT_TITLE+'</strong></td><td>'+(c.APPLICATION_NUMBER||'-')+'</td>' +
                  '<td>'+c.COUNTRY+'</td><td>'+(c.FILING_DATE?formatDate(c.FILING_DATE):'-')+'</td>' +
                  '<td>'+getBadge(c.CURRENT_STATUS)+'</td><td>'+(c.NEXT_DEADLINE?formatDate(c.NEXT_DEADLINE):'-')+'</td>' +
                  '<td>'+(c.ATTORNEY||'-')+'</td>';
                if (isAdmin) {
                  tr += '<td><div style="display:flex;gap:4px;"><button class="btn-primary" style="padding:4px 8px;font-size:11px;" onclick="openCaseModal(\\''+safeCaseData+'\\')">Edit</button>' +
                    '<button class="btn-primary" style="padding:4px 8px;font-size:11px;background:var(--danger);" onclick="deleteItem(\\'Case\\', \\''+c.CASE_ID+'\\')">Del</button></div></td>';
                }
                tr += '</tr>';
                html += tr;
              });
              html += '</tbody></table>';
              if (isAdmin) {
                html += '<div style="margin-top:16px;"><button class="btn-primary" onclick="openCaseModal()">+ Add New Case</button></div>';
              }
            }
            html += '</div>';
            document.getElementById('casesContent').innerHTML = html;
            document.getElementById('casesContent').style.display = 'block';
            document.getElementById('casesLoading').style.display = 'none';
          });
        })
        .withFailureHandler(function(e) {
          showError('casesContent', 'Failed to load cases: ' + e.message);
        })
        .getPortalData('getCases', { token: SESSION_TOKEN });
    }

    // ── Documents ────────────────────────────
    function loadDocuments() {
      google.script.run
        .withSuccessHandler(function(docs) {
          handleResponse(docs, 'documentsContent', function(docs) {
            var html = '';
            if (docs.error) { html = '<div class="alert-msg alert-error">'+docs.error+'</div>'; }
            else if (docs.message) { html = '<div class="alert-msg">'+docs.message+'</div>'; }
            else {
              var icons = {APPLICATIONS:'📝',OFFICE_ACTIONS:'📨',RESPONSES:'📤',CERTIFICATES:'🏆',INVOICES:'🧾',COMMUNICATION:'💬'};
              Object.keys(docs).forEach(function(category) {
                html += '<div class="card doc-section"><h3>'+(icons[category]||'📁')+' '+category.replace(/_/g,' ')+'</h3>';
                if (docs[category].length === 0) {
                  html += '<p style="color:var(--text-muted);font-size:13px;">No documents in this category.</p>';
                } else {
                  docs[category].forEach(function(f) {
                    html += '<div class="doc-item"><span class="material-symbols-outlined" style="color:var(--primary)">draft</span>' +
                      '<div><a href="'+f.url+'" target="_blank">'+f.name+'</a><div class="meta">'+(f.date?formatDate(f.date):'')+
                      ' • '+(f.size > 1048576 ? (f.size/1048576).toFixed(1)+' MB' : (f.size/1024).toFixed(0)+' KB')+'</div></div></div>';
                  });
                }
                html += '</div>';
              });
            }
            document.getElementById('documentsContent').innerHTML = html;
            document.getElementById('documentsContent').style.display = 'block';
            document.getElementById('documentsLoading').style.display = 'none';
          });
        })
        .withFailureHandler(function(e) {
          showError('documentsContent', 'Failed to load documents: ' + e.message);
        })
        .getPortalData('getDocuments', { token: SESSION_TOKEN });
    }

    // ── Invoices ─────────────────────────────
    function loadInvoices() {
      google.script.run
        .withSuccessHandler(function(invoices) {
          handleResponse(invoices, 'invoicesContent', function(invoices) {
            if (invoices.error) { showError('invoicesContent', invoices.error); return; }
            
            var html = '<div class="card" style="margin-bottom: 24px;">';
            html += '<h2>🧾 Finance Dashboard</h2>';

            if (!invoices || invoices.length === 0) {
              html += '<p style="color:var(--text-muted);">No invoices found.</p></div>';
            } else {
              // --- Compute Statistics ---
              var totalCount = invoices.length;
              var paidCount = 0;
              var unpaidCount = 0;
              var paidSum = 0;
              var unpaidSum = 0;

              invoices.forEach(function(inv) {
                var amt = parseFloat(inv.TOTAL) || 0;
                if (inv.PAYMENT_STATUS === 'Paid') {
                  paidCount++;
                  paidSum += amt;
                } else {
                  unpaidCount++;
                  unpaidSum += amt;
                }
              });

              // --- Render Stat Boxes ---
              html += '<div style="display:flex; gap:16px; margin: 24px 0; overflow-x: auto; padding-bottom: 8px;">';
              
              var statBoxStyle = "flex: 1; min-width: 140px; background: var(--bg); padding: 16px; border-radius: 8px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; gap: 8px;";
              var titleStyle = "font-size: 12px; font-weight: 500; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px;";
              var valStyle = "font-size: 20px; font-weight: 600; color: var(--text);";

              html += '<div style="'+statBoxStyle+'"><span style="'+titleStyle+'">Total Invoices</span><span style="'+valStyle+'">'+totalCount+'</span></div>';
              html += '<div style="'+statBoxStyle+'"><span style="'+titleStyle+'">Pending Count</span><span style="'+valStyle+'; color:var(--danger)">'+unpaidCount+'</span></div>';
              html += '<div style="'+statBoxStyle+'"><span style="'+titleStyle+'">Paid Count</span><span style="'+valStyle+'; color:var(--success)">'+paidCount+'</span></div>';
              html += '<div style="'+statBoxStyle+'"><span style="'+titleStyle+'">Pending Amount</span><span style="'+valStyle+'; color:var(--danger)">₹'+unpaidSum.toLocaleString('en-IN', {minimumFractionDigits: 2, maximumFractionDigits: 2})+'</span></div>';
              html += '<div style="'+statBoxStyle+'"><span style="'+titleStyle+'">Paid Amount</span><span style="'+valStyle+'; color:var(--success)">₹'+paidSum.toLocaleString('en-IN', {minimumFractionDigits: 2, maximumFractionDigits: 2})+'</span></div>';
              html += '</div></div>'; // end stats card

              // --- Render Table ---
              html += '<div class="card"><h3>Invoice Ledger</h3>';
              var isAdmin = (document.getElementById('userInfo').textContent.indexOf('(Admin)') > -1);
              html += '<div style="overflow-x: auto;"><table><thead><tr><th>Invoice #</th><th>Date</th><th>Service</th><th>Amount</th><th>GST</th>' +
                '<th>Total</th><th>Status</th><th>PDF</th>' + 
                (isAdmin ? '<th>Action</th>' : '') + '</tr></thead><tbody>';
                
              invoices.forEach(function(inv) {
                var isPaid = inv.PAYMENT_STATUS === 'Paid';
                var safeInvData = encodeURIComponent(JSON.stringify(inv));
                html += '<tr><td style="font-weight:500;">'+inv.INVOICE_ID+'</td><td>'+(inv.INVOICE_DATE?formatDate(inv.INVOICE_DATE):'-')+'</td>' +
                  '<td>'+inv.SERVICE_TYPE+'</td><td>₹'+(parseFloat(inv.AMOUNT)||0).toFixed(2)+'</td>' +
                  '<td>₹'+(parseFloat(inv.GST_AMOUNT)||0).toFixed(2)+'</td>' +
                  '<td><strong>₹'+(parseFloat(inv.TOTAL)||0).toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})+'</strong></td>' +
                  '<td><span class="badge '+(isPaid?'badge-paid':'badge-unpaid')+'">'+ inv.PAYMENT_STATUS+'</span></td>' +
                  '<td>'+(inv.INVOICE_PDF_LINK?'<a href="'+inv.INVOICE_PDF_LINK+'" target="_blank">📄 View</a>':'-')+'</td>';
                
                if (isAdmin) {
                  html += '<td><div style="display:flex;gap:4px;">';
                  if (!isPaid) {
                    html += '<button class="btn-primary" style="padding: 4px 8px; font-size: 11px; background: var(--success);" onclick="markInvoicePaid(\\''+inv.INVOICE_ID+'\\', this)">Paid</button>';
                  }
                  html += '<button class="btn-primary" style="padding:4px 8px;font-size:11px;" onclick="openInvoiceModal(\\''+safeInvData+'\\')">Edit</button>';
                  html += '<button class="btn-primary" style="padding:4px 8px;font-size:11px;background:var(--danger);" onclick="deleteItem(\\'Invoice\\', \\''+inv.INVOICE_ID+'\\')">Del</button>';
                  html += '</div></td>';
                }
                html += '</tr>';
              });
              html += '</tbody></table></div>';
              if (isAdmin) {
                html += '<div style="margin-top:16px;"><button class="btn-primary" onclick="openInvoiceModal()">+ Add New Invoice</button></div>';
              }
              html += '</div>';
            }
            
            document.getElementById('invoicesContent').innerHTML = html;
            document.getElementById('invoicesContent').style.display = 'block';
            document.getElementById('invoicesLoading').style.display = 'none';
          });
        })
        .withFailureHandler(function(e) {
          showError('invoicesContent', 'Failed to load invoices: ' + e.message);
        })
        .getPortalData('getInvoices', { token: SESSION_TOKEN });
    }

    function markInvoicePaid(invoiceId, btnElement) {
      if (!confirm("Are you sure you want to mark " + invoiceId + " as Paid?")) return;
      btnElement.disabled = true;
      btnElement.textContent = "Updating...";
      
      google.script.run.withSuccessHandler(function(r) {
        handleResponse(r, 'invoicesContent', function(res) {
          if (res.success) {
            loadInvoices(); // Reload the table to reflect the new state immediately
            loadDashboard(); // Refresh background dashboard numbers
          } else {
            alert("Error: " + (res.message || res.error || "Failed to update."));
            btnElement.disabled = false;
            btnElement.textContent = "Mark Paid";
          }
        });
      }).withFailureHandler(function(e) {
        alert("Connection Error: " + e.message);
        btnElement.disabled = false;
        btnElement.textContent = "Mark Paid";
      }).getPortalData('markInvoicePaid', { token: SESSION_TOKEN, invoiceId: invoiceId });
    }

    // ── Contact Form ─────────────────────────
    function submitContact() {
      var params = {
        token: SESSION_TOKEN,
        subject: document.getElementById('contactSubject').value,
        caseId: document.getElementById('contactCaseId').value,
        message: document.getElementById('contactMessage').value
      };
      if (!params.subject || !params.message) {
        document.getElementById('contactStatus').innerHTML = '<div class="alert-msg alert-error" style="margin-top:12px">Please fill subject and message.</div>';
        return;
      }
      google.script.run.withSuccessHandler(function(r) {
        handleResponse(r, function(r) {
          document.getElementById('contactStatus').innerHTML = '<div class="alert-msg alert-success" style="margin-top:12px">'+r.message+'</div>';
          document.getElementById('contactSubject').value = '';
          document.getElementById('contactCaseId').value = '';
          document.getElementById('contactMessage').value = '';
        });
      }).withFailureHandler(function(e) {
        document.getElementById('contactStatus').innerHTML = '<div class="alert-msg alert-error" style="margin-top:12px">Error: '+e.message+'</div>';
      }).getPortalData('submitContact', params);
    }

    // ── Change Password ──────────────────────
    function doChangePassword() {
      var oldPwd = document.getElementById('oldPassword').value;
      var newPwd = document.getElementById('newPassword').value;
      var confirmPwd = document.getElementById('confirmPassword').value;
      if (!oldPwd || !newPwd) { showPwdStatus('Please fill all fields.', 'error'); return; }
      if (newPwd !== confirmPwd) { showPwdStatus('New passwords do not match.', 'error'); return; }
      if (newPwd.length < 6) { showPwdStatus('Password must be at least 6 characters.', 'error'); return; }
      
      google.script.run.withSuccessHandler(function(r) {
        handleResponse(r, function(r) {
          showPwdStatus(r.message, r.success ? 'success' : 'error');
          if (r.success) {
            document.getElementById('oldPassword').value = '';
            document.getElementById('newPassword').value = '';
            document.getElementById('confirmPassword').value = '';
          }
        });
      }).getPortalData('changePassword', { token: SESSION_TOKEN, oldPassword: oldPwd, newPassword: newPwd });
    }
    function showPwdStatus(msg, type) {
      document.getElementById('passwordStatus').innerHTML = '<div class="alert-msg '+(type==='error'?'alert-error':'alert-success')+'">'+msg+'</div>';
    }

    // ── Helpers ──────────────────────────────
    function getBadge(status) {
      var cls = 'badge-drafted';
      if (status === 'Filed') cls = 'badge-filed';
      else if (status === 'Granted') cls = 'badge-granted';
      else if (status === 'Under Examination' || status === 'Published') cls = 'badge-examination';
      return '<span class="badge '+cls+'">'+status+'</span>';
    }
    function formatDate(d) {
      try {
        var date = new Date(d);
        return date.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
      } catch(e) { return d; }
    }
    function showError(containerId, msg) {
      document.getElementById(containerId).innerHTML = '<div class="alert-msg alert-error">'+msg+'</div>';
      document.getElementById(containerId).style.display = 'block';
      var loadingId = containerId.replace('Content','Loading');
      var loadingEl = document.getElementById(loadingId);
      if (loadingEl) loadingEl.style.display = 'none';
    }

    // ── Management Tab Handlers (Admin Only) ─────────────────
    
    function loadManageUsers() {
      document.getElementById('managementLoading').style.display = 'block';
      document.getElementById('managementContent').style.display = 'none';
      google.script.run.withSuccessHandler(function(users) {
        handleResponse(users, 'managementContent', function(data) {
          if (data.error) { showError('managementContent', data.error); return; }
          var html = '<h3>Users</h3><table><thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
          data.forEach(function(u) {
            var safeData = encodeURIComponent(JSON.stringify(u));
            html += '<tr><td>'+u.USER_ID+'</td><td>'+u.FULL_NAME+'</td><td>'+u.EMAIL+'</td><td>'+u.ROLE+'</td><td>'+u.STATUS+'</td>';
            html += '<td><div style="display:flex;gap:4px;"><button class="btn-primary" style="padding:4px 8px;font-size:11px;" onclick="openUserModal(\\''+safeData+'\\')">Edit</button>';
            html += '<button class="btn-primary" style="padding:4px 8px;font-size:11px;background:var(--danger);" onclick="deleteItem(\\'User\\', \\''+u.USER_ID+'\\')">Del</button></div></td></tr>';
          });
          html += '</tbody></table><div style="margin-top:16px;"><button class="btn-primary" onclick="openUserModal()">+ Add New User</button></div>';
          document.getElementById('managementContent').innerHTML = html;
          document.getElementById('managementContent').style.display = 'block';
          document.getElementById('managementLoading').style.display = 'none';
        });
      }).getPortalData('getUsers', { token: SESSION_TOKEN });
    }

    function loadManageClients() {
      document.getElementById('managementLoading').style.display = 'block';
      document.getElementById('managementContent').style.display = 'none';
      google.script.run.withSuccessHandler(function(clients) {
        handleResponse(clients, 'managementContent', function(data) {
          if (data.error) { showError('managementContent', data.error); return; }
          var html = '<h3>Clients</h3><table><thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
          data.forEach(function(c) {
            var safeData = encodeURIComponent(JSON.stringify(c));
            html += '<tr><td>'+c.CLIENT_ID+'</td><td>'+c.CLIENT_NAME+'</td><td>'+c.EMAIL+'</td><td>'+c.STATUS+'</td>';
            html += '<td><div style="display:flex;gap:4px;"><button class="btn-primary" style="padding:4px 8px;font-size:11px;" onclick="openClientModal(\\''+safeData+'\\')">Edit</button>';
            html += '<button class="btn-primary" style="padding:4px 8px;font-size:11px;background:var(--danger);" onclick="deleteItem(\\'Client\\', \\''+c.CLIENT_ID+'\\')">Del</button></div></td></tr>';
          });
          html += '</tbody></table><div style="margin-top:16px;"><button class="btn-primary" onclick="openClientModal()">+ Add New Client</button></div>';
          document.getElementById('managementContent').innerHTML = html;
          document.getElementById('managementContent').style.display = 'block';
          document.getElementById('managementLoading').style.display = 'none';
        });
      }).getPortalData('getClients', { token: SESSION_TOKEN });
    }

    function loadManageCases() {
      // Re-use logic for loading cases, just put it in the management section
      document.getElementById('managementLoading').style.display = 'block';
      document.getElementById('managementContent').style.display = 'none';
      google.script.run.withSuccessHandler(function(cases) {
        handleResponse(cases, 'managementContent', function(data) {
          if (data.error) { showError('managementContent', data.error); return; }
          var html = '<h3>Cases</h3><table><thead><tr><th>ID</th><th>Title</th><th>Client ID</th><th>Status</th><th>Actions</th></tr></thead><tbody>';
          data.forEach(function(c) {
            var safeData = encodeURIComponent(JSON.stringify(c));
            html += '<tr><td>'+c.CASE_ID+'</td><td>'+c.PATENT_TITLE+'</td><td>'+c.CLIENT_ID+'</td><td>'+getBadge(c.CURRENT_STATUS)+'</td>';
            html += '<td><div style="display:flex;gap:4px;"><button class="btn-primary" style="padding:4px 8px;font-size:11px;" onclick="openCaseModal(\\''+safeData+'\\')">Edit</button>';
            html += '<button class="btn-primary" style="padding:4px 8px;font-size:11px;background:var(--danger);" onclick="deleteItem(\\'Case\\', \\''+c.CASE_ID+'\\')">Del</button></div></td></tr>';
          });
          html += '</tbody></table><div style="margin-top:16px;"><button class="btn-primary" onclick="openCaseModal()">+ Add New Case</button></div>';
          document.getElementById('managementContent').innerHTML = html;
          document.getElementById('managementContent').style.display = 'block';
          document.getElementById('managementLoading').style.display = 'none';
        });
      }).getPortalData('getCases', { token: SESSION_TOKEN });
    }

    function deleteItem(type, id) {
      if (!confirm("Are you sure you want to delete " + type + " " + id + "?")) return;
      google.script.run.withSuccessHandler(function(r) {
        handleResponse(r, 'managementContent', function(res) {
          if (res.success) {
            alert(res.message);
            if (type === 'User') loadManageUsers();
            if (type === 'Client') loadManageClients();
            if (type === 'Case') { loadManageCases(); loadCases(); } // Update both views
            if (type === 'Invoice') loadInvoices();
          } else { alert("Error: " + (res.message || res.error)); }
        });
      }).getPortalData('delete' + type, { token: SESSION_TOKEN, [type.toLowerCase() + 'Id']: id });
    }

    // ── Common Modal Logic ──
    function closeAllModals() {
      document.querySelectorAll('.modal-overlay').forEach(function(m) { m.style.display = 'none'; });
    }

    function serializeForm(formId) {
      var obj = {};
      var elements = document.getElementById(formId).querySelectorAll('input, select, textarea');
      elements.forEach(function(el) { if (el.id) obj[el.id.replace(formId+'_', '')] = el.value; });
      return obj;
    }

    // ── Modals Opening logic ──
    function openUserModal(safeData) {
      document.getElementById('modal-user').style.display = 'flex';
      var form = document.getElementById('formUser');
      form.reset();
      if (safeData) {
        var u = JSON.parse(decodeURIComponent(safeData));
        document.getElementById('formUser_USER_ID').value = u.USER_ID;
        document.getElementById('formUser_FULL_NAME').value = u.FULL_NAME;
        document.getElementById('formUser_EMAIL').value = u.EMAIL;
        document.getElementById('formUser_ROLE').value = u.ROLE;
        document.getElementById('formUser_CLIENT_ID').value = u.CLIENT_ID || '';
        document.getElementById('formUser_DEPARTMENT').value = u.DEPARTMENT || '';
        document.getElementById('userPasswordHint').textContent = "Leave blank to keep unchanged";
      } else {
        document.getElementById('formUser_USER_ID').value = '';
        document.getElementById('userPasswordHint').textContent = "Required";
      }
    }

    function openClientModal(safeData) {
      document.getElementById('modal-client').style.display = 'flex';
      var form = document.getElementById('formClient');
      form.reset();
      if (safeData) {
        var c = JSON.parse(decodeURIComponent(safeData));
        document.getElementById('formClient_CLIENT_ID').value = c.CLIENT_ID;
        document.getElementById('formClient_CLIENT_NAME').value = c.CLIENT_NAME;
        document.getElementById('formClient_CONTACT_PERSON').value = c.CONTACT_PERSON;
        document.getElementById('formClient_EMAIL').value = c.EMAIL;
        document.getElementById('formClient_PHONE').value = c.PHONE || '';
        document.getElementById('formClient_ADDRESS').value = c.ADDRESS || '';
        document.getElementById('formClient_NOTES').value = c.NOTES || '';
      } else {
        document.getElementById('formClient_CLIENT_ID').value = '';
      }
    }

    function openCaseModal(safeData) {
      document.getElementById('modal-case').style.display = 'flex';
      var form = document.getElementById('formCase');
      form.reset();
      if (safeData) {
        var c = JSON.parse(decodeURIComponent(safeData));
        document.getElementById('formCase_CASE_ID').value = c.CASE_ID;
        document.getElementById('formCase_CLIENT_ID').value = c.CLIENT_ID;
        document.getElementById('formCase_PATENT_TITLE').value = c.PATENT_TITLE;
        document.getElementById('formCase_APPLICATION_NUMBER').value = c.APPLICATION_NUMBER || '';
        document.getElementById('formCase_COUNTRY').value = c.COUNTRY || 'India';
        if (c.FILING_DATE) {
          try { document.getElementById('formCase_FILING_DATE').value = new Date(c.FILING_DATE).toISOString().split('T')[0]; } catch(e){}
        }
        if (c.NEXT_DEADLINE) {
          try { document.getElementById('formCase_NEXT_DEADLINE').value = new Date(c.NEXT_DEADLINE).toISOString().split('T')[0]; } catch(e){}
        }
        document.getElementById('formCase_CURRENT_STATUS').value = c.CURRENT_STATUS;
        document.getElementById('formCase_PATENT_TYPE').value = c.PATENT_TYPE;
        document.getElementById('formCase_ATTORNEY').value = c.ATTORNEY || '';
        document.getElementById('formCase_PRIORITY').value = c.PRIORITY || 'Normal';
        document.getElementById('formCase_NOTES').value = c.NOTES || '';
      } else {
        document.getElementById('formCase_CASE_ID').value = '';
      }
    }

    function openInvoiceModal(safeData) {
      document.getElementById('modal-invoice').style.display = 'flex';
      var form = document.getElementById('formInvoice');
      form.reset();
      if (safeData) {
        var inv = JSON.parse(decodeURIComponent(safeData));
        document.getElementById('formInvoice_INVOICE_ID').value = inv.INVOICE_ID;
        document.getElementById('formInvoice_CLIENT_ID').value = inv.CLIENT_ID;
        document.getElementById('formInvoice_CASE_ID').value = inv.CASE_ID || '';
        document.getElementById('formInvoice_SERVICE_TYPE').value = inv.SERVICE_TYPE;
        document.getElementById('formInvoice_DESCRIPTION').value = inv.DESCRIPTION || '';
        document.getElementById('formInvoice_AMOUNT').value = inv.AMOUNT || 0;
        document.getElementById('formInvoice_GST_RATE').value = inv.GST_RATE || 18;
        document.getElementById('formInvoice_NOTES').value = inv.NOTES || '';
      } else {
        document.getElementById('formInvoice_INVOICE_ID').value = '';
      }
    }

    function saveItem(type, formId) {
      document.getElementById('btnSave'+type).disabled = true;
      document.getElementById('btnSave'+type).textContent = "Saving...";
      var dataObj = serializeForm(formId);
      var params = { token: SESSION_TOKEN };
      params[type.toLowerCase() + 'Data'] = dataObj;

      google.script.run.withSuccessHandler(function(r) {
        document.getElementById('btnSave'+type).disabled = false;
        document.getElementById('btnSave'+type).textContent = "Save";
        handleResponse(r, 'managementContent', function(res) {
          if (res.success) {
            alert(res.message);
            closeAllModals();
            if (type === 'User') loadManageUsers();
            if (type === 'Client') loadManageClients();
            if (type === 'Case') { loadManageCases(); loadCases(); loadDashboard(); }
            if (type === 'Invoice') { loadInvoices(); loadDashboard(); }
          } else { alert("Error: " + (res.message || res.error)); }
        });
      }).getPortalData('save' + type, params);
    }
  </script>

  <!-- HTML Modals overlay and styling -->
  <style>
    .modal-overlay {
      display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: center;
    }
    .modal-content {
      background: white; padding: 24px; border-radius: 8px; width: 100%; max-width: 500px;
      max-height: 90vh; overflow-y: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.2);
    }
    .modal-content h3 { margin-bottom: 16px; color: var(--primary-dark); }
  </style>

  <div class="modal-overlay" id="modal-user">
    <div class="modal-content">
      <h3>User Management</h3>
      <form id="formUser" onsubmit="event.preventDefault(); saveItem('User', 'formUser');">
        <input type="hidden" id="formUser_USER_ID" />
        <div class="form-group"><label>Full Name *</label><input id="formUser_FULL_NAME" required /></div>
        <div class="form-group"><label>Email *</label><input id="formUser_EMAIL" type="email" required /></div>
        <div class="form-group"><label>Password <span id="userPasswordHint" style="font-size:11px;color:#888;"></span></label><input id="formUser_PASSWORD_HASH" type="password" /></div>
        <div class="form-group"><label>Role *</label><select id="formUser_ROLE"><option>Admin</option><option>Attorney</option><option>Client</option></select></div>
        <div class="form-group"><label>Client ID (If role is Client)</label><input id="formUser_CLIENT_ID" /></div>
        <div class="form-group"><label>Department</label><input id="formUser_DEPARTMENT" /></div>
        <div style="display:flex;gap:8px;margin-top:24px;">
          <button type="submit" class="btn-primary" id="btnSaveUser">Save</button>
          <button type="button" class="btn-primary" style="background:var(--text-muted);" onclick="closeAllModals()">Cancel</button>
        </div>
      </form>
    </div>
  </div>

  <div class="modal-overlay" id="modal-client">
    <div class="modal-content">
      <h3>Client Management</h3>
      <form id="formClient" onsubmit="event.preventDefault(); saveItem('Client', 'formClient');">
        <input type="hidden" id="formClient_CLIENT_ID" />
        <div class="form-group"><label>Client Name *</label><input id="formClient_CLIENT_NAME" required /></div>
        <div class="form-group"><label>Contact Person *</label><input id="formClient_CONTACT_PERSON" required /></div>
        <div class="form-group"><label>Email *</label><input id="formClient_EMAIL" type="email" required /></div>
        <div class="form-group"><label>Phone</label><input id="formClient_PHONE" /></div>
        <div class="form-group"><label>Address</label><textarea id="formClient_ADDRESS"></textarea></div>
        <div class="form-group"><label>Notes</label><textarea id="formClient_NOTES"></textarea></div>
        <div style="display:flex;gap:8px;margin-top:24px;">
          <button type="submit" class="btn-primary" id="btnSaveClient">Save</button>
          <button type="button" class="btn-primary" style="background:var(--text-muted);" onclick="closeAllModals()">Cancel</button>
        </div>
      </form>
    </div>
  </div>

  <div class="modal-overlay" id="modal-case">
    <div class="modal-content">
      <h3>Case Management</h3>
      <form id="formCase" onsubmit="event.preventDefault(); saveItem('Case', 'formCase');">
        <input type="hidden" id="formCase_CASE_ID" />
        <div class="form-group"><label>Client ID *</label><input id="formCase_CLIENT_ID" required /></div>
        <div class="form-group"><label>Patent Title *</label><input id="formCase_PATENT_TITLE" required /></div>
        <div style="display:flex;gap:12px;">
          <div class="form-group" style="flex:1;"><label>App. No.</label><input id="formCase_APPLICATION_NUMBER" /></div>
          <div class="form-group" style="flex:1;"><label>Country</label><select id="formCase_COUNTRY"><option>India</option><option>USA</option><option>EPO</option><option>PCT</option><option>Other</option></select></div>
        </div>
        <div style="display:flex;gap:12px;">
          <div class="form-group" style="flex:1;"><label>Filing Date</label><input id="formCase_FILING_DATE" type="date" /></div>
          <div class="form-group" style="flex:1;"><label>Next Deadline</label><input id="formCase_NEXT_DEADLINE" type="date" /></div>
        </div>
        <div style="display:flex;gap:12px;">
          <div class="form-group" style="flex:1;"><label>Status</label><select id="formCase_CURRENT_STATUS"><option>Drafted</option><option>Filed</option><option>Published</option><option>Under Examination</option><option>Granted</option><option>Abandoned</option><option>Lapsed</option></select></div>
          <div class="form-group" style="flex:1;"><label>Type</label><select id="formCase_PATENT_TYPE"><option>Utility</option><option>Design</option><option>PCT</option><option>Provisional</option></select></div>
        </div>
        <div class="form-group"><label>Attorney Email</label><input id="formCase_ATTORNEY" type="email" /></div>
        <div class="form-group"><label>Priority</label><select id="formCase_PRIORITY"><option>Normal</option><option>High</option><option>Urgent</option></select></div>
        <div class="form-group"><label>Notes</label><textarea id="formCase_NOTES"></textarea></div>
        <div style="display:flex;gap:8px;margin-top:24px;">
          <button type="submit" class="btn-primary" id="btnSaveCase">Save</button>
          <button type="button" class="btn-primary" style="background:var(--text-muted);" onclick="closeAllModals()">Cancel</button>
        </div>
      </form>
    </div>
  </div>

  <div class="modal-overlay" id="modal-invoice">
    <div class="modal-content">
      <h3>Invoice Management</h3>
      <form id="formInvoice" onsubmit="event.preventDefault(); saveItem('Invoice', 'formInvoice');">
        <input type="hidden" id="formInvoice_INVOICE_ID" />
        <div class="form-group"><label>Client ID *</label><input id="formInvoice_CLIENT_ID" required /></div>
        <div class="form-group"><label>Case ID (Optional)</label><input id="formInvoice_CASE_ID" /></div>
        <div class="form-group"><label>Service Type</label><select id="formInvoice_SERVICE_TYPE"><option>Patent Filing</option><option>Patent Prosecution</option><option>Patent Search</option><option>Legal Opinion</option><option>Drafting Fees</option><option>Government Fees</option><option>Consultation</option><option>Annual Maintenance</option><option>Other</option></select></div>
        <div class="form-group"><label>Description</label><textarea id="formInvoice_DESCRIPTION"></textarea></div>
        <div style="display:flex;gap:12px;">
          <div class="form-group" style="flex:1;"><label>Amount (₹) *</label><input id="formInvoice_AMOUNT" type="number" step="0.01" required /></div>
          <div class="form-group" style="flex:1;"><label>GST Rate (%)</label><input id="formInvoice_GST_RATE" type="number" value="18" /></div>
        </div>
        <div class="form-group"><label>Notes</label><textarea id="formInvoice_NOTES"></textarea></div>
        <div style="display:flex;gap:8px;margin-top:24px;">
          <button type="submit" class="btn-primary" id="btnSaveInvoice">Save</button>
          <button type="button" class="btn-primary" style="background:var(--text-muted);" onclick="closeAllModals()">Cancel</button>
        </div>
      </form>
    </div>
  </div>

</body>
</html>`;
}
function clerkLogin_(email) {
  if (!email) return { success: false, message: 'Email required.' };

  var user = getUserRecordByEmail_(email);
  if (user && String(user.STATUS || "") === "Active") {
    var token = Utilities.getUuid();
    var cache = CacheService.getScriptCache();
    cache.put('session_' + token, JSON.stringify(buildSessionFromUser_(user)), 21600);

    logActivity_('CLERK_LOGIN', 'USER', user.USER_ID, 'Login via Clerk/GitHub Pages');

    return {
      success: true,
      token: token,
      role: normalizeRole_(user.ROLE),
      name: user.FULL_NAME || user.EMAIL
    };
  }

  return { success: false, message: 'Email not found or account inactive.' };
}
