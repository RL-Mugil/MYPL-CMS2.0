/**
 * ============================================================
 * 07_Addons.gs - Role Access, Organizations, Messages & Ops
 * ============================================================
 */

var ROLE_ORDER_ = {
  "Super Admin": 60,
  "Admin": 50,
  "Galvanizer": 45,
  "Staff": 40,
  "Attorney": 30,
  "Client Admin": 20,
  "Client Employee": 15,
  "Individual Client": 10,
  "Client": 10
};

var _REQUEST_RECORD_CACHE_ = {};
var _REQUEST_HEADERS_CACHE_ = {};

function clearRequestSheetCache_(configKey) {
  delete _REQUEST_RECORD_CACHE_[configKey];
  delete _REQUEST_HEADERS_CACHE_[configKey];
}

function getRecords_(configKey) {
  if (_REQUEST_RECORD_CACHE_.hasOwnProperty(configKey)) {
    return _REQUEST_RECORD_CACHE_[configKey];
  }
  var sheet = getSheet_(configKey);
  var data = sheet.getDataRange().getValues();
  if (!data.length) return [];
  var headers = data[0];
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    records.push(row);
  }
  _REQUEST_RECORD_CACHE_[configKey] = records;
  _REQUEST_HEADERS_CACHE_[configKey] = headers;
  return records;
}

function getSheetHeaders_(configKey) {
  if (_REQUEST_HEADERS_CACHE_.hasOwnProperty(configKey)) {
    return _REQUEST_HEADERS_CACHE_[configKey];
  }
  var sheet = getSheet_(configKey);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  _REQUEST_HEADERS_CACHE_[configKey] = headers;
  return headers;
}

function appendRecord_(configKey, record) {
  var sheet = getSheet_(configKey);
  var headers = getSheetHeaders_(configKey);
  var row = headers.map(function(header) {
    return record.hasOwnProperty(header) ? record[header] : "";
  });
  sheet.appendRow(row);
  clearRequestSheetCache_(configKey);
}

function updateRecordById_(configKey, idField, idValue, updates) {
  var sheet = getSheet_(configKey);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf(idField)] === idValue) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) return false;

  Object.keys(updates).forEach(function(key) {
    var colIndex = headers.indexOf(key);
    if (colIndex > -1) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(updates[key]);
    }
  });
  clearRequestSheetCache_(configKey);
  return true;
}

function getUserRecordByEmail_(email) {
  var users = getRecords_("USERS");
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].EMAIL || "").toLowerCase() === String(email || "").toLowerCase()) {
      return users[i];
    }
  }
  return null;
}

function getUserRecordById_(userId) {
  var users = getRecords_("USERS");
  for (var i = 0; i < users.length; i++) {
    if (users[i].USER_ID === userId) return users[i];
  }
  return null;
}

function normalizeRole_(role) {
  if (role === "Client Manager") return "Client Admin";
  if (role === "Client") return "Individual Client";
  return role || "Individual Client";
}

function getEffectiveRoles_(userOrSession) {
  if (!userOrSession) return [];
  var roles = [];
  if (userOrSession.role || userOrSession.ROLE) {
    roles.push(normalizeRole_(userOrSession.role || userOrSession.ROLE));
  }
  var additional = userOrSession.additionalRoles || userOrSession.ADDITIONAL_ROLES || "";
  String(additional).split(",").forEach(function(role) {
    var trimmed = normalizeRole_(String(role || "").trim());
    if (trimmed && roles.indexOf(trimmed) === -1) roles.push(trimmed);
  });
  return roles.filter(function(role) { return !!role; });
}

function roleLevel_(role) {
  return ROLE_ORDER_[normalizeRole_(role)] || 0;
}

function hasRoleAtLeast_(role, minimumRole) {
  if (Object.prototype.toString.call(role) === "[object Array]") {
    return role.some(function(item) { return roleLevel_(item) >= roleLevel_(minimumRole); });
  }
  return roleLevel_(role) >= roleLevel_(minimumRole);
}

function hasAnyRole_(userOrSession, roles) {
  var effectiveRoles = getEffectiveRoles_(userOrSession);
  return roles.some(function(role) {
    return effectiveRoles.indexOf(normalizeRole_(role)) > -1;
  });
}

function isInternalRole_(role) {
  var roles = Object.prototype.toString.call(role) === "[object Array]" ? role : [role];
  return roles.some(function(item) {
    return ["Super Admin", "Admin", "Galvanizer", "Staff", "Attorney"].indexOf(normalizeRole_(item)) > -1;
  });
}

function isClientSideSession_(session) {
  return !!session && !isInternalRole_(getEffectiveRoles_(session));
}

function canAccessManagement_(session) {
  if (!session) return false;
  return hasAnyRole_(session, ["Super Admin", "Admin", "Client Admin"]);
}

function canViewFinance_(session) {
  if (!session) return false;
  if (hasRoleAtLeast_(session.role, "Admin")) return true;
  return String(session.canViewFinance || "").toLowerCase() === "yes";
}

function canManageAllData_(session) {
  return !!session && hasRoleAtLeast_(getEffectiveRoles_(session), "Admin");
}

function canApproveExpenseClaims_(session) {
  return !!session && normalizeRole_(session.role) === "Super Admin";
}

function getOrganizationById_(orgId) {
  var orgs = getRecords_("ORGANIZATIONS");
  for (var i = 0; i < orgs.length; i++) {
    if (orgs[i].ORG_ID === orgId) return orgs[i];
  }
  return null;
}

function getClientsByOrgId_(orgId) {
  return getAllClients().filter(function(client) {
    return String(client.ORG_ID || "").trim().toUpperCase() === String(orgId || "").trim().toUpperCase();
  });
}

function getResolvedOrgIdForSession_(session) {
  var directOrgId = String(session && session.orgId || "").trim().toUpperCase();
  if (directOrgId) return directOrgId;

  var sessionClientId = String(session && session.clientId || "").trim().toUpperCase();
  var sessionEmail = String(session && session.email || "").trim().toLowerCase();
  var clients = getAllClients();

  if (sessionClientId) {
    for (var i = 0; i < clients.length; i++) {
      if (String(clients[i].CLIENT_ID || "").trim().toUpperCase() === sessionClientId) {
        var clientOrgId = String(clients[i].ORG_ID || "").trim().toUpperCase();
        if (clientOrgId) return clientOrgId;
      }
    }
  }

  if (sessionEmail) {
    for (var j = 0; j < clients.length; j++) {
      if (String(clients[j].EMAIL || "").trim().toLowerCase() === sessionEmail) {
        var emailOrgId = String(clients[j].ORG_ID || "").trim().toUpperCase();
        if (emailOrgId) return emailOrgId;
      }
    }
  }

  return "";
}

function buildSessionFromUser_(user) {
  return {
    email: user.EMAIL,
    role: normalizeRole_(user.ROLE),
    additionalRoles: user.ADDITIONAL_ROLES || "",
    clientId: user.CLIENT_ID || "",
    orgId: user.ORG_ID || "",
    name: user.FULL_NAME || user.EMAIL,
    userId: user.USER_ID,
    canViewFinance: user.CAN_VIEW_FINANCE || "",
    reportsTo: user.REPORTS_TO || ""
  };
}

function normalizeClientCode_(code, region) {
  var cleaned = String(code || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!cleaned) return "";
  if (/[MY]$/.test(cleaned)) return cleaned;
  var suffix = String(region || "").toUpperCase() === "ABROAD" ? "Y" : "M";
  return cleaned + suffix;
}

function looksLikeModernClientCode_(value) {
  return /^[A-Z0-9]{3,4}[MY]$/.test(String(value || "").toUpperCase());
}

function nextCaseSequenceForClient_(clientCode) {
  var cases = getAllCases();
  var prefix = String(clientCode || "").toUpperCase();
  var maxNum = 0;
  cases.forEach(function(caseItem) {
    var caseId = String(caseItem.CASE_ID || "").toUpperCase();
    if (caseId.indexOf(prefix) === 0) {
      var suffix = caseId.substring(prefix.length);
      var parsed = parseInt(suffix, 10);
      if (!isNaN(parsed) && parsed > maxNum) maxNum = parsed;
    }
  });
  return ("000" + (maxNum + 1)).slice(-3);
}

function generateCaseIdForClient_(client) {
  var clientCode = normalizeClientCode_(client.CLIENT_CODE || client.CLIENT_ID, client.CLIENT_REGION);
  if (!looksLikeModernClientCode_(clientCode)) {
    throw new Error("Valid client code is required before creating a case.");
  }
  return clientCode + nextCaseSequenceForClient_(clientCode);
}

function getAccessibleClientsForUser_(session) {
  var role = normalizeRole_(session.role);
  var roles = getEffectiveRoles_(session);
  var clients = getAllClients();
  var sessionOrgId = getResolvedOrgIdForSession_(session);
  var sessionClientId = String(session.clientId || "").trim().toUpperCase();
  var sessionEmail = String(session.email || "").trim().toLowerCase();

  if (hasRoleAtLeast_(roles, "Admin")) return clients;

  if (roles.indexOf("Galvanizer") > -1 || role === "Staff") {
    return clients.filter(function(client) {
      return String(client.ASSIGNED_STAFF_EMAIL || "").trim().toLowerCase() === sessionEmail;
    });
  }

  if (role === "Attorney") {
    var attorneyCases = getAttorneyCases(session.email);
    var caseClientIds = {};
    attorneyCases.forEach(function(item) { caseClientIds[item.CLIENT_ID] = true; });
    return clients.filter(function(client) { return !!caseClientIds[client.CLIENT_ID]; });
  }

  if (role === "Client Admin" || role === "Client Employee") {
    if (sessionOrgId) {
      return clients.filter(function(client) {
        return String(client.ORG_ID || "").trim().toUpperCase() === sessionOrgId;
      });
    }
    if (sessionClientId) {
      return clients.filter(function(client) {
        return String(client.CLIENT_ID || "").trim().toUpperCase() === sessionClientId;
      });
    }
  }

  if (role === "Individual Client" && sessionClientId) {
    return clients.filter(function(client) {
      return String(client.CLIENT_ID || "").trim().toUpperCase() === sessionClientId;
    });
  }

  return [];
}

function getAccessibleClientIdsForUser_(session) {
  return getAccessibleClientsForUser_(session).map(function(client) {
    return client.CLIENT_ID;
  });
}

function getAccessibleCasesForUser_(session) {
  var role = normalizeRole_(session.role);
  var roles = getEffectiveRoles_(session);
  var sessionEmail = String(session.email || "").trim().toLowerCase();
  var sessionOrgId = getResolvedOrgIdForSession_(session);

  if (hasRoleAtLeast_(roles, "Admin")) return getAllCases();
  if (roles.indexOf("Galvanizer") > -1) {
    return getAllCases().filter(function(caseItem) {
      return String(caseItem.ASSIGNED_STAFF_EMAIL || "").trim().toLowerCase() === sessionEmail ||
        String(caseItem.GALVANIZER_EMAIL || "").trim().toLowerCase() === sessionEmail;
    });
  }
  if (role === "Attorney") return getAttorneyCases(session.email);

  var clientIds = getAccessibleClientIdsForUser_(session);
  var clientIdMap = {};
  clientIds.forEach(function(id) { clientIdMap[id] = true; });

  return getAllCases().filter(function(caseItem) {
    if (role === "Staff") {
      return String(caseItem.ASSIGNED_STAFF_EMAIL || "").trim().toLowerCase() === sessionEmail ||
        clientIdMap[caseItem.CLIENT_ID];
    }
    if (sessionOrgId) {
      return String(caseItem.ORG_ID || "").trim().toUpperCase() === sessionOrgId || clientIdMap[caseItem.CLIENT_ID];
    }
    return clientIdMap[caseItem.CLIENT_ID];
  });
}

function normalizeComparableDate_(value) {
  if (!value) return "";
  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) return "";
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function filterCases_(cases, filters) {
  filters = filters || {};
  var clientCode = String(filters.clientCode || "").trim().toUpperCase();
  var caseId = String(filters.caseId || "").trim().toUpperCase();
  var status = String(filters.status || "").trim();
  var country = String(filters.country || "").trim();
  var assignedStaff = String(filters.assignedStaff || "").trim().toLowerCase();
  var attorney = String(filters.attorney || "").trim().toLowerCase();
  var galvanizer = String(filters.galvanizer || "").trim().toLowerCase();
  var workflowStage = String(filters.workflowStage || "").trim();
  var fromDate = normalizeComparableDate_(filters.fromDate || "");
  var toDate = normalizeComparableDate_(filters.toDate || "");
  var query = String(filters.query || "").trim().toLowerCase();

  return (cases || []).filter(function(caseItem) {
    if (clientCode && String(caseItem.CLIENT_ID || "").toUpperCase() !== clientCode) return false;
    if (caseId && String(caseItem.CASE_ID || "").toUpperCase().indexOf(caseId) === -1) return false;
    if (status && String(caseItem.CURRENT_STATUS || "") !== status) return false;
    if (country && String(caseItem.COUNTRY || "") !== country) return false;
    if (assignedStaff && String(caseItem.ASSIGNED_STAFF_EMAIL || "").toLowerCase() !== assignedStaff) return false;
    if (attorney && String(caseItem.ATTORNEY || "").toLowerCase() !== attorney) return false;
    if (galvanizer && String(caseItem.GALVANIZER_EMAIL || "").toLowerCase() !== galvanizer) return false;
    if (workflowStage && String(caseItem.WORKFLOW_STAGE || "") !== workflowStage) return false;
    if (fromDate || toDate) {
      var comparable = normalizeComparableDate_(caseItem.FILING_DATE || caseItem.CREATED_AT || "");
      if (!comparable) return false;
      if (fromDate && comparable < fromDate) return false;
      if (toDate && comparable > toDate) return false;
    }
    if (query) {
      var hay = [
        caseItem.CASE_ID,
        caseItem.CLIENT_ID,
        caseItem.PATENT_TITLE,
        caseItem.APPLICATION_NUMBER,
        caseItem.ATTORNEY,
        caseItem.ASSIGNED_STAFF_EMAIL,
        caseItem.GALVANIZER_EMAIL
      ].join(" ").toLowerCase();
      if (hay.indexOf(query) === -1) return false;
    }
    return true;
  });
}

function filterAccessibleCases_(session, filters) {
  return filterCases_(getAccessibleCasesForUser_(session), filters || {});
}

function filterInvoicesByCases_(invoices, cases) {
  var allowedCaseIds = {};
  var allowedClientIds = {};
  (cases || []).forEach(function(caseItem) {
    allowedCaseIds[caseItem.CASE_ID] = true;
    allowedClientIds[caseItem.CLIENT_ID] = true;
  });
  return (invoices || []).filter(function(invoice) {
    if (invoice.CASE_ID && allowedCaseIds[invoice.CASE_ID]) return true;
    return allowedClientIds[invoice.CLIENT_ID];
  });
}

function filterUsers_(users, filters) {
  filters = filters || {};
  var query = String(filters.query || "").trim().toLowerCase();
  var role = String(filters.role || "").trim();
  if (!query && !role) return users || [];
  return (users || []).filter(function(user) {
    if (role && String(user.ROLE || "") !== role) return false;
    if (!query) return true;
    var hay = [
      user.USER_ID,
      user.FULL_NAME,
      user.EMAIL,
      user.ROLE,
      user.CLIENT_ID,
      user.ORG_ID
    ].join(" ").toLowerCase();
    return hay.indexOf(query) > -1;
  });
}

function getGalvanizerQueue_(session, filters) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("galvanizerQueue", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var cases = getAccessibleCasesForUser_(session);
  cases = filterCases_(cases, filters);
  cases = cases.filter(function(caseItem) {
    return ["Drafting", "Ready for Attorney", "Under Attorney Review"].indexOf(String(caseItem.WORKFLOW_STAGE || "Drafting")) > -1;
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(cases), 60);
}

function getAccessibleInvoicesForUser_(session) {
  if (!canViewFinance_(session)) return [];
  var invoices = getRecords_("INVOICE");
  if (canManageAllData_(session)) return invoices;

  var clientIds = getAccessibleClientIdsForUser_(session);
  var clientIdMap = {};
  clientIds.forEach(function(id) { clientIdMap[id] = true; });

  return invoices.filter(function(invoice) {
    if (session.orgId && invoice.ORG_ID === session.orgId) return true;
    return clientIdMap[invoice.CLIENT_ID];
  });
}

function getCircles_(session) {
  if (!isInternalRole_(getEffectiveRoles_(session))) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("circles", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var circles = getRecords_("CIRCLES").filter(function(circle) {
    return String(circle.STATUS || "Active") === "Active";
  });
  var memberships = getRecords_("CIRCLE_MEMBERS");
  var userEmail = String(session.email || "").toLowerCase();
  if (!canManageAllData_(session)) {
    var allowedCircleIds = memberships.filter(function(item) {
      return String(item.USER_EMAIL || "").toLowerCase() === userEmail && item.STATUS !== "Removed";
    }).map(function(item) { return item.CIRCLE_ID; });
    circles = circles.filter(function(circle) {
      return allowedCircleIds.indexOf(circle.CIRCLE_ID) > -1;
    });
  }
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(circles), 300);
}

function getCircleMembers_(session, circleId) {
  if (!isInternalRole_(getEffectiveRoles_(session))) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("circleMembers", session, circleId || "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var members = getRecords_("CIRCLE_MEMBERS").filter(function(item) {
    return item.CIRCLE_ID === circleId && item.STATUS !== "Removed";
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(members), 300);
}

function saveCircle_(session, circleData) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  return withScriptLock_(function() {
    var sheet = getSheet_("CIRCLES");
    if (circleData.CIRCLE_ID) {
      var ok = updateRecordById_("CIRCLES", "CIRCLE_ID", circleData.CIRCLE_ID, {
        CIRCLE_NAME: circleData.CIRCLE_NAME || "",
        DESCRIPTION: circleData.DESCRIPTION || "",
        STATUS: circleData.STATUS || "Active",
        UPDATED_AT: new Date()
      });
      if (!ok) return { error: "Circle not found." };
      clearPortalCaches_(["circles", "circleMembers", "users"]);
      return { success: true, circleId: circleData.CIRCLE_ID, message: "Circle updated." };
    }
    var circleId = generateId_("CRC", sheet, 0);
    appendRecord_("CIRCLES", {
      CIRCLE_ID: circleId,
      CIRCLE_NAME: circleData.CIRCLE_NAME || "",
      DESCRIPTION: circleData.DESCRIPTION || "",
      STATUS: circleData.STATUS || "Active",
      CREATED_BY: session.email,
      CREATED_AT: new Date(),
      UPDATED_AT: new Date()
    });
    clearPortalCaches_(["circles", "circleMembers", "users"]);
    return { success: true, circleId: circleId, message: "Circle created." };
  });
}

function deleteCircle_(session, circleId) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  var ok = updateRecordById_("CIRCLES", "CIRCLE_ID", circleId, {
    STATUS: "Inactive",
    UPDATED_AT: new Date()
  });
  if (ok) clearPortalCaches_(["circles", "circleMembers", "users"]);
  return ok ? { success: true, message: "Circle deactivated." } : { error: "Circle not found." };
}

function saveCircleMember_(session, memberData) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  return withScriptLock_(function() {
    var user = memberData.USER_ID ? getUserRecordById_(memberData.USER_ID) : getUserRecordByEmail_(memberData.USER_EMAIL);
    if (!user) return { error: "User not found." };
    var members = getRecords_("CIRCLE_MEMBERS");
    for (var i = 0; i < members.length; i++) {
      if (members[i].CIRCLE_ID === memberData.CIRCLE_ID && members[i].USER_ID === user.USER_ID && members[i].STATUS !== "Removed") {
        updateRecordById_("CIRCLE_MEMBERS", "MEMBERSHIP_ID", members[i].MEMBERSHIP_ID, {
          ROLE_IN_CIRCLE: memberData.ROLE_IN_CIRCLE || members[i].ROLE_IN_CIRCLE || "Member",
          STATUS: "Active",
          UPDATED_AT: new Date()
        });
        clearPortalCaches_(["circles", "circleMembers", "users"]);
        return { success: true, membershipId: members[i].MEMBERSHIP_ID, message: "Circle member updated." };
      }
    }
    var sheet = getSheet_("CIRCLE_MEMBERS");
    var membershipId = generateId_("CRM", sheet, 0);
    appendRecord_("CIRCLE_MEMBERS", {
      MEMBERSHIP_ID: membershipId,
      CIRCLE_ID: memberData.CIRCLE_ID,
      USER_ID: user.USER_ID,
      USER_EMAIL: user.EMAIL,
      ROLE_IN_CIRCLE: memberData.ROLE_IN_CIRCLE || "Member",
      STATUS: "Active",
      ADDED_BY: session.email,
      CREATED_AT: new Date(),
      UPDATED_AT: new Date()
    });
    clearPortalCaches_(["circles", "circleMembers", "users"]);
    return { success: true, membershipId: membershipId, message: "Circle member added." };
  });
}

function removeCircleMember_(session, membershipId) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  var ok = updateRecordById_("CIRCLE_MEMBERS", "MEMBERSHIP_ID", membershipId, {
    STATUS: "Removed",
    UPDATED_AT: new Date()
  });
  if (ok) clearPortalCaches_(["circles", "circleMembers", "users"]);
  return ok ? { success: true, message: "Circle member removed." } : { error: "Membership not found." };
}

function bulkUpdateCases_(session, payload) {
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Admin")) return { error: "Access denied." };
  var caseIds = payload.caseIds || [];
  if (!caseIds.length) return { error: "Select at least one case." };

  var updates = {};
  [
    "ASSIGNED_STAFF_EMAIL",
    "GALVANIZER_EMAIL",
    "ATTORNEY",
    "WORKFLOW_STAGE",
    "CURRENT_STATUS",
    "PRIORITY",
    "NEXT_DEADLINE"
  ].forEach(function(key) {
    if (payload.updates && payload.updates[key] !== undefined && payload.updates[key] !== "") {
      updates[key] = key === "NEXT_DEADLINE" ? new Date(payload.updates[key]) : payload.updates[key];
    }
  });

  if (!Object.keys(updates).length) return { error: "No update values supplied." };

  var updated = 0;
  caseIds.forEach(function(caseId) {
    var ok = updateCase(caseId, updates);
    if (ok && ok.success) updated++;
  });

  logActivity_("BULK_UPDATE_CASES", "CASE", caseIds.join(","), JSON.stringify(updates));
  if (updated) clearPortalCaches_(["cases", "casesPage", "dashboard", "dashboardSummary", "dashboardDetails", "galvanizerQueue", "documents"]);
  return { success: true, updated: updated, message: updated + " cases updated." };
}

function normalizeImportRecordType_(value) {
  var text = String(value || "").trim().toLowerCase();
  if (text === "patent") return "Patent";
  if (text === "trademark") return "Trademark";
  return "";
}

function findClientByCodeOrId_(clientCode) {
  var normalized = String(clientCode || "").trim().toUpperCase();
  if (!normalized) return null;
  var clients = getAllClients();
  for (var i = 0; i < clients.length; i++) {
    var code = String(clients[i].CLIENT_CODE || clients[i].CLIENT_ID || "").trim().toUpperCase();
    var id = String(clients[i].CLIENT_ID || "").trim().toUpperCase();
    if (normalized === code || normalized === id) return clients[i];
  }
  return null;
}

function parseImportDateValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }
  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) return "";
  return parsed;
}

function buildImportedCaseNote_(row, selectedClientCode) {
  var parts = [
    "Imported from DocketTrak",
    "Selected Client: " + String(selectedClientCode || ""),
    "Source Record Type: " + String(row.recordType || ""),
    "Source Event: " + String(row.docketingEvent || ""),
    "Source Status: " + String(row.sourceStatus || "")
  ];
  if (row.sourceClient) parts.push("Source Client: " + row.sourceClient);
  if (row.referenceNumber) parts.push("Source Reference: " + row.referenceNumber);
  if (row.notes) parts.push("Import Notes: " + row.notes);
  return parts.join(" | ");
}

function bulkImportDocketTrakRows_(session, payload) {
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  payload = payload || {};
  var selectedClientCode = String(payload.clientCode || "").trim();
  var rows = payload.rows || [];
  if (!selectedClientCode) return { error: "Client code is required." };
  if (!rows.length) return { error: "No import rows supplied." };

  var client = findClientByCodeOrId_(selectedClientCode);
  if (!client) return { error: "Selected client not found." };
  if (String(client.STATUS || "Active") === "Deleted") return { error: "Selected client is inactive." };

  return withScriptLock_(function() {
    var existingCases = getAllCases() || [];
    var duplicateKeys = {};
    existingCases.forEach(function(caseItem) {
      if (caseItem.CLIENT_ID !== client.CLIENT_ID) return;
      var appKey = String(caseItem.APPLICATION_NUMBER || "").trim().toLowerCase();
      var titleKey = [
        String(caseItem.PATENT_TITLE || "").trim().toLowerCase(),
        String(caseItem.COUNTRY || "").trim().toLowerCase(),
        String(caseItem.PATENT_TYPE || "").trim().toLowerCase()
      ].join("|");
      if (appKey) duplicateKeys["app|" + appKey] = true;
      if (titleKey.replace(/\|/g, "")) duplicateKeys["title|" + titleKey] = true;
    });

    var imported = 0;
    var skipped = 0;
    var errors = [];
    var importedCaseIds = [];

    rows.forEach(function(row, index) {
      try {
        var recordType = normalizeImportRecordType_(row.recordType);
        var title = String(row.title || "").trim();
        var country = String(row.country || "").trim() || "India";
        var appNo = String(row.applicationNumber || row.referenceNumber || "").trim();
        var titleKey = [title.toLowerCase(), country.toLowerCase(), recordType.toLowerCase()].join("|");
        var appKey = appNo.toLowerCase();

        if (!recordType) {
          skipped++;
          errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: "Unsupported record type." });
          return;
        }
        if (!title) {
          skipped++;
          errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: "Missing title." });
          return;
        }
        if (appKey && duplicateKeys["app|" + appKey]) {
          skipped++;
          errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: "Duplicate application/reference number for selected client." });
          return;
        }
        if (!appKey && duplicateKeys["title|" + titleKey]) {
          skipped++;
          errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: "Duplicate title/country/type for selected client." });
          return;
        }

        var res = createCase({
          clientId: client.CLIENT_ID,
          patentTitle: title,
          applicationNumber: appNo,
          country: country,
          filingDate: "",
          nextDeadline: parseImportDateValue_(row.eventDate),
          status: "Drafted",
          patentType: recordType,
          attorney: "",
          priority: "Normal",
          orgId: client.ORG_ID || "",
          assignedStaffEmail: client.ASSIGNED_STAFF_EMAIL || "",
          galvanizerEmail: "",
          workflowStage: "Drafting",
          notes: buildImportedCaseNote_(row, selectedClientCode)
        });

        if (res && res.success) {
          imported++;
          importedCaseIds.push(res.caseId);
          if (appKey) duplicateKeys["app|" + appKey] = true;
          duplicateKeys["title|" + titleKey] = true;
        } else {
          skipped++;
          errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: (res && (res.error || res.message)) || "Import failed." });
        }
      } catch (e) {
        skipped++;
        errors.push({ rowNumber: row.sourceRowNumber || (index + 1), reason: e.message });
      }
    });

    if (imported) {
      try {
        logActivity_("BULK_IMPORT_CASES", "CASE", importedCaseIds.join(","), "Imported " + imported + " DocketTrak rows into " + client.CLIENT_ID);
      } catch (e) {}
      clearPortalCaches_(["cases", "casesPage", "dashboard", "dashboardSummary", "dashboardDetails", "documents", "galvanizerQueue", "workflowBoard", "smartSearch"]);
    }

    return {
      success: true,
      imported: imported,
      skipped: skipped,
      errors: sanitizeDataForFrontend_(errors),
      importedCaseIds: importedCaseIds,
      message: imported + " row(s) imported."
    };
  });
}

function getAccessibleDocumentsForUser_(session) {
  var docs = {};
  CONFIG.CLIENT_SUBFOLDERS.forEach(function(subName) { docs[subName] = []; });

  var clients = getAccessibleClientsForUser_(session);
  var showPrefix = canManageAllData_(session) || normalizeRole_(session.role) === "Staff";

  clients.forEach(function(client) {
    if (!client.CLIENT_FOLDER_ID) return;
    try {
      var clientFolder = DriveApp.getFolderById(client.CLIENT_FOLDER_ID);
      CONFIG.CLIENT_SUBFOLDERS.forEach(function(subName) {
        var subFolders = clientFolder.getFoldersByName(subName);
        if (!subFolders.hasNext()) return;
        var folder = subFolders.next();
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          var prefix = showPrefix ? "[" + (client.CLIENT_ID || client.CLIENT_NAME) + "] " : "";
          docs[subName].push({
            name: prefix + file.getName(),
            url: file.getUrl(),
            size: file.getSize(),
            date: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "dd-MMM-yyyy"),
            type: file.getMimeType(),
            clientId: client.CLIENT_ID,
            orgId: client.ORG_ID || ""
          });
        }
      });
    } catch (e) {
      Logger.log("Document access error for " + client.CLIENT_ID + ": " + e.message);
    }
  });

  return docs;
}

function canAccessMessageThreadForSession_(session, thread) {
  if (!thread || String(thread.STATUS || "") === "Deleted") return false;
  var role = normalizeRole_(session.role);
  var email = String(session.email || "").trim().toLowerCase();
  var sessionOrgId = getResolvedOrgIdForSession_(session);
  var sessionClientId = String(session.clientId || "").trim().toUpperCase();

  if (String(thread.THREAD_TYPE || "") === "Direct") {
    return getMessageParticipants_(thread.THREAD_ID).some(function(item) {
      return String(item.USER_EMAIL || "").trim().toLowerCase() === email;
    });
  }

  if (canManageAllData_(session)) return true;

  if (isClientSideSession_(session)) {
    if (String(thread.VISIBLE_TO_CLIENT || "No") !== "Yes") return false;
    if (sessionOrgId && String(thread.ORG_ID || "").trim().toUpperCase() === sessionOrgId) return true;
    if (sessionClientId && String(thread.CLIENT_ID || "").trim().toUpperCase() === sessionClientId) return true;
    return false;
  }

  if (role === "Staff") {
    var relatedClient = thread.CLIENT_ID ? getClientById(thread.CLIENT_ID) : null;
    return String(thread.CREATED_BY || "").trim().toLowerCase() === email ||
      String((relatedClient && relatedClient.ASSIGNED_STAFF_EMAIL) || "").trim().toLowerCase() === email;
  }
  if (role === "Attorney") {
    return String(thread.CREATED_BY || "").trim().toLowerCase() === email ||
      thread.RELATED_ENTITY_TYPE === "CASE";
  }
  if (sessionOrgId) return String(thread.ORG_ID || "").trim().toUpperCase() === sessionOrgId;
  if (sessionClientId) return String(thread.CLIENT_ID || "").trim().toUpperCase() === sessionClientId;
  return false;
}

function isClientVisibleMessage_(session, thread, message) {
  if (!isClientSideSession_(session)) return true;
  if (String(thread.THREAD_TYPE || "") === "Direct") return true;
  return String(message.IS_INTERNAL || "No") !== "Yes";
}

function listOrganizationsForSession_(session) {
  var cacheKey = buildPortalCacheKey_("organizations", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var orgs = getRecords_("ORGANIZATIONS");
  if (canManageAllData_(session)) return writePortalCache_(cacheKey, sanitizeDataForFrontend_(orgs), 300);
  if (session.orgId) {
    return writePortalCache_(cacheKey, sanitizeDataForFrontend_(orgs.filter(function(org) { return org.ORG_ID === session.orgId; })), 300);
  }
  return [];
}

function saveOrganization_(session, orgData) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };

  return withScriptLock_(function() {
    var orgSheet = getSheet_("ORGANIZATIONS");
    if (orgData.ORG_ID) {
      var ok = updateRecordById_("ORGANIZATIONS", "ORG_ID", orgData.ORG_ID, {
        ORG_NAME: orgData.ORG_NAME || "",
        PRIMARY_EMAIL: orgData.PRIMARY_EMAIL || "",
        PRIMARY_PHONE: orgData.PRIMARY_PHONE || "",
        ADDRESS: orgData.ADDRESS || "",
        ORG_CODE: normalizeClientCode_(orgData.ORG_CODE || "", orgData.CLIENT_REGION || "India"),
        CLIENT_ADMIN_USER_ID: orgData.CLIENT_ADMIN_USER_ID || "",
        ASSIGNED_STAFF_EMAIL: orgData.ASSIGNED_STAFF_EMAIL || "",
        STATUS: orgData.STATUS || "Active",
        NOTES: orgData.NOTES || ""
      });
      if (!ok) return { error: "Organization not found." };
      logActivity_("UPDATE_ORGANIZATION", "ORGANIZATION", orgData.ORG_ID, orgData.ORG_NAME || "");
      clearPortalCaches_(["organizations", "clients", "users", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage"]);
      return { success: true, orgId: orgData.ORG_ID, message: "Organization updated successfully." };
    }

    var orgId = generateId_("ORG", orgSheet, 0);
    appendRecord_("ORGANIZATIONS", {
      ORG_ID: orgId,
      ORG_NAME: orgData.ORG_NAME || "",
      PRIMARY_EMAIL: orgData.PRIMARY_EMAIL || "",
      PRIMARY_PHONE: orgData.PRIMARY_PHONE || "",
      ADDRESS: orgData.ADDRESS || "",
      ORG_CODE: normalizeClientCode_(orgData.ORG_CODE || "", orgData.CLIENT_REGION || "India"),
      CLIENT_ADMIN_USER_ID: orgData.CLIENT_ADMIN_USER_ID || "",
      ASSIGNED_STAFF_EMAIL: orgData.ASSIGNED_STAFF_EMAIL || "",
      STATUS: orgData.STATUS || "Active",
      DATE_CREATED: new Date(),
      NOTES: orgData.NOTES || ""
    });
    logActivity_("CREATE_ORGANIZATION", "ORGANIZATION", orgId, orgData.ORG_NAME || "");
    clearPortalCaches_(["organizations", "clients", "users", "dashboard", "dashboardSummary", "dashboardDetails", "cases", "casesPage"]);
    return { success: true, orgId: orgId, message: "Organization created successfully." };
  });
}

function getOrganizationUsers_(session, orgId) {
  if (!canAccessManagement_(session)) return { error: "Access denied." };
  var effectiveOrgId = orgId || session.orgId;
  if (!effectiveOrgId) return [];
  if (!canManageAllData_(session) && effectiveOrgId !== session.orgId) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("users", session, "org:" + effectiveOrgId);
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;

  var users = getRecords_("USERS").filter(function(user) {
    return user.ORG_ID === effectiveOrgId && String(user.STATUS || "") !== "Inactive";
  }).map(function(user) {
    delete user.PASSWORD_HASH;
    return user;
  });

  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(users), 300);
}

function saveDailyPriority_(session, data) {
  if (!isInternalRole_(session.role)) return { error: "Access denied." };

  return withScriptLock_(function() {
    var sheet = getSheet_("DAILY_PRIORITIES");
    var entryDate = data.entryDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var records = getRecords_("DAILY_PRIORITIES");
    var existing = null;
    for (var i = 0; i < records.length; i++) {
      if (records[i].USER_EMAIL === session.email && String(records[i].ENTRY_DATE) === String(entryDate)) {
        existing = records[i];
        break;
      }
    }

    var updates = {
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      ROLE: session.role,
      ENTRY_DATE: entryDate,
      PRIORITY_1: data.priority1 || "",
      PRIORITY_2: data.priority2 || "",
      PRIORITY_3: data.priority3 || "",
      NOTES: data.notes || "",
      STATUS: data.status || "Planned",
      UPDATED_AT: new Date()
    };

    if (existing) {
      updateRecordById_("DAILY_PRIORITIES", "ENTRY_ID", existing.ENTRY_ID, updates);
      logActivity_("UPDATE_DAILY_PRIORITY", "DAILY_PRIORITY", existing.ENTRY_ID, session.email);
      clearPortalCaches_(["dailyOps", "dailyAudit", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
      return { success: true, entryId: existing.ENTRY_ID, message: "Daily priorities updated." };
    }

    var entryId = generateId_("DP", sheet, 0);
    updates.ENTRY_ID = entryId;
    updates.CREATED_AT = new Date();
    appendRecord_("DAILY_PRIORITIES", updates);
    createNotificationsForRole_("Admin", "Daily priorities submitted", session.name + " submitted priorities for " + entryDate, "DAILY_PRIORITY", entryId);
    logActivity_("CREATE_DAILY_PRIORITY", "DAILY_PRIORITY", entryId, session.email);
    clearPortalCaches_(["dailyOps", "dailyAudit", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
    return { success: true, entryId: entryId, message: "Daily priorities submitted." };
  });
}

function saveDailyWrapup_(session, data) {
  if (!isInternalRole_(session.role)) return { error: "Access denied." };

  return withScriptLock_(function() {
    var sheet = getSheet_("DAILY_WRAPUPS");
    var entryDate = data.entryDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var records = getRecords_("DAILY_WRAPUPS");
    var existing = null;
    for (var i = 0; i < records.length; i++) {
      if (records[i].USER_EMAIL === session.email && String(records[i].ENTRY_DATE) === String(entryDate)) {
        existing = records[i];
        break;
      }
    }

    var updates = {
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      ROLE: session.role,
      ENTRY_DATE: entryDate,
      HIGH_POINTS: data.highPoints || "",
      LOW_POINTS: data.lowPoints || "",
      HELP_NEEDED: data.helpNeeded || "",
      ADMIN_REVIEW_STATUS: data.adminReviewStatus || "Pending Review",
      ADMIN_REVIEW_NOTES: data.adminReviewNotes || "",
      UPDATED_AT: new Date()
    };

    if (existing) {
      updateRecordById_("DAILY_WRAPUPS", "WRAPUP_ID", existing.WRAPUP_ID, updates);
      logActivity_("UPDATE_DAILY_WRAPUP", "DAILY_WRAPUP", existing.WRAPUP_ID, session.email);
      clearPortalCaches_(["dailyOps", "dailyAudit", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
      return { success: true, wrapupId: existing.WRAPUP_ID, message: "Day-end log updated." };
    }

    var wrapupId = generateId_("WU", sheet, 0);
    updates.WRAPUP_ID = wrapupId;
    updates.CREATED_AT = new Date();
    appendRecord_("DAILY_WRAPUPS", updates);
    createNotificationsForRole_("Admin", "Day-end log submitted", session.name + " submitted the day-end log.", "DAILY_WRAPUP", wrapupId);
    logActivity_("CREATE_DAILY_WRAPUP", "DAILY_WRAPUP", wrapupId, session.email);
    clearPortalCaches_(["dailyOps", "dailyAudit", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
    return { success: true, wrapupId: wrapupId, message: "Day-end log submitted." };
  });
}

function getDailyOpsOverview_(session) {
  if (!isInternalRole_(session.role) && !canManageAllData_(session)) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("dailyOps", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;

  var priorities = getRecords_("DAILY_PRIORITIES");
  var wrapups = getRecords_("DAILY_WRAPUPS");
  var claims = getRecords_("EXPENSE_CLAIMS");

  if (!canManageAllData_(session)) {
    priorities = priorities.filter(function(item) { return item.USER_EMAIL === session.email; });
    wrapups = wrapups.filter(function(item) { return item.USER_EMAIL === session.email; });
    claims = claims.filter(function(item) { return item.USER_EMAIL === session.email; });
  }

  return writePortalCache_(cacheKey, sanitizeDataForFrontend_({
    priorities: priorities,
    wrapups: wrapups,
    claims: claims
  }), 120);
}

function getDailyAuditView_(session, filters) {
  if (!isInternalRole_(getEffectiveRoles_(session))) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("dailyAudit", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var priorities = getRecords_("DAILY_PRIORITIES");
  var wrapups = getRecords_("DAILY_WRAPUPS");
  filters = filters || {};
  var dateFilter = String(filters.entryDate || "").trim();
  var userFilter = String(filters.userName || "").trim().toLowerCase();

  function match(item) {
    if (dateFilter && normalizeComparableDate_(item.ENTRY_DATE) !== dateFilter) return false;
    if (userFilter && String(item.USER_NAME || "").toLowerCase() !== userFilter) return false;
    return true;
  }

  if (!canManageAllData_(session)) {
    priorities = priorities.filter(function(item) { return item.USER_EMAIL === session.email; });
    wrapups = wrapups.filter(function(item) { return item.USER_EMAIL === session.email; });
  }

  priorities = priorities.filter(match);
  wrapups = wrapups.filter(match);
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_({ priorities: priorities, wrapups: wrapups }), 120);
}

function submitExpenseClaim_(session, data) {
  if (!isInternalRole_(session.role)) return { error: "Access denied." };

  return withScriptLock_(function() {
    var sheet = getSheet_("EXPENSE_CLAIMS");
    var claimId = data.CLAIM_ID || generateId_("EXP", sheet, 0);
    var submittedTo = data.submittedTo || "Super Admin";

    if (data.CLAIM_ID) {
      var ok = updateRecordById_("EXPENSE_CLAIMS", "CLAIM_ID", data.CLAIM_ID, {
        CLAIM_DATE: data.claimDate || "",
        CATEGORY: data.category || "",
        AMOUNT: data.amount || 0,
        DESCRIPTION: data.description || "",
        BILL_LINK: data.billLink || "",
        STATUS: data.status || "Submitted",
        SUBMITTED_TO: submittedTo,
        ADMIN_REMARKS: data.adminRemarks || "",
        UPDATED_AT: new Date()
      });
      if (!ok) return { error: "Expense claim not found." };
      clearPortalCaches_(["expenses", "dailyOps", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
      return { success: true, claimId: data.CLAIM_ID, message: "Expense claim updated." };
    }

    appendRecord_("EXPENSE_CLAIMS", {
      CLAIM_ID: claimId,
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      ROLE: session.role,
      CLAIM_DATE: data.claimDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      CATEGORY: data.category || "",
      AMOUNT: data.amount || 0,
      DESCRIPTION: data.description || "",
      BILL_LINK: data.billLink || "",
      STATUS: data.status || "Submitted",
      SUBMITTED_TO: submittedTo,
      ADMIN_REMARKS: "",
      CREATED_AT: new Date(),
      UPDATED_AT: new Date()
    });

    createNotificationsForRole_("Super Admin", "New expense claim", session.name + " submitted expense claim " + claimId, "EXPENSE_CLAIM", claimId);
    logActivity_("CREATE_EXPENSE_CLAIM", "EXPENSE_CLAIM", claimId, session.email);
    clearPortalCaches_(["expenses", "dailyOps", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
    return { success: true, claimId: claimId, message: "Expense claim submitted." };
  });
}

function getExpenseClaims_(session, filters) {
  var cacheKey = buildPortalCacheKey_("expenses", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var claims = getRecords_("EXPENSE_CLAIMS");
  filters = filters || {};
  var fromDate = normalizeComparableDate_(filters.fromDate || "");
  var toDate = normalizeComparableDate_(filters.toDate || "");
  if (String(filters.period || "") === "month") {
    var now = new Date();
    claims = claims.filter(function(claim) {
      var rawDate = claim.CLAIM_DATE || claim.CREATED_AT;
      var date = rawDate ? new Date(rawDate) : null;
      return date && !isNaN(date.getTime()) && date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
    });
  }
  if (fromDate || toDate) {
    claims = claims.filter(function(claim) {
      var normalized = normalizeComparableDate_(claim.CLAIM_DATE || claim.CREATED_AT);
      if (!normalized) return false;
      if (fromDate && normalized < fromDate) return false;
      if (toDate && normalized > toDate) return false;
      return true;
    });
  }
  if (canApproveExpenseClaims_(session) || canManageAllData_(session)) {
    return writePortalCache_(cacheKey, sanitizeDataForFrontend_(claims), 120);
  }
  if (isInternalRole_(session.role)) {
    return writePortalCache_(cacheKey, sanitizeDataForFrontend_(claims.filter(function(claim) { return claim.USER_EMAIL === session.email; })), 120);
  }
  return { error: "Access denied." };
}

function saveExpenseClaimReview_(session, claimId, reviewData) {
  if (!canApproveExpenseClaims_(session)) return { error: "Access denied." };
  var ok = updateRecordById_("EXPENSE_CLAIMS", "CLAIM_ID", claimId, {
    STATUS: reviewData.status || "Reviewed",
    ADMIN_REMARKS: reviewData.adminRemarks || "",
    UPDATED_AT: new Date()
  });
  if (!ok) return { error: "Expense claim not found." };
  var claim = null;
  var claims = getRecords_("EXPENSE_CLAIMS");
  for (var i = 0; i < claims.length; i++) {
    if (claims[i].CLAIM_ID === claimId) {
      claim = claims[i];
      break;
    }
  }
  if (claim) {
    createNotification_(claim.USER_EMAIL, "Expense claim updated", "Your expense claim " + claimId + " is now " + (reviewData.status || "Reviewed"), "EXPENSE_CLAIM", claimId);
  }
  clearPortalCaches_(["expenses", "dailyOps", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
  return { success: true, message: "Expense claim review saved." };
}

function getNotifications_(session) {
  var cacheKey = buildPortalCacheKey_("notifications", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var notifications = getRecords_("NOTIFICATIONS").filter(function(item) {
    return item.USER_EMAIL === session.email && String(item.STATUS || "Active") !== "Deleted";
  });
  notifications.sort(function(a, b) {
    return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0);
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(notifications.slice(0, 100)), 45);
}

function markNotificationRead_(session, notificationId) {
  var ok = updateRecordById_("NOTIFICATIONS", "NOTIFICATION_ID", notificationId, {
    IS_READ: "Yes",
    READ_AT: new Date()
  });
  if (!ok) return { error: "Notification not found." };
  clearPortalCaches_(["notifications", "dashboard", "dashboardSummary", "dashboardDetails"]);
  return { success: true };
}

function createNotification_(userEmail, title, body, entityType, entityId) {
  ensureHeaders_(getSheet_("NOTIFICATIONS"), [
    "NOTIFICATION_ID", "USER_EMAIL", "TITLE", "BODY", "TYPE",
    "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "IS_READ",
    "CREATED_AT", "READ_AT", "STATUS"
  ]);
  var sheet = getSheet_("NOTIFICATIONS");
  var notificationId = generateId_("NTF", sheet, 0);
  appendRecord_("NOTIFICATIONS", {
    NOTIFICATION_ID: notificationId,
    USER_EMAIL: userEmail,
    TITLE: title,
    BODY: body,
    TYPE: entityType || "GENERAL",
    RELATED_ENTITY_TYPE: entityType || "",
    RELATED_ENTITY_ID: entityId || "",
    IS_READ: "No",
    CREATED_AT: new Date(),
    READ_AT: "",
    STATUS: "Active"
  });
  clearPortalCaches_(["notifications", "dashboard", "dashboardSummary", "dashboardDetails"]);
}

function deleteNotification_(session, notificationId) {
  ensureHeaders_(getSheet_("NOTIFICATIONS"), [
    "NOTIFICATION_ID", "USER_EMAIL", "TITLE", "BODY", "TYPE",
    "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "IS_READ",
    "CREATED_AT", "READ_AT", "STATUS"
  ]);
  var notifications = getRecords_("NOTIFICATIONS");
  var item = notifications.find(function(row) {
    return row.NOTIFICATION_ID === notificationId && row.USER_EMAIL === session.email;
  });
  if (!item) return { error: "Notification not found." };
  var ok = updateRecordById_("NOTIFICATIONS", "NOTIFICATION_ID", notificationId, {
    STATUS: "Deleted",
    IS_READ: "Yes",
    READ_AT: new Date()
  });
  if (!ok) return { error: "Notification delete failed." };
  clearPortalCaches_(["notifications", "dashboard", "dashboardSummary", "dashboardDetails"]);
  return { success: true };
}

function clearNotifications_(session) {
  ensureHeaders_(getSheet_("NOTIFICATIONS"), [
    "NOTIFICATION_ID", "USER_EMAIL", "TITLE", "BODY", "TYPE",
    "RELATED_ENTITY_TYPE", "RELATED_ENTITY_ID", "IS_READ",
    "CREATED_AT", "READ_AT", "STATUS"
  ]);
  var notifications = getRecords_("NOTIFICATIONS").filter(function(item) {
    return item.USER_EMAIL === session.email && String(item.STATUS || "Active") !== "Deleted";
  });
  notifications.forEach(function(item) {
    updateRecordById_("NOTIFICATIONS", "NOTIFICATION_ID", item.NOTIFICATION_ID, {
      STATUS: "Deleted",
      IS_READ: "Yes",
      READ_AT: new Date()
    });
  });
  clearPortalCaches_(["notifications", "dashboard", "dashboardSummary", "dashboardDetails"]);
  return { success: true, cleared: notifications.length };
}

function decodeDataUrlToBlob_(dataUrl, fileName, mimeTypeFallback) {
  var raw = String(dataUrl || "");
  if (!raw) throw new Error("File data missing.");
  var parts = raw.split(",");
  if (parts.length < 2) throw new Error("Invalid file payload.");
  var meta = parts[0];
  var mimeMatch = meta.match(/data:(.*?);base64/);
  var mimeType = (mimeMatch && mimeMatch[1]) || mimeTypeFallback || "application/octet-stream";
  var bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, mimeType, fileName || ("upload_" + new Date().getTime()));
}

function uploadExpenseBill_(session, fileData) {
  if (!isInternalRole_(session.role)) return { error: "Access denied." };
  ensureHeaders_(getSheet_("EXPENSE_CLAIMS"), [
    "CLAIM_ID", "USER_EMAIL", "USER_NAME", "ROLE", "CLAIM_DATE",
    "CATEGORY", "AMOUNT", "DESCRIPTION", "BILL_LINK", "STATUS",
    "SUBMITTED_TO", "ADMIN_REMARKS", "CREATED_AT", "UPDATED_AT",
    "BILL_FILE_ID", "BILL_PREVIEW"
  ]);
  var safeName = String(fileData.fileName || "expense_bill").replace(/[^\w.\- ]/g, "_");
  var blob = decodeDataUrlToBlob_(fileData.dataUrl, safeName, fileData.mimeType);
  var folder = getFolder_("04_INVOICE_SYSTEM");
  var subFolders = folder.getFoldersByName("EXPENSE_BILLS");
  var targetFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder("EXPENSE_BILLS");
  var file = targetFolder.createFile(blob);
  return {
    success: true,
    fileUrl: file.getUrl(),
    fileId: file.getId(),
    fileName: file.getName()
  };
}

function uploadPortalDocument_(session, data) {
  if (!session) return { error: "Access denied." };
  var clientCode = String(data.clientCode || "").trim().toUpperCase();
  var accessibleClientMap = {};
  getAccessibleClientsForUser_(session).forEach(function(item) {
    accessibleClientMap[String(item.CLIENT_ID || "").trim().toUpperCase()] = true;
    accessibleClientMap[String(item.CLIENT_CODE || "").trim().toUpperCase()] = true;
  });
  var client = getAllClients().find(function(item) {
    return String(item.CLIENT_CODE || item.CLIENT_ID || "").toUpperCase() === clientCode || String(item.CLIENT_ID || "").toUpperCase() === clientCode;
  });
  if (!client) return { error: "Client not found." };
  if (!canManageAllData_(session) && !accessibleClientMap[String(client.CLIENT_ID || "").trim().toUpperCase()] && !accessibleClientMap[String(client.CLIENT_CODE || "").trim().toUpperCase()]) {
    return { error: "Access denied." };
  }
  if (!client.CLIENT_FOLDER_ID) return { error: "Client folder not available." };
  var fileName = String(data.fileName || "document").replace(/[^\w.\- ]/g, "_");
  var blob = decodeDataUrlToBlob_(data.dataUrl, fileName, data.mimeType);
  var upload = uploadToClientFolder(client.CLIENT_ID, data.subfolderName || "COMMUNICATION", blob, fileName);
  if (upload && upload.success) {
    createTimelineEvent_({
      EVENT_TYPE: "DOCUMENT_UPLOADED",
      TITLE: fileName,
      DESCRIPTION: "Document uploaded by " + session.name,
      ENTITY_TYPE: "DOCUMENT",
      ENTITY_ID: upload.fileUrl,
      CLIENT_ID: client.CLIENT_ID || "",
      ORG_ID: client.ORG_ID || "",
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      VISIBILITY: "Shared"
    });
    clearPortalCaches_(["documents", "documentRequests", "timeline", "notifications"]);
  }
  return upload;
}

function createNotificationsForRole_(role, title, body, entityType, entityId) {
  var users = getRecords_("USERS");
  users.forEach(function(user) {
    if (normalizeRole_(user.ROLE) === normalizeRole_(role) && String(user.STATUS || "") === "Active") {
      createNotification_(user.EMAIL, title, body, entityType, entityId);
    }
  });
}

function getMessageThreads_(session) {
  var cacheKey = buildPortalCacheKey_("messages", session, "threads");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var threads = getRecords_("MESSAGE_THREADS");
  var visibleThreads = threads.filter(function(thread) {
    return canAccessMessageThreadForSession_(session, thread);
  });

  visibleThreads.sort(function(a, b) {
    return new Date(b.LAST_MESSAGE_AT || b.CREATED_AT || 0) - new Date(a.LAST_MESSAGE_AT || a.CREATED_AT || 0);
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(visibleThreads), 45);
}

function deleteMessageThread_(session, threadId) {
  if (!threadId) return { error: "Thread ID is required." };
  var visibleThreads = getMessageThreads_(session);
  var exists = visibleThreads.some(function(thread) { return thread.THREAD_ID === threadId; });
  if (!exists) return { error: "Thread not found or access denied." };
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  var ok = updateRecordById_("MESSAGE_THREADS", "THREAD_ID", threadId, {
    STATUS: "Deleted",
    LAST_MESSAGE_AT: new Date()
  });
  if (ok) clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
  return ok ? { success: true, message: "Thread deleted." } : { error: "Thread not found." };
}

function getMessagesForThread_(session, threadId) {
  var threads = getMessageThreads_(session);
  var thread = null;
  for (var i = 0; i < threads.length; i++) {
    if (threads[i].THREAD_ID === threadId) {
      thread = threads[i];
      break;
    }
  }
  if (!thread) return { error: "Access denied." };
  var messages = getRecords_("MESSAGES").filter(function(item) {
    return item.THREAD_ID === threadId && isClientVisibleMessage_(session, thread, item);
  });
  return sanitizeDataForFrontend_(messages);
}

function parseReadBy_(value) {
  return String(value || "")
    .split(",")
    .map(function(item) { return String(item || "").trim().toLowerCase(); })
    .filter(function(item) { return !!item; });
}

function markThreadRead_(session, threadId) {
  if (!threadId) return { error: "Thread ID is required." };
  var allowed = getMessageThreads_(session).some(function(thread) { return thread.THREAD_ID === threadId; });
  if (!allowed) return { error: "Access denied." };
  var sheet = getSheet_("MESSAGES");
  var data = sheet.getDataRange().getValues();
  if (!data.length) return { success: true };
  var headers = data[0];
  var threadIdx = headers.indexOf("THREAD_ID");
  var readIdx = headers.indexOf("READ_BY");
  var senderIdx = headers.indexOf("SENDER_EMAIL");
  if (threadIdx === -1 || readIdx === -1) return { success: true };
  var email = String(session.email || "").toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (data[i][threadIdx] !== threadId) continue;
    if (senderIdx > -1 && String(data[i][senderIdx] || "").toLowerCase() === email) continue;
    var readers = parseReadBy_(data[i][readIdx]);
    if (readers.indexOf(email) > -1) continue;
    readers.push(email);
    sheet.getRange(i + 1, readIdx + 1).setValue(readers.join(","));
  }
  clearRequestSheetCache_("MESSAGES");
  clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
  return { success: true };
}

function saveMessageThread_(session, data) {
  return withScriptLock_(function() {
    var sheet = getSheet_("MESSAGE_THREADS");
    var threadId = data.THREAD_ID || generateId_("THR", sheet, 0);

    if (data.THREAD_ID) {
      updateRecordById_("MESSAGE_THREADS", "THREAD_ID", threadId, {
        TITLE: data.TITLE || "",
        THREAD_TYPE: data.THREAD_TYPE || "General",
        RELATED_ENTITY_TYPE: data.RELATED_ENTITY_TYPE || "",
        RELATED_ENTITY_ID: data.RELATED_ENTITY_ID || "",
        STATUS: data.STATUS || "Open",
        LAST_MESSAGE_AT: new Date()
      });
      clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
      return { success: true, threadId: threadId, message: "Thread updated." };
    }

    appendRecord_("MESSAGE_THREADS", {
      THREAD_ID: threadId,
      THREAD_TYPE: data.THREAD_TYPE || "General",
      TITLE: data.TITLE || "Untitled Thread",
      RELATED_ENTITY_TYPE: data.RELATED_ENTITY_TYPE || "",
      RELATED_ENTITY_ID: data.RELATED_ENTITY_ID || "",
      ORG_ID: data.ORG_ID || session.orgId || "",
      CLIENT_ID: data.CLIENT_ID || session.clientId || "",
      VISIBLE_TO_CLIENT: data.VISIBLE_TO_CLIENT || "No",
      CREATED_BY: session.email,
      LAST_MESSAGE_AT: new Date(),
      STATUS: data.STATUS || "Open",
      CREATED_AT: new Date()
    });
    logActivity_("CREATE_MESSAGE_THREAD", "MESSAGE_THREAD", threadId, session.email);
    clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
    return { success: true, threadId: threadId, message: "Thread created." };
  });
}

function sendThreadMessage_(session, data) {
  return withScriptLock_(function() {
    var threadId = data.threadId;
    if (!threadId) return { error: "Thread ID is required." };
    var accessibleThreads = getMessageThreads_(session);
    var thread = null;
    for (var i = 0; i < accessibleThreads.length; i++) {
      if (accessibleThreads[i].THREAD_ID === threadId) {
        thread = accessibleThreads[i];
        break;
      }
    }
    if (!thread) return { error: "Access denied." };

    var messageSheet = getSheet_("MESSAGES");
    var messageId = generateId_("MSG", messageSheet, 0);
    appendRecord_("MESSAGES", {
      MESSAGE_ID: messageId,
      THREAD_ID: threadId,
      SENDER_EMAIL: session.email,
      SENDER_NAME: session.name,
      SENDER_ROLE: session.role,
      MESSAGE_TEXT: data.messageText || "",
      IS_INTERNAL: data.isInternal || "No",
      CREATED_AT: new Date(),
      READ_BY: session.email
    });
    updateRecordById_("MESSAGE_THREADS", "THREAD_ID", threadId, {
      LAST_MESSAGE_AT: new Date(),
      STATUS: data.threadStatus || "Open"
    });

    var recipients = resolveThreadRecipients_(session, thread);

    recipients.forEach(function(email) {
      createNotification_(email, "New message in " + (thread.TITLE || threadId), session.name + " sent a new message.", "MESSAGE_THREAD", threadId);
    });

    logActivity_("SEND_MESSAGE", "MESSAGE", messageId, session.email + " -> " + threadId);
    clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
    return { success: true, messageId: messageId, message: "Message sent." };
  });
}

function createTimelineEvent_(eventData) {
  var sheet = getSheet_("ACTIVITY_TIMELINE");
  var eventId = generateId_("EVT", sheet, 0);
  appendRecord_("ACTIVITY_TIMELINE", {
    EVENT_ID: eventId,
    EVENT_TYPE: eventData.EVENT_TYPE || "GENERAL",
    TITLE: eventData.TITLE || "",
    DESCRIPTION: eventData.DESCRIPTION || "",
    ENTITY_TYPE: eventData.ENTITY_TYPE || "",
    ENTITY_ID: eventData.ENTITY_ID || "",
    CLIENT_ID: eventData.CLIENT_ID || "",
    ORG_ID: eventData.ORG_ID || "",
    CASE_ID: eventData.CASE_ID || "",
    USER_EMAIL: eventData.USER_EMAIL || "",
    USER_NAME: eventData.USER_NAME || "",
    VISIBILITY: eventData.VISIBILITY || "Internal",
    CREATED_AT: new Date()
  });
  clearPortalCaches_(["timeline", "dashboard", "dashboardSummary", "dashboardDetails"]);
  return eventId;
}

function getMessageParticipants_(threadId) {
  return getRecords_("MESSAGE_PARTICIPANTS").filter(function(item) {
    return item.THREAD_ID === threadId && String(item.STATUS || "Active") !== "Removed";
  });
}

function saveMessageParticipants_(threadId, participants) {
  participants = participants || [];
  var existing = getRecords_("MESSAGE_PARTICIPANTS");
  var now = new Date();
  participants.forEach(function(item) {
    if (!item || !item.USER_EMAIL) return;
    var found = existing.find(function(p) {
      return p.THREAD_ID === threadId && String(p.USER_EMAIL || "").toLowerCase() === String(item.USER_EMAIL || "").toLowerCase();
    });
    if (found) {
      updateRecordById_("MESSAGE_PARTICIPANTS", "PARTICIPANT_ID", found.PARTICIPANT_ID, {
        USER_NAME: item.USER_NAME || found.USER_NAME || "",
        USER_ROLE: item.USER_ROLE || found.USER_ROLE || "",
        PARTICIPANT_TYPE: item.PARTICIPANT_TYPE || found.PARTICIPANT_TYPE || "Direct",
        STATUS: "Active",
        UPDATED_AT: now
      });
      return;
    }
    var partId = generateId_("PAR", getSheet_("MESSAGE_PARTICIPANTS"), 0);
    appendRecord_("MESSAGE_PARTICIPANTS", {
      PARTICIPANT_ID: partId,
      THREAD_ID: threadId,
      USER_EMAIL: item.USER_EMAIL,
      USER_NAME: item.USER_NAME || "",
      USER_ROLE: item.USER_ROLE || "",
      PARTICIPANT_TYPE: item.PARTICIPANT_TYPE || "Direct",
      STATUS: "Active",
      CREATED_AT: now,
      UPDATED_AT: now
    });
  });
}

function createDirectThread_(session, data) {
  if (!isInternalRole_(getEffectiveRoles_(session))) return { error: "Access denied." };
  var recipientEmail = String(data.recipientEmail || "").trim().toLowerCase();
  if (!recipientEmail) return { error: "Recipient email is required." };
  if (recipientEmail === String(session.email || "").toLowerCase()) return { error: "You cannot start a direct chat with yourself." };
  var recipient = getUserRecordByEmail_(recipientEmail);
  if (!recipient || String(recipient.STATUS || "") !== "Active") return { error: "Recipient not found." };

  var existing = getRecords_("MESSAGE_THREADS").find(function(thread) {
    if (String(thread.STATUS || "") === "Deleted" || String(thread.THREAD_TYPE || "") !== "Direct") return false;
    var participants = getMessageParticipants_(thread.THREAD_ID).map(function(item) {
      return String(item.USER_EMAIL || "").toLowerCase();
    }).sort();
    var expected = [String(session.email || "").toLowerCase(), recipientEmail].sort();
    return participants.length === 2 && participants[0] === expected[0] && participants[1] === expected[1];
  });
  if (existing) return { success: true, threadId: existing.THREAD_ID, message: "Direct thread already exists." };

  var result = saveMessageThread_(session, {
    TITLE: data.title || ("Direct: " + (recipient.FULL_NAME || recipient.EMAIL)),
    THREAD_TYPE: "Direct",
    STATUS: "Open",
    VISIBLE_TO_CLIENT: "No"
  });
  if (result && result.success) {
    saveMessageParticipants_(result.threadId, [
      { USER_EMAIL: session.email, USER_NAME: session.name, USER_ROLE: session.role, PARTICIPANT_TYPE: "Direct" },
      { USER_EMAIL: recipient.EMAIL, USER_NAME: recipient.FULL_NAME, USER_ROLE: recipient.ROLE, PARTICIPANT_TYPE: "Direct" }
    ]);
    createNotification_(recipient.EMAIL, "New direct message", session.name + " started a direct conversation.", "MESSAGE_THREAD", result.threadId);
  }
  clearPortalCaches_(["messages", "notifications", "directInbox"]);
  return result;
}

function getDirectInbox_(session) {
  if (!isInternalRole_(getEffectiveRoles_(session))) return { error: "Access denied." };
  var cacheKey = buildPortalCacheKey_("directInbox", session, "all");
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var email = String(session.email || "").toLowerCase();
  var threads = getMessageThreads_(session).filter(function(thread) {
    return String(thread.THREAD_TYPE || "") === "Direct" && String(thread.STATUS || "") !== "Deleted";
  });
  var messages = getRecords_("MESSAGES");
  var inbox = threads.map(function(thread) {
    var participants = getMessageParticipants_(thread.THREAD_ID);
    var counterpart = participants.find(function(item) {
      return String(item.USER_EMAIL || "").toLowerCase() !== email;
    }) || null;
    var threadMessages = messages.filter(function(item) {
      return item.THREAD_ID === thread.THREAD_ID;
    }).sort(function(a, b) {
      return new Date(a.CREATED_AT || 0) - new Date(b.CREATED_AT || 0);
    });
    var lastMessage = threadMessages.length ? threadMessages[threadMessages.length - 1] : null;
    var unreadCount = threadMessages.filter(function(item) {
      if (String(item.SENDER_EMAIL || "").toLowerCase() === email) return false;
      return parseReadBy_(item.READ_BY).indexOf(email) === -1;
    }).length;
    return {
      THREAD_ID: thread.THREAD_ID,
      TITLE: thread.TITLE || "",
      THREAD_TYPE: "Direct",
      STATUS: thread.STATUS || "Open",
      LAST_MESSAGE_AT: thread.LAST_MESSAGE_AT || thread.CREATED_AT,
      counterpartEmail: counterpart ? counterpart.USER_EMAIL || "" : "",
      counterpartName: counterpart ? counterpart.USER_NAME || counterpart.USER_EMAIL || "Unknown User" : "Unknown User",
      counterpartRole: counterpart ? counterpart.USER_ROLE || "" : "",
      lastMessageText: lastMessage ? lastMessage.MESSAGE_TEXT || "" : "",
      lastSenderEmail: lastMessage ? lastMessage.SENDER_EMAIL || "" : "",
      unreadCount: unreadCount
    };
  }).sort(function(a, b) {
    return new Date(b.LAST_MESSAGE_AT || 0) - new Date(a.LAST_MESSAGE_AT || 0);
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(inbox), 45);
}

function resolveThreadRecipients_(session, thread) {
  var participants = getMessageParticipants_(thread.THREAD_ID);
  if (participants.length) {
    return participants
      .map(function(item) { return item.USER_EMAIL; })
      .filter(function(email) { return String(email || "").toLowerCase() !== String(session.email || "").toLowerCase(); });
  }
  if (thread.ORG_ID) {
    return getRecords_("USERS").filter(function(user) {
      return user.ORG_ID === thread.ORG_ID && user.EMAIL !== session.email && String(user.STATUS || "") === "Active";
    }).map(function(user) { return user.EMAIL; });
  }
  if (thread.CLIENT_ID) {
    return getRecords_("USERS").filter(function(user) {
      return user.CLIENT_ID === thread.CLIENT_ID && user.EMAIL !== session.email && String(user.STATUS || "") === "Active";
    }).map(function(user) { return user.EMAIL; });
  }
  return getRecords_("USERS").filter(function(user) {
    return isInternalRole_(user.ROLE) && user.EMAIL !== session.email && String(user.STATUS || "") === "Active";
  }).map(function(user) { return user.EMAIL; });
}

var _ORIGINAL_SEND_THREAD_MESSAGE_ = sendThreadMessage_;
sendThreadMessage_ = function(session, data) {
  var result = _ORIGINAL_SEND_THREAD_MESSAGE_(session, data);
  if (!(result && result.success)) return result;
  createTimelineEvent_({
    EVENT_TYPE: "MESSAGE",
    TITLE: "Message sent",
    DESCRIPTION: session.name + " sent a message to thread " + data.threadId,
    ENTITY_TYPE: "MESSAGE_THREAD",
    ENTITY_ID: data.threadId,
    USER_EMAIL: session.email,
    USER_NAME: session.name,
    VISIBILITY: String(data.isInternal || "No") === "Yes" ? "Internal" : "Shared"
  });
  clearPortalCaches_(["messages", "notifications", "dashboard", "dashboardSummary", "dashboardDetails", "directInbox"]);
  return result;
};

function getTasks_(session, filters) {
  var cacheKey = buildPortalCacheKey_("tasks", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  filters = filters || {};
  var tasks = getRecords_("TASKS").filter(function(task) {
    if (String(task.STATUS || "") === "Deleted") return false;
    if (canManageAllData_(session)) return true;
    if (String(task.ASSIGNED_TO_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    if (String(task.ASSIGNED_BY_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    if (session.orgId && task.ORG_ID === session.orgId) return true;
    if (session.clientId && task.CLIENT_ID === session.clientId) return true;
    return false;
  });
  if (filters.status) tasks = tasks.filter(function(item) { return String(item.STATUS || "") === String(filters.status); });
  if (filters.assignedTo) tasks = tasks.filter(function(item) { return String(item.ASSIGNED_TO_EMAIL || "").toLowerCase() === String(filters.assignedTo || "").toLowerCase(); });
  if (filters.relatedEntityId) tasks = tasks.filter(function(item) { return item.RELATED_ENTITY_ID === filters.relatedEntityId; });
  if (filters.query) {
    var query = String(filters.query).toLowerCase();
    tasks = tasks.filter(function(item) {
      return [item.TITLE, item.DESCRIPTION, item.ASSIGNED_TO_NAME, item.RELATED_ENTITY_ID, item.CLIENT_ID, item.TAGS].join(" ").toLowerCase().indexOf(query) > -1;
    });
  }
  tasks.sort(function(a, b) { return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0); });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(tasks), 60);
}

function saveTask_(session, data) {
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  return withScriptLock_(function() {
    var now = new Date();
    if (data.TASK_ID) {
      var ok = updateRecordById_("TASKS", "TASK_ID", data.TASK_ID, {
        TITLE: data.title || "",
        DESCRIPTION: data.description || "",
        STATUS: data.status || "Open",
        PRIORITY: data.priority || "Normal",
        ASSIGNED_TO_EMAIL: data.assignedToEmail || "",
        ASSIGNED_TO_NAME: data.assignedToName || "",
        RELATED_ENTITY_TYPE: data.relatedEntityType || "",
        RELATED_ENTITY_ID: data.relatedEntityId || "",
        CLIENT_ID: data.clientId || "",
        ORG_ID: data.orgId || "",
        DUE_DATE: data.dueDate || "",
        START_DATE: data.startDate || "",
        TAGS: data.tags || "",
        NOTES: data.notes || "",
        UPDATED_AT: now
      });
      if (!ok) return { error: "Task not found." };
      createTimelineEvent_({
        EVENT_TYPE: "TASK_UPDATED",
        TITLE: "Task updated",
        DESCRIPTION: (data.title || data.TASK_ID) + " updated by " + session.name,
        ENTITY_TYPE: "TASK",
        ENTITY_ID: data.TASK_ID,
        CLIENT_ID: data.clientId || "",
        ORG_ID: data.orgId || "",
        USER_EMAIL: session.email,
        USER_NAME: session.name,
        VISIBILITY: "Internal"
      });
      clearPortalCaches_(["tasks", "timeline", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
      return { success: true, taskId: data.TASK_ID, message: "Task updated." };
    }
    var taskId = generateId_("TSK", getSheet_("TASKS"), 0);
    appendRecord_("TASKS", {
      TASK_ID: taskId,
      TITLE: data.title || "",
      DESCRIPTION: data.description || "",
      STATUS: data.status || "Open",
      PRIORITY: data.priority || "Normal",
      ASSIGNED_TO_EMAIL: data.assignedToEmail || "",
      ASSIGNED_TO_NAME: data.assignedToName || "",
      ASSIGNED_BY_EMAIL: session.email,
      ASSIGNED_BY_NAME: session.name,
      RELATED_ENTITY_TYPE: data.relatedEntityType || "",
      RELATED_ENTITY_ID: data.relatedEntityId || "",
      CLIENT_ID: data.clientId || "",
      ORG_ID: data.orgId || "",
      DUE_DATE: data.dueDate || "",
      START_DATE: data.startDate || "",
      COMPLETED_AT: "",
      TAGS: data.tags || "",
      NOTES: data.notes || "",
      CREATED_AT: now,
      UPDATED_AT: now
    });
    if (data.assignedToEmail) {
      createNotification_(data.assignedToEmail, "New task assigned", session.name + " assigned task " + (data.title || taskId), "TASK", taskId);
    }
    createTimelineEvent_({
      EVENT_TYPE: "TASK_CREATED",
      TITLE: "Task created",
      DESCRIPTION: (data.title || taskId) + " created by " + session.name,
      ENTITY_TYPE: "TASK",
      ENTITY_ID: taskId,
      CLIENT_ID: data.clientId || "",
      ORG_ID: data.orgId || "",
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      VISIBILITY: "Internal"
    });
    clearPortalCaches_(["tasks", "timeline", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
    return { success: true, taskId: taskId, message: "Task created." };
  });
}

function updateTaskStatus_(session, taskId, status, notes) {
  var task = getRecords_("TASKS").find(function(item) { return item.TASK_ID === taskId; });
  if (!task) return { error: "Task not found." };
  if (!canManageAllData_(session) && String(task.ASSIGNED_TO_EMAIL || "").toLowerCase() !== String(session.email || "").toLowerCase() && String(task.ASSIGNED_BY_EMAIL || "").toLowerCase() !== String(session.email || "").toLowerCase()) {
    return { error: "Access denied." };
  }
  var ok = updateRecordById_("TASKS", "TASK_ID", taskId, {
    STATUS: status || "Open",
    COMPLETED_AT: status === "Completed" ? new Date() : "",
    NOTES: notes || task.NOTES || "",
    UPDATED_AT: new Date()
  });
  if (!ok) return { error: "Task update failed." };
  createTimelineEvent_({
    EVENT_TYPE: "TASK_STATUS",
    TITLE: "Task status updated",
    DESCRIPTION: task.TITLE + " moved to " + status,
    ENTITY_TYPE: "TASK",
    ENTITY_ID: taskId,
    CLIENT_ID: task.CLIENT_ID || "",
    ORG_ID: task.ORG_ID || "",
    USER_EMAIL: session.email,
    USER_NAME: session.name,
    VISIBILITY: "Internal"
  });
  clearPortalCaches_(["tasks", "timeline", "dashboard", "dashboardSummary", "dashboardDetails", "notifications"]);
  return { success: true };
}

function getActivityTimeline_(session, filters) {
  var cacheKey = buildPortalCacheKey_("timeline", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  filters = filters || {};
  var events = getRecords_("ACTIVITY_TIMELINE").filter(function(item) {
    if (canManageAllData_(session)) return true;
    if (item.VISIBILITY === "Internal" && !isInternalRole_(getEffectiveRoles_(session))) return false;
    if (session.orgId && item.ORG_ID === session.orgId) return true;
    if (session.clientId && item.CLIENT_ID === session.clientId) return true;
    if (String(item.USER_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    return isInternalRole_(getEffectiveRoles_(session));
  });
  if (filters.entityType) events = events.filter(function(item) { return item.ENTITY_TYPE === filters.entityType; });
  if (filters.entityId) events = events.filter(function(item) { return item.ENTITY_ID === filters.entityId || item.CASE_ID === filters.entityId || item.CLIENT_ID === filters.entityId; });
  events.sort(function(a, b) { return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0); });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(events.slice(0, 200)), 60);
}

function getWorkflowBoard_(session, filters) {
  var cacheKey = buildPortalCacheKey_("workflowBoard", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  var stages = ["Drafting", "Ready for Attorney", "Under Attorney Review", "Filed", "Under Examination", "Granted", "Closed"];
  var cases = filterAccessibleCases_(session, filters || {});
  var board = {};
  stages.forEach(function(stage) { board[stage] = []; });
  cases.forEach(function(caseItem) {
    var stage = String(caseItem.WORKFLOW_STAGE || "Drafting");
    if (!board[stage]) board[stage] = [];
    board[stage].push(caseItem);
  });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(board), 60);
}

function getAttorneyWorkspace_(session, filters) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Attorney"])) return { error: "Access denied." };
  filters = filters || {};
  var cases = getAccessibleCasesForUser_(session).filter(function(caseItem) {
    return String(caseItem.ATTORNEY || "").toLowerCase() === String(session.email || "").toLowerCase() || hasRoleAtLeast_(getEffectiveRoles_(session), "Admin");
  });
  if (filters.status) cases = cases.filter(function(item) { return String(item.CURRENT_STATUS || "") === String(filters.status); });
  if (filters.stage) cases = cases.filter(function(item) { return String(item.WORKFLOW_STAGE || "") === String(filters.stage); });
  var tasks = getTasks_(session, { assignedTo: session.email });
  return sanitizeDataForFrontend_({
    cases: cases,
    pendingReview: cases.filter(function(item) { return String(item.WORKFLOW_STAGE || "") === "Ready for Attorney"; }),
    activeReview: cases.filter(function(item) { return String(item.WORKFLOW_STAGE || "") === "Under Attorney Review"; }),
    tasks: tasks && !tasks.error ? tasks : []
  });
}

function getSmartSearch_(session, query, scope) {
  query = String(query || "").trim().toLowerCase();
  if (!query) return sanitizeDataForFrontend_({ cases: [], clients: [], users: [], tasks: [], threads: [] });
  scope = scope || "all";
  var out = { cases: [], clients: [], users: [], tasks: [], threads: [] };
  if (scope === "all" || scope === "cases") {
    out.cases = filterAccessibleCases_(session, { query: query }).slice(0, 20);
  }
  if (scope === "all" || scope === "clients") {
    out.clients = getAccessibleClientsForUser_(session).filter(function(item) {
      return [item.CLIENT_ID, item.CLIENT_CODE, item.CLIENT_NAME, item.EMAIL, item.CONTACT_PERSON].join(" ").toLowerCase().indexOf(query) > -1;
    }).slice(0, 20);
  }
  if (scope === "all" || scope === "users") {
    if (canAccessManagement_(session)) {
      out.users = getRecords_("USERS").filter(function(item) {
        return [item.USER_ID, item.FULL_NAME, item.EMAIL, item.ROLE].join(" ").toLowerCase().indexOf(query) > -1;
      }).slice(0, 20);
    }
  }
  if (scope === "all" || scope === "tasks") {
    var tasks = getTasks_(session, { query: query });
    out.tasks = tasks && !tasks.error ? tasks.slice(0, 20) : [];
  }
  if (scope === "all" || scope === "messages") {
    var threads = getMessageThreads_(session);
    out.threads = (threads || []).filter(function(item) {
      return [item.THREAD_ID, item.TITLE, item.RELATED_ENTITY_ID, item.RELATED_ENTITY_TYPE].join(" ").toLowerCase().indexOf(query) > -1;
    }).slice(0, 20);
  }
  return sanitizeDataForFrontend_(out);
}

function getApprovalRequests_(session, filters) {
  var cacheKey = buildPortalCacheKey_("approvals", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  filters = filters || {};
  var approvals = getRecords_("APPROVALS").filter(function(item) {
    if (canManageAllData_(session)) return true;
    if (String(item.APPROVER_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    if (String(item.REQUESTED_BY_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    if (session.orgId && item.ORG_ID === session.orgId) return true;
    return false;
  });
  if (filters.status) approvals = approvals.filter(function(item) { return String(item.STATUS || "") === String(filters.status); });
  approvals.sort(function(a, b) { return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0); });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(approvals), 60);
}

function saveApprovalRequest_(session, data) {
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  return withScriptLock_(function() {
    var approvalId = generateId_("APR", getSheet_("APPROVALS"), 0);
    appendRecord_("APPROVALS", {
      APPROVAL_ID: approvalId,
      APPROVAL_TYPE: data.approvalType || "General",
      TITLE: data.title || "",
      DESCRIPTION: data.description || "",
      STATUS: "Pending",
      REQUESTED_BY_EMAIL: session.email,
      REQUESTED_BY_NAME: session.name,
      APPROVER_EMAIL: data.approverEmail || "",
      APPROVER_ROLE: data.approverRole || "",
      RELATED_ENTITY_TYPE: data.relatedEntityType || "",
      RELATED_ENTITY_ID: data.relatedEntityId || "",
      CLIENT_ID: data.clientId || "",
      ORG_ID: data.orgId || "",
      REQUEST_DATE: new Date(),
      DECISION_DATE: "",
      DECISION_NOTES: "",
      CREATED_AT: new Date(),
      UPDATED_AT: new Date()
    });
    if (data.approverEmail) {
      createNotification_(data.approverEmail, "Approval request pending", session.name + " requested approval for " + (data.title || approvalId), "APPROVAL", approvalId);
    }
    createTimelineEvent_({
      EVENT_TYPE: "APPROVAL_REQUESTED",
      TITLE: "Approval requested",
      DESCRIPTION: (data.title || approvalId) + " requested by " + session.name,
      ENTITY_TYPE: "APPROVAL",
      ENTITY_ID: approvalId,
      CLIENT_ID: data.clientId || "",
      ORG_ID: data.orgId || "",
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      VISIBILITY: "Internal"
    });
    clearPortalCaches_(["approvals", "timeline", "notifications"]);
    return { success: true, approvalId: approvalId };
  });
}

function reviewApprovalRequest_(session, approvalId, reviewData) {
  var approval = getRecords_("APPROVALS").find(function(item) { return item.APPROVAL_ID === approvalId; });
  if (!approval) return { error: "Approval request not found." };
  if (!canManageAllData_(session) && String(approval.APPROVER_EMAIL || "").toLowerCase() !== String(session.email || "").toLowerCase()) {
    return { error: "Access denied." };
  }
  var status = reviewData.status || "Approved";
  var ok = updateRecordById_("APPROVALS", "APPROVAL_ID", approvalId, {
    STATUS: status,
    DECISION_DATE: new Date(),
    DECISION_NOTES: reviewData.notes || "",
    UPDATED_AT: new Date()
  });
  if (!ok) return { error: "Approval update failed." };
  createNotification_(approval.REQUESTED_BY_EMAIL, "Approval " + status.toLowerCase(), "Your approval request " + approvalId + " is " + status.toLowerCase(), "APPROVAL", approvalId);
  createTimelineEvent_({
    EVENT_TYPE: "APPROVAL_" + String(status || "").toUpperCase(),
    TITLE: "Approval " + status,
    DESCRIPTION: approval.TITLE + " was " + status.toLowerCase(),
    ENTITY_TYPE: "APPROVAL",
    ENTITY_ID: approvalId,
    CLIENT_ID: approval.CLIENT_ID || "",
    ORG_ID: approval.ORG_ID || "",
    USER_EMAIL: session.email,
    USER_NAME: session.name,
    VISIBILITY: "Internal"
  });
  clearPortalCaches_(["approvals", "timeline", "notifications"]);
  return { success: true };
}

function getDocumentRequests_(session, filters) {
  var cacheKey = buildPortalCacheKey_("documentRequests", session, JSON.stringify(filters || {}));
  var cached = readPortalCache_(cacheKey);
  if (cached) return cached;
  filters = filters || {};
  var requests = getRecords_("DOCUMENT_REQUESTS").filter(function(item) {
    if (canManageAllData_(session)) return true;
    var sessionOrgId = getResolvedOrgIdForSession_(session);
    var sessionClientId = String(session.clientId || "").trim().toUpperCase();
    if (isClientSideSession_(session) && String(item.CLIENT_VISIBLE || "Yes") !== "Yes") return false;
    if (sessionOrgId && String(item.ORG_ID || "").trim().toUpperCase() === sessionOrgId) return true;
    if (sessionClientId && String(item.CLIENT_ID || "").trim().toUpperCase() === sessionClientId) return true;
    if (String(item.REQUESTED_BY_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    if (String(item.ASSIGNED_TO_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase()) return true;
    return false;
  });
  if (filters.status) requests = requests.filter(function(item) { return String(item.STATUS || "") === String(filters.status); });
  if (filters.caseId) requests = requests.filter(function(item) { return String(item.CASE_ID || "") === String(filters.caseId); });
  requests.sort(function(a, b) { return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0); });
  return writePortalCache_(cacheKey, sanitizeDataForFrontend_(requests), 60);
}

function saveDocumentRequest_(session, data) {
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  return withScriptLock_(function() {
    var requestId = data.REQUEST_ID || generateId_("DCR", getSheet_("DOCUMENT_REQUESTS"), 0);
    var payload = {
      TITLE: data.title || "",
      DESCRIPTION: data.description || "",
      REQUEST_TYPE: data.requestType || "Document",
      STATUS: data.status || "Open",
      CLIENT_ID: data.clientId || "",
      ORG_ID: data.orgId || "",
      CASE_ID: data.caseId || "",
      REQUESTED_BY_EMAIL: session.email,
      ASSIGNED_TO_EMAIL: data.assignedToEmail || "",
      DUE_DATE: data.dueDate || "",
      DRIVE_LINK: data.driveLink || "",
      CLIENT_VISIBLE: data.clientVisible || "Yes",
      APPROVAL_STATUS: data.approvalStatus || "Not Required",
      UPDATED_AT: new Date()
    };
    if (data.REQUEST_ID) {
      var ok = updateRecordById_("DOCUMENT_REQUESTS", "REQUEST_ID", requestId, payload);
      if (!ok) return { error: "Document request not found." };
    } else {
      payload.REQUEST_ID = requestId;
      payload.CREATED_AT = new Date();
      appendRecord_("DOCUMENT_REQUESTS", payload);
    }
    if (payload.ASSIGNED_TO_EMAIL) {
      createNotification_(payload.ASSIGNED_TO_EMAIL, "Document request assigned", payload.TITLE || requestId, "DOCUMENT_REQUEST", requestId);
    }
    createTimelineEvent_({
      EVENT_TYPE: data.REQUEST_ID ? "DOCUMENT_REQUEST_UPDATED" : "DOCUMENT_REQUEST_CREATED",
      TITLE: payload.TITLE || "Document request",
      DESCRIPTION: "Document workflow updated by " + session.name,
      ENTITY_TYPE: "DOCUMENT_REQUEST",
      ENTITY_ID: requestId,
      CLIENT_ID: payload.CLIENT_ID || "",
      ORG_ID: payload.ORG_ID || "",
      CASE_ID: payload.CASE_ID || "",
      USER_EMAIL: session.email,
      USER_NAME: session.name,
      VISIBILITY: payload.CLIENT_VISIBLE === "Yes" ? "Shared" : "Internal"
    });
    clearPortalCaches_(["documentRequests", "timeline", "notifications"]);
    return { success: true, requestId: requestId };
  });
}

function reviewDocumentRequest_(session, requestId, reviewData) {
  var request = getRecords_("DOCUMENT_REQUESTS").find(function(item) { return item.REQUEST_ID === requestId; });
  if (!request) return { error: "Document request not found." };
  if (!hasRoleAtLeast_(getEffectiveRoles_(session), "Staff")) return { error: "Access denied." };
  var ok = updateRecordById_("DOCUMENT_REQUESTS", "REQUEST_ID", requestId, {
    STATUS: reviewData.status || request.STATUS || "Open",
    DRIVE_LINK: reviewData.driveLink || request.DRIVE_LINK || "",
    APPROVAL_STATUS: reviewData.approvalStatus || request.APPROVAL_STATUS || "",
    UPDATED_AT: new Date()
  });
  if (!ok) return { error: "Document request update failed." };
  if (request.REQUESTED_BY_EMAIL) {
    createNotification_(request.REQUESTED_BY_EMAIL, "Document request updated", request.TITLE + " is now " + (reviewData.status || request.STATUS), "DOCUMENT_REQUEST", requestId);
  }
  clearPortalCaches_(["documentRequests", "timeline", "notifications"]);
  return { success: true };
}

function getGalvanizerCommandCenter_(session, filters) {
  if (!hasAnyRole_(session, ["Super Admin", "Admin", "Galvanizer"])) return { error: "Access denied." };
  filters = filters || {};
  var queue = getGalvanizerQueue_(session, filters);
  var tasks = getTasks_(session, {});
  var approvals = getApprovalRequests_(session, { status: "Pending" });
  return sanitizeDataForFrontend_({
    incoming: (queue || []).filter(function(item) { return String(item.WORKFLOW_STAGE || "") === "Drafting"; }),
    readyForAttorney: (queue || []).filter(function(item) { return String(item.WORKFLOW_STAGE || "") === "Ready for Attorney"; }),
    underReview: (queue || []).filter(function(item) { return String(item.WORKFLOW_STAGE || "") === "Under Attorney Review"; }),
    tasks: tasks && !tasks.error ? tasks.filter(function(item) {
      return String(item.ASSIGNED_TO_EMAIL || "").toLowerCase() === String(session.email || "").toLowerCase();
    }) : [],
    pendingApprovals: approvals && !approvals.error ? approvals : []
  });
}
