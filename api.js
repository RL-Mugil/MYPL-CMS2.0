/**
 * api.js - API client for the portal frontend.
 *
 * This frontend should talk to the Cloudflare Worker only.
 * The Worker can then verify Clerk and forward trusted calls to GAS.
 */

const API_BASE = (window.APP_CONFIG && window.APP_CONFIG.apiBase) || '';
const API_CACHE_PREFIX = 'mg_api_cache_v1';
const API_MEMORY_CACHE = new Map();
const SAFE_LOOKUP_TTL_MS = 5 * 24 * 60 * 60 * 1000;

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sanitizePayload(value) {
  if (value == null) return value;
  if (Array.isArray(value)) return value.map(sanitizePayload);
  if (typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach((key) => {
      out[key] = sanitizePayload(value[key]);
    });
    return out;
  }
  return typeof value === 'string' ? escapeHtml(value) : value;
}

function ensureApiBase() {
  if (!API_BASE) {
    throw new Error('API endpoint is not configured.');
  }
}

function stableStringify(value) {
  if (value == null) return '';
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
  if (typeof value === 'object') {
    return `{${Object.keys(value).sort().map((key) => `${key}:${stableStringify(value[key])}`).join(',')}}`;
  }
  return JSON.stringify(value);
}

function getCacheScope() {
  const session = window.Auth?.getGasSession?.();
  return session?.email || session?.token || 'public';
}

function buildClientCacheKey(action, params = {}) {
  return `${API_CACHE_PREFIX}:${getCacheScope()}:${action}:${stableStringify(params)}`;
}

function readClientCache(key) {
  const memory = API_MEMORY_CACHE.get(key);
  if (memory && memory.expiresAt > Date.now()) {
    return memory.value;
  }
  if (memory) {
    API_MEMORY_CACHE.delete(key);
  }

  try {
    const raw = sessionStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || parsed.expiresAt <= Date.now()) {
      sessionStorage.removeItem(key);
      return null;
    }
    API_MEMORY_CACHE.set(key, parsed);
    return parsed.value;
  } catch {
    return null;
  }
}

function writeClientCache(key, value, ttlMs) {
  const record = { value, expiresAt: Date.now() + ttlMs };
  API_MEMORY_CACHE.set(key, record);
  try {
    sessionStorage.setItem(key, JSON.stringify(record));
  } catch {}
  return value;
}

function readPersistentLookupCache(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || parsed.expiresAt <= Date.now()) {
      localStorage.removeItem(key);
      return null;
    }
    return parsed.value;
  } catch {
    return null;
  }
}

function writePersistentLookupCache(key, value, ttlMs = SAFE_LOOKUP_TTL_MS) {
  try {
    localStorage.setItem(key, JSON.stringify({
      value,
      expiresAt: Date.now() + ttlMs,
    }));
  } catch {}
  return value;
}

function peekCachedValue(action, params = {}, allowPersistent = false) {
  const key = buildClientCacheKey(action, params);
  const sessionCached = readClientCache(key);
  if (sessionCached) return sessionCached;
  if (allowPersistent) return readPersistentLookupCache(key);
  return null;
}

function clearClientCache(matchers = []) {
  const shouldClear = (key) => !matchers.length || matchers.some((matcher) => key.includes(`:${matcher}:`));

  Array.from(API_MEMORY_CACHE.keys()).forEach((key) => {
    if (key.startsWith(API_CACHE_PREFIX) && shouldClear(key)) {
      API_MEMORY_CACHE.delete(key);
    }
  });

  try {
    Object.keys(sessionStorage).forEach((key) => {
      if (key.startsWith(API_CACHE_PREFIX) && shouldClear(key)) {
        sessionStorage.removeItem(key);
      }
    });
  } catch {}

  try {
    Object.keys(localStorage).forEach((key) => {
      if (key.startsWith(API_CACHE_PREFIX) && shouldClear(key)) {
        localStorage.removeItem(key);
      }
    });
  } catch {}
}

function reportError(context, error, extra = {}) {
  const payload = {
    context,
    message: error && error.message ? error.message : String(error),
    stack: error && error.stack ? error.stack : '',
    extra,
    path: window.location.pathname,
    env: window.APP_CONFIG && window.APP_CONFIG.environment,
    timestamp: new Date().toISOString(),
  };

  console.error('[portal-error]', payload);

  if (window.APP_CONFIG && window.APP_CONFIG.errorEndpoint) {
    fetch(window.APP_CONFIG.errorEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      keepalive: true,
    }).catch(() => {});
  }
}

async function requestJson(payload, context) {
  ensureApiBase();

  try {
    const res = await fetch(API_BASE, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      redirect: 'follow',
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return sanitizePayload(await res.json());
  } catch (error) {
    reportError(context, error, { action: payload.action });
    throw error;
  }
}

async function callGAS(action, params = {}) {
  const session = window.Auth.getGasSession();
  const token = session?.token || null;
  const data = await requestJson({ action, params: { ...params, token } }, `callGAS:${action}`);

  if (data.sessionExpired) {
    window.Auth.clearGasSession();
    showGlobalError('Session expired. Please sign in again.');
    setTimeout(() => location.reload(), 2000);
    throw new Error('session_expired');
  }

  if (data.error) {
    throw new Error(data.error);
  }

  return data;
}

async function callGASPublic(action, params = {}) {
  return requestJson({ action, params }, `callGASPublic:${action}`);
}

async function callGASCached(action, params = {}, ttlMs = 60000) {
  const key = buildClientCacheKey(action, params);
  const cached = readClientCache(key);
  if (cached) return cached;
  const fresh = await callGAS(action, params);
  return writeClientCache(key, fresh, ttlMs);
}

async function callGASLookupCached(action, params = {}, ttlMs = SAFE_LOOKUP_TTL_MS) {
  const key = buildClientCacheKey(action, params);
  const sessionCached = readClientCache(key);
  if (sessionCached) return sessionCached;

  const persistentCached = readPersistentLookupCache(key);
  if (persistentCached) {
    writeClientCache(key, persistentCached, Math.min(ttlMs, 60 * 60 * 1000));
    return persistentCached;
  }

  const fresh = await callGAS(action, params);
  writeClientCache(key, fresh, Math.min(ttlMs, 60 * 60 * 1000));
  return writePersistentLookupCache(key, fresh, ttlMs);
}

async function callGASPersistentCached(action, params = {}, ttlMs = 15 * 60 * 1000) {
  const key = buildClientCacheKey(action, params);
  const sessionCached = readClientCache(key);
  if (sessionCached) return sessionCached;

  const persistentCached = readPersistentLookupCache(key);
  if (persistentCached) {
    writeClientCache(key, persistentCached, Math.min(ttlMs, 60 * 1000));
    return persistentCached;
  }

  const fresh = await callGAS(action, params);
  writeClientCache(key, fresh, Math.min(ttlMs, 60 * 1000));
  return writePersistentLookupCache(key, fresh, ttlMs);
}

async function clerkLogin(email) {
  return callGASPublic('clerklogin', { email });
}

async function getDashboard(filters = {}) { return callGASCached('getDashboard', { filters }, 30000); }
async function getDashboardSummary(filters = {}) { return callGASPersistentCached('getDashboardSummary', { filters }, 15 * 60 * 1000); }
async function getDashboardDetails(filters = {}) { return callGASPersistentCached('getDashboardDetails', { filters }, 15 * 60 * 1000); }
async function getCases(filters = {}) { return callGASCached('getCases', { filters }, 30000); }
async function getCasesPage(filters = {}, limit = 50, offset = 0) {
  return callGASCached('getCasesPage', { filters, limit, offset }, 30000);
}
async function getInvoices() { return callGAS('getInvoices', {}); }
async function getDocuments() { return callGAS('getDocuments', {}); }
async function getOrganizations() { return callGASLookupCached('getOrganizations', {}, SAFE_LOOKUP_TTL_MS); }
async function saveOrganization(orgData) {
  const result = await callGAS('saveOrganization', { orgData });
  clearClientCache(['getOrganizations', 'getClients', 'getUsers', 'getDashboard', 'getCases']);
  return result;
}
async function getOrganizationUsers(orgId) { return callGASLookupCached('getOrganizationUsers', { orgId }, SAFE_LOOKUP_TTL_MS); }
async function getCircles() { return callGASLookupCached('getCircles', {}, SAFE_LOOKUP_TTL_MS); }
async function getCircleMembers(circleId) { return callGASLookupCached('getCircleMembers', { circleId }, SAFE_LOOKUP_TTL_MS); }
async function saveCircle(circleData) {
  const result = await callGAS('saveCircle', { circleData });
  clearClientCache(['getCircles', 'getCircleMembers', 'getUsers']);
  return result;
}
async function deleteCircle(circleId) {
  const result = await callGAS('deleteCircle', { circleId });
  clearClientCache(['getCircles', 'getCircleMembers', 'getUsers']);
  return result;
}
async function saveCircleMember(memberData) {
  const result = await callGAS('saveCircleMember', { memberData });
  clearClientCache(['getCircles', 'getCircleMembers', 'getUsers']);
  return result;
}
async function removeCircleMember(membershipId) {
  const result = await callGAS('removeCircleMember', { membershipId });
  clearClientCache(['getCircles', 'getCircleMembers', 'getUsers']);
  return result;
}
async function saveDailyPriority(priorityData) {
  const result = await callGAS('saveDailyPriority', { priorityData });
  clearClientCache(['getDailyOpsOverview', 'getDailyAudit', 'getDashboard', 'getNotifications']);
  return result;
}
async function saveDailyWrapup(wrapupData) {
  const result = await callGAS('saveDailyWrapup', { wrapupData });
  clearClientCache(['getDailyOpsOverview', 'getDailyAudit', 'getDashboard', 'getNotifications']);
  return result;
}
async function getDailyOpsOverview() { return callGASCached('getDailyOpsOverview', {}, 60000); }
async function getDailyAudit(filters = {}) { return callGASCached('getDailyAudit', { filters }, 60000); }
async function submitExpenseClaim(claimData) {
  const result = await callGAS('submitExpenseClaim', { claimData });
  clearClientCache(['getExpenseClaims', 'getDailyOpsOverview', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getNotifications']);
  return result;
}
async function uploadExpenseBill(fileData) { return callGAS('uploadExpenseBill', { fileData }); }
async function getExpenseClaims(filters = {}) { return callGASCached('getExpenseClaims', { filters }, 60000); }
async function reviewExpenseClaim(claimId, reviewData) {
  const result = await callGAS('reviewExpenseClaim', { claimId, reviewData });
  clearClientCache(['getExpenseClaims', 'getDailyOpsOverview', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getNotifications']);
  return result;
}
async function getNotifications() { return callGASCached('getNotifications', {}, 30000); }
async function markNotificationRead(notificationId) {
  const result = await callGAS('markNotificationRead', { notificationId });
  clearClientCache(['getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function deleteNotification(notificationId) {
  const result = await callGAS('deleteNotification', { notificationId });
  clearClientCache(['getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function clearNotifications() {
  const result = await callGAS('clearNotifications', {});
  clearClientCache(['getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function getDirectInbox() { return callGASCached('getDirectInbox', {}, 30000); }
async function getMessageThreads() { return callGASCached('getMessageThreads', {}, 30000); }
async function getGalvanizerQueue(filters = {}) { return callGASCached('getGalvanizerQueue', { filters }, 30000); }
async function getThreadMessages(threadId) { return callGAS('getThreadMessages', { threadId }); }
async function markThreadRead(threadId) {
  const result = await callGAS('markThreadRead', { threadId });
  clearClientCache(['getDirectInbox', 'getMessageThreads', 'getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function saveMessageThread(threadData) {
  const result = await callGAS('saveMessageThread', { threadData });
  clearClientCache(['getDirectInbox', 'getMessageThreads', 'getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function deleteMessageThread(threadId) {
  const result = await callGAS('deleteMessageThread', { threadId });
  clearClientCache(['getDirectInbox', 'getMessageThreads', 'getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function sendThreadMessage(messageData) {
  const result = await callGAS('sendThreadMessage', { messageData });
  clearClientCache(['getDirectInbox', 'getMessageThreads', 'getNotifications', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails']);
  return result;
}
async function createDirectThread(threadData) {
  const result = await callGAS('createDirectThread', { threadData });
  clearClientCache(['getDirectInbox', 'getMessageThreads', 'getNotifications']);
  return result;
}
async function getTasks(filters = {}) { return callGASCached('getTasks', { filters }, 60000); }
async function saveTask(taskData) {
  const result = await callGAS('saveTask', { taskData });
  clearClientCache(['getTasks', 'getActivityTimeline', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getNotifications']);
  return result;
}
async function updateTaskStatus(taskId, status, notes = '') {
  const result = await callGAS('updateTaskStatus', { taskId, status, notes });
  clearClientCache(['getTasks', 'getActivityTimeline', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getNotifications']);
  return result;
}
async function getActivityTimeline(filters = {}) { return callGASCached('getActivityTimeline', { filters }, 60000); }
async function getWorkflowBoard(filters = {}) { return callGASCached('getWorkflowBoard', { filters }, 60000); }
async function getAttorneyWorkspace(filters = {}) { return callGASCached('getAttorneyWorkspace', { filters }, 60000); }
async function getSmartSearch(query, scope = 'all') { return callGASCached('getSmartSearch', { query, scope }, 60000); }
async function getApprovalRequests(filters = {}) { return callGASCached('getApprovalRequests', { filters }, 60000); }
async function saveApprovalRequest(approvalData) {
  const result = await callGAS('saveApprovalRequest', { approvalData });
  clearClientCache(['getApprovalRequests', 'getActivityTimeline', 'getNotifications']);
  return result;
}
async function reviewApprovalRequest(approvalId, reviewData) {
  const result = await callGAS('reviewApprovalRequest', { approvalId, reviewData });
  clearClientCache(['getApprovalRequests', 'getActivityTimeline', 'getNotifications']);
  return result;
}
async function getDocumentRequests(filters = {}) { return callGASCached('getDocumentRequests', { filters }, 60000); }
async function saveDocumentRequest(requestData) {
  const result = await callGAS('saveDocumentRequest', { requestData });
  clearClientCache(['getDocumentRequests', 'getActivityTimeline', 'getNotifications']);
  return result;
}
async function uploadPortalDocument(documentData) {
  const result = await callGAS('uploadPortalDocument', { documentData });
  clearClientCache(['getDocuments', 'getDocumentRequests', 'getActivityTimeline', 'getNotifications']);
  return result;
}
async function reviewDocumentRequest(requestId, reviewData) {
  const result = await callGAS('reviewDocumentRequest', { requestId, reviewData });
  clearClientCache(['getDocumentRequests', 'getActivityTimeline', 'getNotifications']);
  return result;
}
async function getGalvanizerCommandCenter(filters = {}) { return callGASCached('getGalvanizerCommandCenter', { filters }, 60000); }

async function submitContact(subject, caseId, message) {
  return callGAS('submitContact', { subject, caseId, message });
}

async function markInvoicePaid(invoiceId) {
  return callGAS('markInvoicePaid', { invoiceId });
}

async function getUsers(filters = {}) {
  const hasFilters = filters && Object.keys(filters).some((key) => String(filters[key] || '').trim() !== '');
  return hasFilters
    ? callGASCached('getUsers', { filters }, 300000)
    : callGASLookupCached('getUsers', { filters }, SAFE_LOOKUP_TTL_MS);
}
async function saveUser(userData) {
  const result = await callGAS('saveUser', { userData });
  clearClientCache(['getUsers', 'getOrganizations', 'getCircles', 'getCircleMembers', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getCases', 'getCasesPage']);
  return result;
}
async function deleteUser(userId) {
  const result = await callGAS('deleteUser', { userId });
  clearClientCache(['getUsers', 'getOrganizations', 'getCircles', 'getCircleMembers', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getCases', 'getCasesPage']);
  return result;
}

async function getClients() { return callGASLookupCached('getClients', {}, SAFE_LOOKUP_TTL_MS); }
async function saveClient(clientData) {
  const result = await callGAS('saveClient', { clientData });
  clearClientCache(['getClients', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getCases', 'getCasesPage', 'getOrganizations']);
  return result;
}
async function deleteClient(clientId) {
  const result = await callGAS('deleteClient', { clientId });
  clearClientCache(['getClients', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getCases', 'getCasesPage', 'getOrganizations']);
  return result;
}

async function saveCase(caseData) {
  const result = await callGAS('saveCase', { caseData });
  clearClientCache(['getCases', 'getCasesPage', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getGalvanizerQueue', 'getDocuments']);
  return result;
}
async function bulkUpdateCases(bulkData) {
  const result = await callGAS('bulkUpdateCases', { bulkData });
  clearClientCache(['getCases', 'getCasesPage', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getGalvanizerQueue', 'getDocuments']);
  return result;
}
async function deleteCase(caseId) {
  const result = await callGAS('deleteCase', { caseId });
  clearClientCache(['getCases', 'getCasesPage', 'getDashboard', 'getDashboardSummary', 'getDashboardDetails', 'getGalvanizerQueue', 'getDocuments']);
  return result;
}

async function saveInvoice(invoiceData) { return callGAS('saveInvoice', { invoiceData }); }
async function deleteInvoice(invoiceId) { return callGAS('deleteInvoice', { invoiceId }); }

function showGlobalError(msg) {
  let el = document.getElementById('global-error');
  if (!el) {
    el = document.createElement('div');
    el.id = 'global-error';
    el.style.cssText = 'position:fixed;top:20px;right:20px;z-index:9999;max-width:360px;';
    document.body.appendChild(el);
  }

  el.textContent = msg;
  el.className = 'alert alert-error';

  setTimeout(() => {
    if (el) {
      el.textContent = '';
      el.className = '';
    }
  }, 5000);
}

function getCachedDashboardSummary(filters = {}) {
  return peekCachedValue('getDashboardSummary', { filters }, true);
}

function getCachedDashboardDetails(filters = {}) {
  return peekCachedValue('getDashboardDetails', { filters }, true);
}

function getCachedUsers(filters = {}) {
  const hasFilters = filters && Object.keys(filters).some((key) => String(filters[key] || '').trim() !== '');
  return peekCachedValue('getUsers', { filters }, !hasFilters);
}

function getCachedClients() {
  return peekCachedValue('getClients', {}, true);
}

function getCachedNotifications() {
  return peekCachedValue('getNotifications', {}, false);
}

function getCachedMessageThreads() {
  return peekCachedValue('getMessageThreads', {}, false);
}

function getCachedDirectInbox() {
  return peekCachedValue('getDirectInbox', {}, false);
}

window.API = {
  callGAS,
  callGASPublic,
  clerkLogin,
  getDashboard,
  getDashboardSummary,
  getDashboardDetails,
  getCachedDashboardSummary,
  getCachedDashboardDetails,
  getCachedUsers,
  getCachedClients,
  getCachedNotifications,
  getCachedMessageThreads,
  getCachedDirectInbox,
  getCases,
  getCasesPage,
  getInvoices,
  getDocuments,
  getOrganizations,
  saveOrganization,
  getOrganizationUsers,
  getCircles,
  getCircleMembers,
  saveCircle,
  deleteCircle,
  saveCircleMember,
  removeCircleMember,
  saveDailyPriority,
  saveDailyWrapup,
  getDailyOpsOverview,
  getDailyAudit,
  submitExpenseClaim,
  uploadExpenseBill,
  getExpenseClaims,
  reviewExpenseClaim,
  getNotifications,
  markNotificationRead,
  deleteNotification,
  clearNotifications,
  getDirectInbox,
  getMessageThreads,
  getGalvanizerQueue,
  getThreadMessages,
  markThreadRead,
  saveMessageThread,
  createDirectThread,
  deleteMessageThread,
  sendThreadMessage,
  getTasks,
  saveTask,
  updateTaskStatus,
  getActivityTimeline,
  getWorkflowBoard,
  getAttorneyWorkspace,
  getSmartSearch,
  getApprovalRequests,
  saveApprovalRequest,
  reviewApprovalRequest,
  getDocumentRequests,
  saveDocumentRequest,
  uploadPortalDocument,
  reviewDocumentRequest,
  getGalvanizerCommandCenter,
  submitContact,
  markInvoicePaid,
  getUsers, saveUser, deleteUser,
  getClients, saveClient, deleteClient,
  saveCase, deleteCase,
  bulkUpdateCases,
  saveInvoice, deleteInvoice,
  showGlobalError,
  reportError,
};

window.Safe = {
  escapeHtml,
  sanitizePayload,
};

window.addEventListener('error', (event) => {
  reportError('window.error', event.error || new Error(event.message), {
    filename: event.filename,
    lineno: event.lineno,
    colno: event.colno,
  });
});

window.addEventListener('unhandledrejection', (event) => {
  const reason = event.reason instanceof Error ? event.reason : new Error(String(event.reason));
  reportError('window.unhandledrejection', reason);
});
