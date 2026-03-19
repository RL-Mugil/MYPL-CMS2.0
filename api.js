/**
 * api.js - API client for the portal frontend.
 *
 * This frontend should talk to the Cloudflare Worker only.
 * The Worker can then verify Clerk and forward trusted calls to GAS.
 */

const API_BASE = (window.APP_CONFIG && window.APP_CONFIG.apiBase) || '';

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

async function clerkLogin(email) {
  return callGASPublic('clerklogin', { email });
}

async function getDashboard(filters = {}) { return callGAS('getDashboard', { filters }); }
async function getCases(filters = {}) { return callGAS('getCases', { filters }); }
async function getInvoices() { return callGAS('getInvoices', {}); }
async function getDocuments() { return callGAS('getDocuments', {}); }
async function getOrganizations() { return callGAS('getOrganizations', {}); }
async function saveOrganization(orgData) { return callGAS('saveOrganization', { orgData }); }
async function getOrganizationUsers(orgId) { return callGAS('getOrganizationUsers', { orgId }); }
async function getCircles() { return callGAS('getCircles', {}); }
async function getCircleMembers(circleId) { return callGAS('getCircleMembers', { circleId }); }
async function saveCircle(circleData) { return callGAS('saveCircle', { circleData }); }
async function deleteCircle(circleId) { return callGAS('deleteCircle', { circleId }); }
async function saveCircleMember(memberData) { return callGAS('saveCircleMember', { memberData }); }
async function removeCircleMember(membershipId) { return callGAS('removeCircleMember', { membershipId }); }
async function saveDailyPriority(priorityData) { return callGAS('saveDailyPriority', { priorityData }); }
async function saveDailyWrapup(wrapupData) { return callGAS('saveDailyWrapup', { wrapupData }); }
async function getDailyOpsOverview() { return callGAS('getDailyOpsOverview', {}); }
async function getDailyAudit(filters = {}) { return callGAS('getDailyAudit', { filters }); }
async function submitExpenseClaim(claimData) { return callGAS('submitExpenseClaim', { claimData }); }
async function getExpenseClaims(filters = {}) { return callGAS('getExpenseClaims', { filters }); }
async function reviewExpenseClaim(claimId, reviewData) { return callGAS('reviewExpenseClaim', { claimId, reviewData }); }
async function getNotifications() { return callGAS('getNotifications', {}); }
async function markNotificationRead(notificationId) { return callGAS('markNotificationRead', { notificationId }); }
async function getMessageThreads() { return callGAS('getMessageThreads', {}); }
async function getGalvanizerQueue(filters = {}) { return callGAS('getGalvanizerQueue', { filters }); }
async function getThreadMessages(threadId) { return callGAS('getThreadMessages', { threadId }); }
async function saveMessageThread(threadData) { return callGAS('saveMessageThread', { threadData }); }
async function deleteMessageThread(threadId) { return callGAS('deleteMessageThread', { threadId }); }
async function sendThreadMessage(messageData) { return callGAS('sendThreadMessage', { messageData }); }

async function submitContact(subject, caseId, message) {
  return callGAS('submitContact', { subject, caseId, message });
}

async function markInvoicePaid(invoiceId) {
  return callGAS('markInvoicePaid', { invoiceId });
}

async function getUsers(filters = {}) { return callGAS('getUsers', { filters }); }
async function saveUser(userData) { return callGAS('saveUser', { userData }); }
async function deleteUser(userId) { return callGAS('deleteUser', { userId }); }

async function getClients() { return callGAS('getClients', {}); }
async function saveClient(clientData) { return callGAS('saveClient', { clientData }); }
async function deleteClient(clientId) { return callGAS('deleteClient', { clientId }); }

async function saveCase(caseData) { return callGAS('saveCase', { caseData }); }
async function bulkUpdateCases(bulkData) { return callGAS('bulkUpdateCases', { bulkData }); }
async function deleteCase(caseId) { return callGAS('deleteCase', { caseId }); }

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

window.API = {
  callGAS,
  callGASPublic,
  clerkLogin,
  getDashboard,
  getCases,
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
  getExpenseClaims,
  reviewExpenseClaim,
  getNotifications,
  markNotificationRead,
  getMessageThreads,
  getGalvanizerQueue,
  getThreadMessages,
  saveMessageThread,
  deleteMessageThread,
  sendThreadMessage,
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
