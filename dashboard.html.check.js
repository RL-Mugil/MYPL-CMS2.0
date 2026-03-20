// ── Globals ──────────────────────────────────────────────────────────────────
let SESSION = null;
let _charts = {};
let DASHBOARD_FILTERS = { clientCode: '', caseId: '', fromDate: '', toDate: '' };
let CASE_FILTERS = { clientCode: '', caseId: '', status: '', country: '', workflowStage: '', query: '' };
let USER_FILTERS = { query: '', role: '' };
let CASES_PAGE_LIMIT = 50;
let CASES_PAGE_OFFSET = 0;
let CASES_HAS_MORE = false;
let CASES_TOTAL = 0;
let CASES_ITEMS = [];
let USERS_PAGE_LIMIT = 50;
let USERS_VISIBLE_COUNT = 50;
let CLIENTS_PAGE_LIMIT = 50;
let CLIENTS_VISIBLE_COUNT = 50;
let MESSAGES_PAGE_LIMIT = 25;
let MESSAGES_VISIBLE_COUNT = 25;
let NOTICES_PAGE_LIMIT = 25;
let NOTICES_VISIBLE_COUNT = 25;
let SAFE_PREFETCH_STARTED = false;
let CURRENT_THEME = 'dark';
let CURRENT_PAGE = 'dashboard';
let CASE_IMPORT_STATE = { parsedRows: [], importableRows: [], skippedRows: [], fileName: '' };
const NAV_SECTION_STORAGE_KEY = 'mg_sidebar_sections_v1';
const SIDEBAR_COLLAPSED_STORAGE_KEY = 'mg_sidebar_collapsed_v1';

function hasRoleAtLeast(role, minimumRole) {
  const order = {
    'Super Admin': 60,
    'Admin': 50,
    'Galvanizer': 45,
    'Staff': 40,
    'Attorney': 30,
    'Client Admin': 20,
    'Client Employee': 15,
    'Individual Client': 10,
  };
  if (Array.isArray(role)) {
    return role.some((item) => (order[item] || 0) >= (order[minimumRole] || 0));
  }
  return (order[role] || 0) >= (order[minimumRole] || 0);
}

function getEffectiveSessionRoles() {
  const roles = [];
  if (SESSION?.role) roles.push(SESSION.role);
  String(SESSION?.additionalRoles || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean)
    .forEach((role) => {
      if (!roles.includes(role)) roles.push(role);
    });
  return roles;
}

function sessionHasRoleAtLeast(minimumRole) {
  return hasRoleAtLeast(getEffectiveSessionRoles(), minimumRole);
}

function sessionHasAnyRole(roles) {
  const effectiveRoles = getEffectiveSessionRoles();
  return roles.some((role) => effectiveRoles.includes(role));
}

function canAccessFinance() {
  return !!SESSION?.canViewFinance || sessionHasRoleAtLeast('Admin');
}

function canAccessManagement() {
  return sessionHasAnyRole(['Super Admin', 'Admin', 'Client Admin']);
}

function isInternalUser() {
  return sessionHasAnyRole(['Super Admin', 'Admin', 'Galvanizer', 'Staff', 'Attorney']);
}

// ── Init ─────────────────────────────────────────────────────────────────────
(async () => {
  applyTheme(getStoredTheme());
  SESSION = await Auth.requireAuth();
  if (!SESSION) {
    window.location.href = 'index.html';
    return;
  }

  if (!canAccessManagement()) {
    document.getElementById('nav-mgmt-section').style.display = 'none';
  }
  if (!canAccessFinance()) {
    document.getElementById('nav-finance').style.display = 'none';
  }
  if (!isInternalUser()) {
    ['messages','notifications','dailyaudit','expenses','galvanizer','contact'].forEach((page) => {
      document.querySelector(`[data-page="${page}"]`)?.style.setProperty('display', 'none');
    });
    document.getElementById('circle-nav-section').style.display = 'none';
  }

  // Populate sidebar user info
  const initials = (SESSION.name || 'U').split(' ').map(n => n[0]).join('').slice(0,2).toUpperCase();
  document.getElementById('sidebar-avatar').textContent = initials;
  document.getElementById('sidebar-name').textContent = SESSION.name || SESSION.email;
  document.getElementById('sidebar-role').textContent = SESSION.role;
  applySidebarState();
  applyNavSectionState();
  await populateCircleNav();
  startSafePrefetch();

  showPage('dashboard');
})();

function toggleSidebarCollapsed() {
  const next = !document.body.classList.contains('sidebar-hidden');
  document.body.classList.toggle('sidebar-hidden', next);
  try { localStorage.setItem(SIDEBAR_COLLAPSED_STORAGE_KEY, next ? '1' : '0'); } catch {}
}

function applySidebarState() {
  try {
    const hidden = localStorage.getItem(SIDEBAR_COLLAPSED_STORAGE_KEY) === '1';
    document.body.classList.toggle('sidebar-hidden', hidden);
  } catch {}
}

function readNavSectionState() {
  try {
    return JSON.parse(localStorage.getItem(NAV_SECTION_STORAGE_KEY) || '{}');
  } catch {
    return {};
  }
}

function writeNavSectionState(state) {
  try { localStorage.setItem(NAV_SECTION_STORAGE_KEY, JSON.stringify(state)); } catch {}
}

function toggleNavSection(sectionName) {
  const state = readNavSectionState();
  const current = state[sectionName] !== false;
  state[sectionName] = !current;
  writeNavSectionState(state);
  applyNavSectionState();
}

function applyNavSectionState() {
  const state = readNavSectionState();
  document.querySelectorAll('[data-nav-section]').forEach((section) => {
    const name = section.getAttribute('data-nav-section');
    const expanded = state[name] !== false;
    section.classList.toggle('collapsed', !expanded);
  });
}

async function populateCircleNav() {
  const wrap = document.getElementById('circle-nav-items');
  try {
    const circles = await API.getCircles();
    if (!Array.isArray(circles) || !circles.length) {
      document.getElementById('circle-nav-section').style.display = 'none';
      return;
    }
    wrap.innerHTML = circles.map(circle => `
      <button class="nav-item" data-page="circles" onclick='loadCircles(${JSON.stringify(JSON.stringify(circle))})'>
        <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="8" cy="8" r="5"/><path d="M8 5v6M5 8h6"/></svg>
        ${circle.CIRCLE_NAME}
      </button>`).join('');
  } catch (e) {
    document.getElementById('circle-nav-section').style.display = 'none';
  }
}

// ── Navigation ───────────────────────────────────────────────────────────────
const PAGE_TITLES = {
  dashboard: 'Dashboard',
  cases: 'Cases',
  documents: 'Documents',
  inbox: 'Inbox',
  workflow: 'Workflow Board',
  attorney: 'Attorney Workspace',
  docworkflow: 'Document Workflow',
  finance: 'Finance',
  messages: 'Threads',
  notifications: 'Notification Center',
  tasks: 'Task System',
  timeline: 'Activity Timeline',
  search: 'Smart Search',
  approvals: 'Approval Workflow',
  dailyaudit: 'Daily Audit',
  expenses: 'Expense Claims',
  galvanizer: 'Galvanizer Queue',
  circles: 'Circles',
  management: 'Management',
  contact: 'Contact',
};

function showPage(page) {
  CURRENT_PAGE = page;
  document.querySelectorAll('.nav-item').forEach(el => {
    el.classList.toggle('active', el.dataset.page === page);
  });
  document.getElementById('topbar-title').textContent = PAGE_TITLES[page] || page;
  document.querySelector('.sidebar').classList.remove('open');

  const content = document.getElementById('page-content');
  content.innerHTML = '<div class="loading-wrap"><div class="spinner"></div><span>Loading…</span></div>';

  setTopbarActions('');

  const loaders = {
    dashboard: loadDashboard,
    cases: loadCases,
    documents: loadDocuments,
    inbox: loadInbox,
    workflow: loadWorkflowBoard,
    attorney: loadAttorneyWorkspace,
    docworkflow: loadDocumentWorkflow,
    finance: loadFinance,
    messages: loadMessages,
    notifications: loadNotifications,
    tasks: loadTasks,
    timeline: loadTimeline,
    search: loadSmartSearch,
    approvals: loadApprovals,
    dailyaudit: loadDailyAudit,
    expenses: loadExpenses,
    galvanizer: loadGalvanizerQueue,
    circles: loadCircles,
    management: loadManagement,
    contact: loadContact,
  };
  if (loaders[page]) loaders[page]();
}

// ── Sign out ──────────────────────────────────────────────────────────────────
async function handleSignOut() {
  await Auth.signOut();
  window.location.href = 'index.html';
}

// ── Modal helpers ─────────────────────────────────────────────────────────────
function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }
function openGenericModal(title, bodyHtml) {
  document.getElementById('modal-generic-title').textContent = title || 'Dialog';
  document.getElementById('modal-generic-body').innerHTML = bodyHtml || '';
  openModal('modal-generic');
}
document.querySelectorAll('.modal-overlay').forEach(m => {
  m.addEventListener('click', e => { if (e.target === m) m.classList.remove('open'); });
});

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmt(v) {
  if (!v) return '—';
  try { return new Date(v).toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' }); }
  catch { return v; }
}
function toInputDate(v) {
  if (!v) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(v))) return String(v);
  const parsed = new Date(v);
  if (!isNaN(parsed.getTime())) return parsed.toISOString().split('T')[0];
  const match = String(v).match(/^(\d{2})-([A-Za-z]{3})-(\d{4})$/);
  if (!match) return '';
  const months = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  return `${match[3]}-${months[match[2]] || '01'}-${match[1]}`;
}
function money(v) {
  const n = parseFloat(v) || 0;
  return '₹' + n.toLocaleString('en-IN', { minimumFractionDigits:2, maximumFractionDigits:2 });
}
function statusBadge(s) {
  const map = {
    'Filed': 'badge-filed', 'Granted': 'badge-granted',
    'Under Examination': 'badge-exam', 'Published': 'badge-exam',
    'Drafted': 'badge-drafted', 'Abandoned': 'badge-abandoned',
    'Lapsed': 'badge-abandoned', 'Refused': 'badge-abandoned',
    'Paid': 'badge-paid', 'Unpaid': 'badge-unpaid',
    'Active': 'badge-active', 'Inactive': 'badge-inactive',
    'Admin': 'badge-admin', 'Attorney': 'badge-attorney', 'Client': 'badge-client',
  };
  return `<span class="badge ${map[s] || 'badge-drafted'}">${s}</span>`;
}
function err(msg) { return `<div class="alert alert-error">${msg}</div>`; }
function showPageError(context, error) {
  const message = error && error.message ? error.message : String(error);
  const content = document.getElementById('page-content');
  if (!content) return;
  content.innerHTML = `
    <div class="card">
      <div class="card-title">Error</div>
      <div class="alert alert-error">${message}</div>
      <div class="text-muted" style="margin-top:10px;font-size:12px">Context: ${context}</div>
      ${error && error.stack ? `<pre style="white-space:pre-wrap;margin-top:12px;font-size:11px;color:var(--text-2)">${error.stack}</pre>` : ''}
    </div>`;
  try { API.reportError(context, error); } catch {}
}

let CLIENT_LOOKUP_CACHE = [];
let USER_LOOKUP_CACHE = [];
let CIRCLE_LOOKUP_CACHE = [];
let ORG_LOOKUP_CACHE = [];
let DASHBOARD_FILTER_OPEN = false;
let DAILY_AUDIT_FILTER_OPEN = false;

async function ensureClientLookupLoaded() {
  if (CLIENT_LOOKUP_CACHE.length) return CLIENT_LOOKUP_CACHE;
  const clients = await API.getClients();
  CLIENT_LOOKUP_CACHE = (Array.isArray(clients) ? clients : []).sort((a, b) =>
    String(a.CLIENT_CODE || a.CLIENT_ID || '').localeCompare(String(b.CLIENT_CODE || b.CLIENT_ID || ''))
  );
  return CLIENT_LOOKUP_CACHE;
}

function getEffectiveRoleList(user) {
  const roles = [];
  if (user?.ROLE) roles.push(user.ROLE);
  String(user?.ADDITIONAL_ROLES || '').split(',').map(v => v.trim()).filter(Boolean).forEach(role => {
    if (!roles.includes(role)) roles.push(role);
  });
  return roles;
}

async function ensureUserLookupLoaded(force = false) {
  if (!force && USER_LOOKUP_CACHE.length) return USER_LOOKUP_CACHE;
  const users = await API.getUsers({});
  USER_LOOKUP_CACHE = (Array.isArray(users) ? users : []).sort((a, b) =>
    String(a.FULL_NAME || '').localeCompare(String(b.FULL_NAME || ''))
  );
  return USER_LOOKUP_CACHE;
}

async function ensureCircleLookupLoaded(force = false) {
  if (!force && CIRCLE_LOOKUP_CACHE.length) return CIRCLE_LOOKUP_CACHE;
  const circles = await API.getCircles();
  CIRCLE_LOOKUP_CACHE = (Array.isArray(circles) ? circles : []).sort((a, b) =>
    String(a.CIRCLE_NAME || '').localeCompare(String(b.CIRCLE_NAME || ''))
  );
  return CIRCLE_LOOKUP_CACHE;
}

async function ensureOrgLookupLoaded(force = false) {
  if (!force && ORG_LOOKUP_CACHE.length) return ORG_LOOKUP_CACHE;
  const orgs = await API.getOrganizations();
  ORG_LOOKUP_CACHE = (Array.isArray(orgs) ? orgs : []).sort((a, b) =>
    String(a.ORG_ID || '').localeCompare(String(b.ORG_ID || ''))
  );
  return ORG_LOOKUP_CACHE;
}

function buildClientOptions(list) {
  return (list || []).map(c => {
    const code = c.CLIENT_CODE || c.CLIENT_ID || '';
    return `<option value="${code}">${code} - ${c.CLIENT_NAME || ''}</option>`;
  }).join('');
}

function buildUserOptions(list, role) {
  return (list || [])
    .filter(user => !role || getEffectiveRoleList(user).includes(role))
    .map(user => `<option value="${user.EMAIL}">${user.FULL_NAME || user.EMAIL} (${user.EMAIL})</option>`)
    .join('');
}

function buildUserIdentityOptions(list, filterFn) {
  return (list || [])
    .filter(user => !filterFn || filterFn(user))
    .map(user => {
      const org = user.ORG_ID ? resolveOrganization(user.ORG_ID) : null;
      const summary = [user.USER_ID, user.FULL_NAME || user.EMAIL, org?.ORG_NAME || user.ORG_ID || ''].filter(Boolean).join(' | ');
      return `<option value="${user.USER_ID}">${summary}</option>`;
    }).join('');
}

function buildUserNameEmailOptions(list, filterFn) {
  return (list || [])
    .filter(user => !filterFn || filterFn(user))
    .map(user => {
      const org = user.ORG_ID ? resolveOrganization(user.ORG_ID) : null;
      const summary = [user.FULL_NAME || user.EMAIL, user.EMAIL, org?.ORG_NAME || user.ORG_ID || ''].filter(Boolean).join(' | ');
      return `<option value="${user.EMAIL}">${summary}</option>`;
    }).join('');
}

function buildNameOptions(list) {
  return (list || []).map(user => `<option value="${user.FULL_NAME || ''}">${user.FULL_NAME || ''}</option>`).join('');
}

function filterClientsForOrgAttach(orgId) {
  return (CLIENT_LOOKUP_CACHE || []).filter(client => String(client.STATUS || 'Active') !== 'Deleted' && String(client.ORG_ID || '') !== String(orgId || ''));
}

function renderSearchableClientAttachList(orgId, query = '') {
  const normalized = String(query || '').trim().toLowerCase();
  const list = filterClientsForOrgAttach(orgId).filter(client => {
    if (!normalized) return true;
    return [client.CLIENT_ID, client.CLIENT_CODE, client.CLIENT_NAME, client.EMAIL, client.CONTACT_PERSON]
      .join(' ')
      .toLowerCase()
      .includes(normalized);
  });
  const target = document.getElementById('org-client-pick-results');
  if (!target) return;
  target.innerHTML = list.length ? list.map(client => `
    <button class="chat-list-item" style="margin-bottom:8px" onclick="attachClientToOrganization('${orgId}','${client.CLIENT_ID}')">
      <div class="chat-avatar">${String(client.CLIENT_NAME || client.CLIENT_CODE || 'C').slice(0,1).toUpperCase()}</div>
      <div class="chat-meta">
        <div class="chat-head"><div class="chat-name">${client.CLIENT_NAME || client.CLIENT_ID}</div><div class="chat-time">${client.CLIENT_TYPE || 'Client'}</div></div>
        <div class="chat-preview">${[client.CLIENT_CODE || client.CLIENT_ID, client.EMAIL || '', client.CONTACT_PERSON || ''].filter(Boolean).join(' | ')}</div>
      </div>
    </button>`).join('') : '<div class="text-muted" style="padding:12px 0">No matching clients found.</div>';
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsDataURL(file);
  });
}

function normalizeExcelImportDate(value) {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().split('T')[0];
  if (typeof value === 'number' && window.XLSX?.SSF?.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      const date = new Date(parsed.y, parsed.m - 1, parsed.d);
      if (!Number.isNaN(date.getTime())) return date.toISOString().split('T')[0];
    }
  }
  const parsedDate = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? '' : parsedDate.toISOString().split('T')[0];
}

function normalizeImportRecordType(recordType) {
  const normalized = String(recordType || '').trim().toLowerCase();
  if (normalized === 'patent') return 'Patent';
  if (normalized === 'trademark') return 'Trademark';
  return '';
}

function updateBulkImportProgress(percent, message, detailsHtml = '') {
  const bar = document.getElementById('bulk-import-progress-bar');
  const label = document.getElementById('bulk-import-progress-label');
  const details = document.getElementById('bulk-import-progress-details');
  if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  if (label) label.textContent = message || '';
  if (details) details.innerHTML = detailsHtml || '';
}

async function parseDocketTrakWorkbook(file) {
  const sourceRows = await readDocketTrakRows_(file);
  if (!sourceRows.length) throw new Error('The uploaded sheet is empty.');

  const parsedRows = sourceRows.map((row, index) => {
    const recordType = normalizeImportRecordType(row['Record Type']);
    const title = String(row['Description/Title'] || '').trim();
    const issues = [];
    if (!recordType) issues.push('Unsupported record type');
    if (!title) issues.push('Missing title');
    return {
      sourceRowNumber: index + 2,
      recordType,
      title,
      country: String(row['Country'] || '').trim() || 'India',
      applicationNumber: String(row['Reference Number'] || '').trim(),
      referenceNumber: String(row['Reference Number'] || '').trim(),
      docketingEvent: String(row['Docketing Event'] || '').trim(),
      sourceStatus: String(row['Status'] || '').trim(),
      sourceClient: String(row['Client'] || '').trim(),
      eventDate: normalizeExcelImportDate(row['Docketing Event Date']),
      notes: '',
      issues,
      importable: issues.length === 0
    };
  });

  const importableRows = parsedRows.filter(row => row.importable);
  const skippedRows = parsedRows.filter(row => !row.importable);
  return {
    fileName: file.name,
    parsedRows,
    importableRows,
    skippedRows,
    headers: Object.keys(sourceRows[0] || {}),
    patentCount: importableRows.filter(row => row.recordType === 'Patent').length,
    trademarkCount: importableRows.filter(row => row.recordType === 'Trademark').length
  };
}

async function readDocketTrakRows_(file) {
  const arrayBuffer = await file.arrayBuffer();
  const primaryError = [];
  if (window.XLSX) {
    try {
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, WTF: false });
      const sheet = workbook.Sheets['IP Records'] || workbook.Sheets[workbook.SheetNames[0]];
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
        if (rows.length) return rows;
      }
    } catch (e) {
      primaryError.push(e?.message || String(e));
    }
  }
  if (window.JSZip) {
    try {
      return await readDocketTrakRowsFromZip_(arrayBuffer);
    } catch (e) {
      primaryError.push(e?.message || String(e));
    }
  }
  throw new Error(primaryError[0] || 'Could not read the uploaded Excel file.');
}

function columnLettersToIndex_(letters) {
  let index = 0;
  const text = String(letters || '').toUpperCase();
  for (let i = 0; i < text.length; i += 1) {
    index = (index * 26) + (text.charCodeAt(i) - 64);
  }
  return Math.max(0, index - 1);
}

function excelSerialToIsoDate_(serial) {
  const numeric = Number(serial);
  if (!Number.isFinite(numeric)) return String(serial || '');
  const utcDays = Math.floor(numeric - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  return Number.isNaN(dateInfo.getTime()) ? String(serial || '') : dateInfo.toISOString().split('T')[0];
}

function getXmlTextContent_(node) {
  if (!node) return '';
  return Array.from(node.getElementsByTagName('*'))
    .filter(child => child.tagName && child.tagName.split(':').pop() === 't')
    .map(child => child.textContent || '')
    .join('');
}

async function readDocketTrakRowsFromZip_(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();

  async function readXml(path) {
    const file = zip.file(path);
    if (!file) return null;
    const xml = await file.async('string');
    return parser.parseFromString(xml, 'application/xml');
  }

  const workbookXml = await readXml('xl/workbook.xml');
  if (!workbookXml) throw new Error('Missing workbook.xml');
  const relsXml = await readXml('xl/_rels/workbook.xml.rels');
  if (!relsXml) throw new Error('Missing workbook relationships');

  const workbookSheets = Array.from(workbookXml.getElementsByTagName('*')).filter(node => node.tagName && node.tagName.split(':').pop() === 'sheet');
  const targetSheet = workbookSheets.find(node => String(node.getAttribute('name') || '').trim() === 'IP Records') || workbookSheets[0];
  if (!targetSheet) throw new Error('No worksheet definition found');

  const relId = targetSheet.getAttribute('r:id') || targetSheet.getAttribute('id');
  const relNodes = Array.from(relsXml.getElementsByTagName('*')).filter(node => node.tagName && node.tagName.split(':').pop() === 'Relationship');
  const relationship = relNodes.find(node => String(node.getAttribute('Id') || '') === String(relId || ''));
  if (!relationship) throw new Error('Worksheet relationship not found');

  let targetPath = relationship.getAttribute('Target') || '';
  if (!targetPath) throw new Error('Worksheet target not found');
  if (!/^xl\//i.test(targetPath)) targetPath = `xl/${targetPath.replace(/^\//, '')}`;
  const sheetXml = await readXml(targetPath);
  if (!sheetXml) throw new Error('Worksheet XML could not be read');

  const sharedStringsXml = await readXml('xl/sharedStrings.xml');
  const sharedStrings = sharedStringsXml
    ? Array.from(sharedStringsXml.getElementsByTagName('*'))
        .filter(node => node.tagName && node.tagName.split(':').pop() === 'si')
        .map(getXmlTextContent_)
    : [];

  const rows = Array.from(sheetXml.getElementsByTagName('*')).filter(node => node.tagName && node.tagName.split(':').pop() === 'row');
  if (!rows.length) return [];

  const matrix = rows.map(rowNode => {
    const cells = [];
    Array.from(rowNode.getElementsByTagName('*'))
      .filter(node => node.tagName && node.tagName.split(':').pop() === 'c')
      .forEach(cellNode => {
        const ref = cellNode.getAttribute('r') || '';
        const colLetters = (ref.match(/[A-Z]+/i) || [''])[0];
        const colIndex = columnLettersToIndex_(colLetters);
        const type = cellNode.getAttribute('t') || '';
        const valueNode = Array.from(cellNode.childNodes).find(node => node.nodeType === 1 && node.tagName && node.tagName.split(':').pop() === 'v');
        const inlineNode = Array.from(cellNode.childNodes).find(node => node.nodeType === 1 && node.tagName && node.tagName.split(':').pop() === 'is');
        let value = '';
        if (type === 's' && valueNode) value = sharedStrings[Number(valueNode.textContent || 0)] || '';
        else if (type === 'inlineStr' && inlineNode) value = getXmlTextContent_(inlineNode);
        else if (valueNode) value = valueNode.textContent || '';
        cells[colIndex] = value;
      });
    return cells;
  });

  const headers = (matrix[0] || []).map(value => String(value || '').trim());
  return matrix.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (!header) return;
      let value = row[index] != null ? row[index] : '';
      if (/Date/i.test(header)) value = excelSerialToIsoDate_(value);
      obj[header] = value;
    });
    return obj;
  }).filter(row => Object.keys(row).length > 0);
}

function getStoredTheme() {
  try {
    return localStorage.getItem('portalTheme') || 'dark';
  } catch (e) {
    return 'dark';
  }
}

function applyTheme(theme) {
  CURRENT_THEME = theme === 'light' ? 'light' : 'dark';
  document.body.classList.toggle('theme-light', CURRENT_THEME === 'light');
  try { localStorage.setItem('portalTheme', CURRENT_THEME); } catch (e) {}
}

function getThemeToggleHtml() {
  return `<button class="btn btn-ghost btn-sm" onclick="toggleTheme()">${CURRENT_THEME === 'light' ? 'Dark' : 'Light'} Theme</button>`;
}

function setTopbarActions(extraHtml = '') {
  window.__lastTopbarExtra = extraHtml || '';
  const right = document.getElementById('topbar-right');
  if (!right) return;
  right.innerHTML = `${getThemeToggleHtml()}${extraHtml ? `<span class="topbar-actions-extra">${extraHtml}</span>` : ''}`;
}

function toggleTheme() {
  applyTheme(CURRENT_THEME === 'light' ? 'dark' : 'light');
  setTopbarActions(window.__lastTopbarExtra || '');
  if (typeof CURRENT_PAGE === 'string' && CURRENT_PAGE) {
    showPage(CURRENT_PAGE);
  }
}

function buildCaseIdOptions(list) {
  return (list || []).map(c => `<option value="${c}">${c}</option>`).join('');
}

function buildCircleOptions(list) {
  return (list || []).map(circle => `<option value="${circle.CIRCLE_ID}">${circle.CIRCLE_NAME}</option>`).join('');
}

function buildOrgOptions(list) {
  return (list || []).map(org => `<option value="${org.ORG_ID}">${org.ORG_ID} | ${org.ORG_NAME || ''}</option>`).join('');
}

function resolveClientByCode(code) {
  const lookup = CLIENT_LOOKUP_CACHE || [];
  const normalized = String(code || '').trim().toUpperCase();
  return lookup.find(c => String(c.CLIENT_CODE || c.CLIENT_ID || '').toUpperCase() === normalized) || null;
}

function resolveOrganization(value) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!normalized) return null;
  return (ORG_LOOKUP_CACHE || []).find(org => {
    const id = String(org.ORG_ID || '').toUpperCase();
    const name = String(org.ORG_NAME || '').toUpperCase();
    return id === normalized || `${id} | ${name}` === normalized || name === normalized;
  }) || null;
}

function resolveUserById(value) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!normalized) return null;
  return (USER_LOOKUP_CACHE || []).find(user => String(user.USER_ID || '').toUpperCase() === normalized) || null;
}

function getCaseIdLookupFromCases(cases) {
  return Array.from(new Set((cases || []).map(c => c.CASE_ID).filter(Boolean))).sort();
}

function renderRoleCheckboxes(selectedCsv) {
  const selected = String(selectedCsv || '').split(',').map(v => v.trim()).filter(Boolean);
  const roles = ['Super Admin','Admin','Galvanizer','Staff','Attorney','Client Admin','Client Employee','Individual Client'];
  return roles.map(role => `
    <label style="display:inline-flex;align-items:center;gap:6px;margin:0 12px 8px 0;font-size:13px">
      <input type="checkbox" class="mu-extra-role" value="${role}" ${selected.includes(role) ? 'checked' : ''} />
      <span>${role}</span>
    </label>`).join('');
}

function getSelectedAdditionalRoles() {
  return Array.from(document.querySelectorAll('.mu-extra-role:checked')).map(el => el.value).join(', ');
}

function startSafePrefetch() {
  if (SAFE_PREFETCH_STARTED) return;
  SAFE_PREFETCH_STARTED = true;
  Promise.allSettled([
    ensureClientLookupLoaded(),
    ensureUserLookupLoaded(),
    ensureCircleLookupLoaded(),
    API.getOrganizations(),
    API.getCasesPage({}, CASES_PAGE_LIMIT, 0)
  ]).catch(() => {});
}

// ── Destroy charts helper ─────────────────────────────────────────────────────
function destroyCharts() {
  Object.values(_charts).forEach(c => { try { c.destroy(); } catch {} });
  _charts = {};
}

// ═══════════════════════════════════════════════════════════════════════════════
//  DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
async function loadDashboard() {
  destroyCharts();
  try {
    await ensureClientLookupLoaded();
    const cachedSummary = API.getCachedDashboardSummary(DASHBOARD_FILTERS);
    const cachedDetails = API.getCachedDashboardDetails(DASHBOARD_FILTERS);
    if (cachedSummary) {
      renderDashboardSummary(cachedSummary);
      if (cachedDetails) {
        hydrateDashboardDetails(cachedSummary, cachedDetails);
      }
    }
    const summary = await API.getDashboardSummary(DASHBOARD_FILTERS);
    renderDashboardSummary(summary);
    if (!cachedDetails) {
      const loadingBlock = document.getElementById('dashboard-detail-loading');
      if (loadingBlock) {
        loadingBlock.style.display = '';
      }
    }
    const details = await API.getDashboardDetails(DASHBOARD_FILTERS);
    hydrateDashboardDetails(summary, details);
  } catch(e) {
    showPageError('loadDashboard', e);
  }
}

function renderDashboardSummary(d) {
  const caseIds = [];
  setTopbarActions(`<button class="btn btn-ghost btn-sm" onclick="toggleDashboardFilters()">${DASHBOARD_FILTER_OPEN ? 'Close Filter' : 'Filter'}</button>`);
  document.getElementById('page-content').innerHTML = `
  <div class="page">
    <div class="filter-toggle-row"></div>
    <div class="card filter-panel filter-drawer ${DASHBOARD_FILTER_OPEN ? 'open' : ''}" style="margin-bottom:20px">
      <div class="filter-head">
        <div class="card-title">Dashboard Filters</div>
        <div class="filter-actions">
          <button class="btn btn-primary btn-sm" onclick="applyDashboardFilters()">Apply</button>
          <button class="btn btn-ghost btn-sm" onclick="resetDashboardFilters()">Reset</button>
        </div>
      </div>
      <div class="filter-grid compact-2">
        <div class="form-group">
          <label class="form-label">Client Code</label>
          <input class="form-control" id="dash-client-code" list="dash-client-options" value="${DASHBOARD_FILTERS.clientCode || ''}" placeholder="A61M or 870Y" />
          <datalist id="dash-client-options">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">Case ID</label>
          <input class="form-control" id="dash-case-id" list="dash-case-options" value="${DASHBOARD_FILTERS.caseId || ''}" placeholder="A61M002" />
          <datalist id="dash-case-options">${buildCaseIdOptions(caseIds)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">From Date</label>
          <input type="date" class="form-control" id="dash-from-date" value="${DASHBOARD_FILTERS.fromDate || ''}" />
        </div>
        <div class="form-group">
          <label class="form-label">To Date</label>
          <input type="date" class="form-control" id="dash-to-date" value="${DASHBOARD_FILTERS.toDate || ''}" />
        </div>
      </div>
    </div>
    <div class="stats-grid">
      <div class="stat-card accent-primary"><div class="stat-label">Granted</div><div class="stat-value">${d.totalGranted || 0}</div><div class="stat-sub">patents</div></div>
      <div class="stat-card accent-sky"><div class="stat-label">Pending</div><div class="stat-value">${d.totalPending || 0}</div><div class="stat-sub">active cases</div></div>
      <div class="stat-card accent-amber"><div class="stat-label">Deadlines</div><div class="stat-value">${d.upcomingDeadlineCount || 0}</div><div class="stat-sub">upcoming 30d</div></div>
      <div class="stat-card accent-rose"><div class="stat-label">Unpaid</div><div class="stat-value">${d.pendingInvoicesCount || 0}</div><div class="stat-sub">invoices</div></div>
      <div class="stat-card accent-sky"><div class="stat-label">Unread Alerts</div><div class="stat-value">${d.unreadNotifications || 0}</div><div class="stat-sub">notifications</div></div>
      <div class="stat-card accent-primary"><div class="stat-label">Open Threads</div><div class="stat-value">${d.openThreads || 0}</div><div class="stat-sub">messages</div></div>
      <div class="stat-card accent-emerald"><div class="stat-label">My Clients</div><div class="stat-value">${d.myClientCount || 0}</div><div class="stat-sub">in scope</div></div>
    </div>
    <div id="dashboard-detail-loading" class="card" style="margin-top:20px">
      <div class="card-title">Loading details</div>
      <div class="loading-wrap" style="min-height:220px"><div class="spinner"></div><span>Loading charts and activity…</span></div>
    </div>
    <div id="dashboard-detail-block"></div>
  </div>`;
}

function hydrateDashboardDetails(summary, details) {
  const d = Object.assign({}, summary || {}, details || {});
  renderDashboard(d);
  const caseIds = getCaseIdLookupFromCases((details && details.recentActiveCases) || []);
  const caseList = document.getElementById('dash-case-options');
  if (caseList) {
    caseList.innerHTML = buildCaseIdOptions(caseIds);
  }
}

function renderDashboard(d) {
  const COUNTRY_COLORS = {
    India:'#818cf8', USA:'#38bdf8', EPO:'#fb7185', PCT:'#fbbf24',
    Japan:'#34d399', China:'#f97316', Other:'#6b7097'
  };

  const deadlineRows = (d.upcomingDeadlines || []).map(u => `
    <div class="list-entry">
      <div>
        <div class="list-entry-label">Docket</div>
        <div class="list-entry-val">${u.docket}</div>
      </div>
      <div style="text-align:right">
        <div class="list-entry-label">${u.country} · Due</div>
        <div class="list-entry-val" style="color:var(--amber)">${u.dateStr}</div>
      </div>
    </div>`).join('') || '<div class="empty-state text-muted" style="padding:20px">No upcoming deadlines</div>';

  const invoiceRows = (d.pendingInvoicesList || []).map(inv => `
    <div class="list-entry">
      <div>
        <div class="list-entry-label">Invoice</div>
        <div class="list-entry-val">${inv.docket}</div>
        <div class="list-entry-label" style="margin-top:3px">${inv.dateStr}</div>
      </div>
      <div class="list-entry-amount">${money(inv.amount)}</div>
    </div>`).join('') || '<div class="empty-state text-muted" style="padding:20px">No pending payments</div>';

  const actionRows = (d.recentActiveCases || []).map(rc => `
    <tr>
      <td class="text-em">${rc.docket}</td>
      <td>${rc.date}</td>
      <td>${statusBadge(rc.status)}</td>
      <td class="text-muted" style="font-size:12px">Review case</td>
    </tr>`).join('') || '<tr><td colspan="4" class="text-muted" style="text-align:center;padding:20px">No active cases pending action</td></tr>';

  const caseIds = getCaseIdLookupFromCases(d.recentActiveCases || []);
  setTopbarActions(`<button class="btn btn-ghost btn-sm" onclick="toggleDashboardFilters()">${DASHBOARD_FILTER_OPEN ? 'Close Filter' : 'Filter'}</button>`);
  document.getElementById('page-content').innerHTML = `
  <div class="page">
    <div class="filter-toggle-row"></div>
    <div class="card filter-panel filter-drawer ${DASHBOARD_FILTER_OPEN ? 'open' : ''}" style="margin-bottom:20px">
      <div class="filter-head">
        <div class="card-title">Dashboard Filters</div>
        <div class="filter-actions">
          <button class="btn btn-primary btn-sm" onclick="applyDashboardFilters()">Apply</button>
          <button class="btn btn-ghost btn-sm" onclick="resetDashboardFilters()">Reset</button>
        </div>
      </div>
      <div class="filter-grid compact-2">
        <div class="form-group">
          <label class="form-label">Client Code</label>
          <input class="form-control" id="dash-client-code" list="dash-client-options" value="${DASHBOARD_FILTERS.clientCode || ''}" placeholder="A61M or 870Y" />
          <datalist id="dash-client-options">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">Case ID</label>
          <input class="form-control" id="dash-case-id" list="dash-case-options" value="${DASHBOARD_FILTERS.caseId || ''}" placeholder="A61M002" />
          <datalist id="dash-case-options">${buildCaseIdOptions(caseIds)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">From Date</label>
          <input type="date" class="form-control" id="dash-from-date" value="${DASHBOARD_FILTERS.fromDate || ''}" />
        </div>
        <div class="form-group">
          <label class="form-label">To Date</label>
          <input type="date" class="form-control" id="dash-to-date" value="${DASHBOARD_FILTERS.toDate || ''}" />
        </div>
      </div>
    </div>
    <!-- Stat strip -->
    <div class="stats-grid">
      <div class="stat-card accent-primary">
        <div class="stat-label">Granted</div>
        <div class="stat-value">${d.totalGranted || 0}</div>
        <div class="stat-sub">patents</div>
      </div>
      <div class="stat-card accent-sky">
        <div class="stat-label">Pending</div>
        <div class="stat-value">${d.totalPending || 0}</div>
        <div class="stat-sub">active cases</div>
      </div>
      <div class="stat-card accent-amber">
        <div class="stat-label">Deadlines</div>
        <div class="stat-value">${(d.upcomingDeadlines||[]).length}</div>
        <div class="stat-sub">upcoming 30d</div>
      </div>
      <div class="stat-card accent-rose">
        <div class="stat-label">Unpaid</div>
        <div class="stat-value">${(d.pendingInvoicesList||[]).length}</div>
        <div class="stat-sub">invoices</div>
      </div>
      <div class="stat-card accent-sky">
        <div class="stat-label">Unread Alerts</div>
        <div class="stat-value">${d.unreadNotifications || 0}</div>
        <div class="stat-sub">notifications</div>
      </div>
      <div class="stat-card accent-primary">
        <div class="stat-label">Open Threads</div>
        <div class="stat-value">${d.openThreads || 0}</div>
        <div class="stat-sub">messages</div>
      </div>
      <div class="stat-card accent-emerald">
        <div class="stat-label">My Clients</div>
        <div class="stat-value">${d.myClientCount || 0}</div>
        <div class="stat-sub">in scope</div>
      </div>
    </div>

    <!-- Row 1 -->
    <div class="grid-7-5" style="margin-bottom:20px">
      <div class="card">
        <div style="display:flex;justify-content:space-around;margin-bottom:20px">
          <span style="font-family:'Syne',sans-serif;font-weight:700;font-size:13px">Granted Patents</span>
          <span style="font-family:'Syne',sans-serif;font-weight:700;font-size:13px">Pending Patents</span>
        </div>
        <div style="display:flex;justify-content:space-around;align-items:center">
          <div class="donut-wrap"><canvas id="chart-granted"></canvas><div class="donut-center"><div class="donut-num">${d.totalGranted||0}</div><div class="donut-lbl">granted</div></div></div>
          <div class="donut-wrap"><canvas id="chart-pending"></canvas><div class="donut-center"><div class="donut-num">${d.totalPending||0}</div><div class="donut-lbl">pending</div></div></div>
        </div>
        <div class="chart-legend">
          ${Object.keys(Object.assign({}, d.grantedByCountry, d.pendingByCountry)).filter((v,i,a)=>a.indexOf(v)===i).map(c=>`
          <div class="legend-item"><div class="legend-dot" style="background:${COUNTRY_COLORS[c]||'#6b7097'}"></div><span>${c}</span></div>`).join('')}
        </div>
      </div>

      <div style="display:flex;flex-direction:column;gap:16px">
        <div class="card">
          <div class="card-title">Upcoming Renewals</div>
          ${deadlineRows}
        </div>
        <div class="card">
          <div class="card-title">Pending Payments</div>
          ${invoiceRows}
          <div style="text-align:right;margin-top:8px"><button class="btn btn-ghost btn-sm" onclick="showPage('finance')">View all →</button></div>
        </div>
      </div>
    </div>

    <!-- Row 2 -->
    <div class="grid-2">
      <div class="card">
        <div class="card-title">Pending by Status</div>
        <div style="position:relative;height:200px"><canvas id="chart-status"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Action Required</div>
        <div class="table-wrap"><table>
          <thead><tr><th>Docket</th><th>Filed</th><th>Status</th><th>Action</th></tr></thead>
          <tbody>${actionRows}</tbody>
        </table></div>
      </div>
    </div>
  </div>`;

  // Draw charts
  const styles = getComputedStyle(document.body);
  const chartText = styles.getPropertyValue('--chart-text').trim() || styles.getPropertyValue('--text-2').trim() || '#9191b4';
  const chartGrid = styles.getPropertyValue('--chart-grid').trim() || 'rgba(255,255,255,0.08)';
  const chartEmpty = styles.getPropertyValue('--chart-empty').trim() || '#1e1e30';
  const primaryDark = styles.getPropertyValue('--primary-d').trim() || '#6366f1';
  const primary = styles.getPropertyValue('--primary').trim() || '#818cf8';
  Chart.defaults.color = chartText;
  Chart.defaults.font.family = "'DM Sans', sans-serif";

  const makeDonut = (canvasId, dataObj) => {
    const labels = Object.keys(dataObj);
    const data = labels.map(k => dataObj[k]);
    const bgColors = labels.map(k => COUNTRY_COLORS[k] || '#6b7097');
    if (data.length === 0 || data.every(v => v === 0)) {
      labels.push('None'); data.push(1); bgColors.push(chartEmpty);
    }
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    _charts[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: bgColors, borderWidth: 0, cutout: '78%' }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.label}: ${c.raw}` } } }
      }
    });
  };

  makeDonut('chart-granted', d.grantedByCountry || {});
  makeDonut('chart-pending', d.pendingByCountry || {});

  const statusObj = d.pendingByStatus || {};
  const sLabels = Object.keys(statusObj);
  const sData = sLabels.map(k => statusObj[k]);
  const barCtx = document.getElementById('chart-status');
  if (barCtx && sLabels.length) {
    _charts['chart-status'] = new Chart(barCtx, {
      type: 'bar',
      data: {
        labels: sLabels,
        datasets: [{ data: sData, backgroundColor: primaryDark, hoverBackgroundColor: primary, borderRadius: 4, barThickness: 14 }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { display: false, grid: { display: false, color: chartGrid }, border: { display: false } },
          y: { grid: { display: false, color: chartGrid }, border: { display: false }, ticks: { padding: 8, font: { size: 11 }, color: chartText } }
        }
      }
    });
  }
}

function applyDashboardFilters() {
  DASHBOARD_FILTERS = {
    clientCode: document.getElementById('dash-client-code')?.value.trim().toUpperCase() || '',
    caseId: document.getElementById('dash-case-id')?.value.trim().toUpperCase() || '',
    fromDate: document.getElementById('dash-from-date')?.value || '',
    toDate: document.getElementById('dash-to-date')?.value || '',
  };
  DASHBOARD_FILTER_OPEN = false;
  loadDashboard();
}

function resetDashboardFilters() {
  DASHBOARD_FILTERS = { clientCode: '', caseId: '', fromDate: '', toDate: '' };
  DASHBOARD_FILTER_OPEN = false;
  loadDashboard();
}

function toggleDashboardFilters() {
  DASHBOARD_FILTER_OPEN = !DASHBOARD_FILTER_OPEN;
  loadDashboard();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  CASES
// ═══════════════════════════════════════════════════════════════════════════════
async function loadCases() {
  destroyCharts();
  const canEdit = sessionHasRoleAtLeast('Staff');
  const canBulk = sessionHasRoleAtLeast('Admin');
  if (canEdit) {
    setTopbarActions(`${canBulk ? `<button class="btn btn-ghost btn-sm" onclick="openBulkCaseUpdate()">Bulk Update</button>` : ''}<button class="btn btn-ghost btn-sm" onclick="openBulkDocketTrakImport()">Bulk Import</button><button class="btn btn-primary btn-sm" onclick="openCaseModal()">+ New Case</button>`);
  } else {
    setTopbarActions('');
  }
  try {
    await ensureClientLookupLoaded();
    await ensureUserLookupLoaded();
    CASES_PAGE_OFFSET = 0;
    CASES_ITEMS = [];
    const page = await API.getCasesPage(CASE_FILTERS, CASES_PAGE_LIMIT, CASES_PAGE_OFFSET);
    CASES_ITEMS = Array.isArray(page.items) ? page.items : [];
    CASES_TOTAL = page.total || CASES_ITEMS.length;
    CASES_HAS_MORE = !!page.hasMore;
    CASES_PAGE_OFFSET = page.nextOffset || CASES_ITEMS.length;
    renderCases(CASES_ITEMS, canEdit, canBulk);
  } catch(e) {
    showPageError('loadCases', e);
  }
}

function renderCases(cases, canEdit, canBulk) {
  if (!cases.length) {
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      ${renderCaseFilters()}
      <div class="empty-state"><p>No cases found.</p></div>
    </div>`;
    return;
  }
  const rows = cases.map(c => `
    <tr>
      ${canBulk ? `<td><input type="checkbox" class="case-bulk-check" value="${c.CASE_ID}" /></td>` : ''}
      <td class="text-em" style="font-family:'Syne',sans-serif;font-size:12px">${c.CASE_ID}</td>
      <td><div style="font-weight:500;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${c.PATENT_TITLE}</div>
          <div style="font-size:11px;color:var(--text-3)">${c.APPLICATION_NUMBER||''}</div></td>
      <td>${c.COUNTRY||'—'}</td>
      <td>${fmt(c.FILING_DATE)}</td>
      <td>${statusBadge(c.CURRENT_STATUS)}</td>
      <td style="color:${c.NEXT_DEADLINE?'var(--amber)':'var(--text-3)'}">${fmt(c.NEXT_DEADLINE)}</td>
      <td style="font-size:12px;color:var(--text-2)">${c.ATTORNEY||'—'}</td>
      ${canEdit ? `<td><div class="action-row">
        <button class="btn btn-ghost btn-sm" onclick='openCaseModal(${JSON.stringify(JSON.stringify(c))})'>Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteItem('Case','${c.CASE_ID}')">Del</button>
      </div></td>` : ''}
    </tr>`).join('');

  document.getElementById('page-content').innerHTML = `
  <div class="page">
    ${renderCaseFilters()}
    <div class="card"><div class="table-wrap"><table>
      <thead><tr>${canBulk?'<th></th>':''}<th>Case ID</th><th>Patent Title</th><th>Country</th><th>Filed</th><th>Status</th><th>Deadline</th><th>Attorney</th>${canEdit?'<th>Actions</th>':''}</tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>
    <div style="display:flex;justify-content:space-between;align-items:center;margin-top:12px">
      <div class="text-muted" style="font-size:12px">${cases.length} of ${CASES_TOTAL} cases loaded</div>
      ${CASES_HAS_MORE ? `<button class="btn btn-ghost btn-sm" id="cases-load-more" onclick="loadMoreCases()">Load More</button>` : ''}
    </div>
  </div>`;
}

async function loadMoreCases() {
  const btn = document.getElementById('cases-load-more');
  if (btn) {
    btn.disabled = true;
    btn.textContent = 'Loading...';
  }
  try {
    const page = await API.getCasesPage(CASE_FILTERS, CASES_PAGE_LIMIT, CASES_PAGE_OFFSET);
    const nextItems = Array.isArray(page.items) ? page.items : [];
    CASES_ITEMS = CASES_ITEMS.concat(nextItems);
    CASES_TOTAL = page.total || CASES_ITEMS.length;
    CASES_HAS_MORE = !!page.hasMore;
    CASES_PAGE_OFFSET = page.nextOffset || CASES_ITEMS.length;
    renderCases(CASES_ITEMS, sessionHasRoleAtLeast('Staff'), sessionHasRoleAtLeast('Admin'));
  } catch (e) {
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Load More';
    }
    alert('Error loading more cases: ' + e.message);
  }
}

function renderCaseFilters() {
  const statuses = ['','Drafted','Filed','Published','Under Examination','Granted','Abandoned','Lapsed','Refused'];
  const countries = ['','India','USA','EPO','PCT','China','Japan','Other'];
  const stages = ['','Drafting','Ready for Attorney','Under Attorney Review','Filed'];
  return `
  <div class="card filter-panel">
    <div class="filter-head">
      <div class="card-title">Case Filters</div>
      <div class="filter-actions">
        <button class="btn btn-primary btn-sm" onclick="applyCaseFilters()">Apply</button>
        <button class="btn btn-ghost btn-sm" onclick="resetCaseFilters()">Reset</button>
      </div>
    </div>
    <div class="filter-grid">
      <div class="form-group">
        <label class="form-label">Client Code</label>
        <input class="form-control" id="case-filter-client" list="case-filter-clients" value="${CASE_FILTERS.clientCode || ''}" placeholder="A61M" />
        <datalist id="case-filter-clients">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
      </div>
      <div class="form-group"><label class="form-label">Case ID</label><input class="form-control" id="case-filter-caseid" value="${CASE_FILTERS.caseId || ''}" placeholder="A61M002" /></div>
      <div class="form-group"><label class="form-label">Search</label><input class="form-control" id="case-filter-query" value="${CASE_FILTERS.query || ''}" placeholder="Title, app no, attorney" /></div>
      <div class="form-group"><label class="form-label">Status</label><select class="form-control" id="case-filter-status">${statuses.map(v=>`<option value="${v}"${CASE_FILTERS.status===v?' selected':''}>${v||'All'}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Country</label><select class="form-control" id="case-filter-country">${countries.map(v=>`<option value="${v}"${CASE_FILTERS.country===v?' selected':''}>${v||'All'}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Workflow Stage</label><select class="form-control" id="case-filter-stage">${stages.map(v=>`<option value="${v}"${CASE_FILTERS.workflowStage===v?' selected':''}>${v||'All'}</option>`).join('')}</select></div>
    </div>
  </div>`;
}

function applyCaseFilters() {
  CASE_FILTERS = {
    clientCode: document.getElementById('case-filter-client')?.value.trim().toUpperCase() || '',
    caseId: document.getElementById('case-filter-caseid')?.value.trim().toUpperCase() || '',
    query: document.getElementById('case-filter-query')?.value.trim() || '',
    status: document.getElementById('case-filter-status')?.value || '',
    country: document.getElementById('case-filter-country')?.value || '',
    workflowStage: document.getElementById('case-filter-stage')?.value || '',
  };
  loadCases();
}

function resetCaseFilters() {
  CASE_FILTERS = { clientCode: '', caseId: '', status: '', country: '', workflowStage: '', query: '' };
  loadCases();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  DOCUMENTS
// ═══════════════════════════════════════════════════════════════════════════════
async function loadDocuments() {
  destroyCharts();
  try {
    await ensureClientLookupLoaded();
    const docs = await API.getDocuments();
    renderDocuments(docs);
  } catch(e) {
    showPageError('loadDocuments', e);
  }
}

async function submitDocumentUpload() {
  const clientCode = document.getElementById('docs-upload-client')?.value.trim().toUpperCase() || '';
  const category = document.getElementById('docs-upload-category')?.value || 'COMMUNICATION';
  const file = document.getElementById('docs-upload-file')?.files?.[0];
  if (!clientCode || !file) {
    alert('Select client code and document file.');
    return;
  }
  const dataUrl = await fileToDataUrl(file);
  try {
    await API.uploadPortalDocument({
      clientCode,
      subfolderName: category,
      fileName: file.name,
      mimeType: file.type,
      dataUrl
    });
    document.getElementById('docs-upload-file').value = '';
    loadDocuments();
  } catch (e) {
    alert('Document upload failed: ' + e.message);
  }
}

async function loadWorkflowBoard() {
  destroyCharts();
  try {
    const board = await API.getWorkflowBoard({});
    const stages = Object.keys(board || {});
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="grid-3">
          ${stages.map(stage => `
            <div class="card">
              <div class="card-title">${stage}</div>
              ${(board[stage] || []).slice(0, 12).map(item => `
                <div class="card" style="padding:12px;margin-bottom:10px">
                  <div class="text-em">${item.CASE_ID}</div>
                  <div style="font-weight:600">${item.PATENT_TITLE || 'Untitled case'}</div>
                  <div class="text-muted" style="font-size:12px">${item.CLIENT_NAME || item.CLIENT_ID || ''}</div>
                  <div class="text-muted" style="font-size:12px">${item.ASSIGNED_STAFF_EMAIL || 'Unassigned'} | ${item.ATTORNEY || 'No attorney'}</div>
                </div>
              `).join('') || '<div class="text-muted">No cases in this stage.</div>'}
            </div>
          `).join('')}
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadWorkflowBoard', e);
  }
}

async function loadAttorneyWorkspace() {
  destroyCharts();
  try {
    const data = await API.getAttorneyWorkspace({});
    const cases = data.cases || [];
    const tasks = data.tasks || [];
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="stats-grid">
          <div class="stat-card accent-primary"><div class="stat-label">Pending Review</div><div class="stat-value">${(data.pendingReview || []).length}</div></div>
          <div class="stat-card accent-sky"><div class="stat-label">Active Review</div><div class="stat-value">${(data.activeReview || []).length}</div></div>
          <div class="stat-card accent-amber"><div class="stat-label">My Tasks</div><div class="stat-value">${tasks.length}</div></div>
        </div>
        <div class="grid-2">
          <div class="card">
            <div class="card-title">Cases</div>
            <div class="table-wrap"><table>
              <thead><tr><th>Case</th><th>Title</th><th>Stage</th><th>Status</th></tr></thead>
              <tbody>${cases.map(item => `<tr><td class="text-em">${item.CASE_ID}</td><td>${item.PATENT_TITLE || ''}</td><td>${item.WORKFLOW_STAGE || ''}</td><td>${statusBadge(item.CURRENT_STATUS || '')}</td></tr>`).join('') || '<tr><td colspan="4" class="text-muted" style="padding:20px;text-align:center">No attorney cases.</td></tr>'}</tbody>
            </table></div>
          </div>
          <div class="card">
            <div class="card-title">Assigned Tasks</div>
            <div class="table-wrap"><table>
              <thead><tr><th>Task</th><th>Priority</th><th>Due</th><th>Status</th></tr></thead>
              <tbody>${tasks.map(item => `<tr><td>${item.TITLE || item.TASK_ID}</td><td>${item.PRIORITY || ''}</td><td>${fmt(item.DUE_DATE)}</td><td>${statusBadge(item.STATUS || '')}</td></tr>`).join('') || '<tr><td colspan="4" class="text-muted" style="padding:20px;text-align:center">No tasks.</td></tr>'}</tbody>
            </table></div>
          </div>
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadAttorneyWorkspace', e);
  }
}

async function loadTasks() {
  destroyCharts();
  setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="openTaskModal()">+ New Task</button>`);
  try {
    const tasks = await API.getTasks({});
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="card">
          <div class="table-wrap"><table>
            <thead><tr><th>ID</th><th>Title</th><th>Assignee</th><th>Priority</th><th>Due</th><th>Status</th><th></th></tr></thead>
            <tbody>${(tasks || []).map(item => `
              <tr>
                <td class="text-em">${item.TASK_ID}</td>
                <td>${item.TITLE || ''}</td>
                <td>${item.ASSIGNED_TO_NAME || item.ASSIGNED_TO_EMAIL || ''}</td>
                <td>${item.PRIORITY || ''}</td>
                <td>${fmt(item.DUE_DATE)}</td>
                <td>${statusBadge(item.STATUS || 'Open')}</td>
                <td><button class="btn btn-ghost btn-sm" onclick="advanceTaskStatus('${item.TASK_ID}','${item.STATUS || 'Open'}')">Next</button></td>
              </tr>`).join('') || '<tr><td colspan="7" class="text-muted" style="padding:20px;text-align:center">No tasks found.</td></tr>'}
            </tbody>
          </table></div>
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadTasks', e);
  }
}

async function quickCreateTask() {
  return openTaskModal();
}

async function openTaskModal() {
  await ensureUserLookupLoaded();
  await ensureClientLookupLoaded();
  openGenericModal('New Task', `
    <div class="form-group"><label class="form-label">Title *</label><input class="form-control" id="task-title" /></div>
    <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="task-description" rows="4"></textarea></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Assign To</label><input class="form-control" id="task-assigned" list="task-user-list" placeholder="Search user" /><datalist id="task-user-list">${buildUserOptions(USER_LOOKUP_CACHE, '')}</datalist></div>
      <div class="form-group"><label class="form-label">Priority</label><select class="form-control" id="task-priority"><option>Normal</option><option>High</option><option>Urgent</option></select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Due Date</label><input type="date" class="form-control" id="task-due" /></div>
      <div class="form-group"><label class="form-label">Client Code</label><input class="form-control" id="task-client" list="task-client-list" placeholder="Search client code" /><datalist id="task-client-list">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitTaskModal()">Save Task</button>
    </div>`);
}

async function submitTaskModal() {
  const assignedToEmail = document.getElementById('task-assigned').value.trim();
  const assignee = (USER_LOOKUP_CACHE || []).find(u => String(u.EMAIL || '').toLowerCase() === assignedToEmail.toLowerCase());
  const client = resolveClientByCode(document.getElementById('task-client').value.trim());
  await API.saveTask({
    title: document.getElementById('task-title').value.trim(),
    description: document.getElementById('task-description').value.trim(),
    assignedToEmail,
    assignedToName: assignee?.FULL_NAME || '',
    dueDate: document.getElementById('task-due').value,
    priority: document.getElementById('task-priority').value,
    status: 'Open',
    clientId: client?.CLIENT_ID || client?.CLIENT_CODE || ''
  });
  closeModal('modal-generic');
  loadTasks();
}

async function advanceTaskStatus(taskId, currentStatus) {
  const flow = { 'Open': 'In Progress', 'In Progress': 'Completed', 'Completed': 'Completed' };
  const next = flow[currentStatus] || 'In Progress';
  await API.updateTaskStatus(taskId, next, '');
  loadTasks();
}

async function loadTimeline() {
  destroyCharts();
  try {
    const items = await API.getActivityTimeline({});
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="card">
          ${(items || []).map(item => `
            <div style="padding:14px 0;border-bottom:1px solid var(--border)">
              <div style="display:flex;justify-content:space-between;gap:12px">
                <div>
                  <div style="font-weight:600">${item.TITLE || item.EVENT_TYPE}</div>
                  <div class="text-muted" style="font-size:13px">${item.DESCRIPTION || ''}</div>
                  <div class="text-muted" style="font-size:11px">${item.ENTITY_TYPE || ''} ${item.ENTITY_ID || ''}</div>
                </div>
                <div class="text-muted" style="font-size:11px;white-space:nowrap">${fmt(item.CREATED_AT)}</div>
              </div>
            </div>`).join('') || '<div class="text-muted">No activity yet.</div>'}
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadTimeline', e);
  }
}

async function loadSmartSearch() {
  destroyCharts();
  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card">
        <div class="form-row">
          <div class="form-group"><label class="form-label">Search</label><input class="form-control" id="smart-search-query" placeholder="Case ID, client code, name, thread, task" /></div>
          <div class="form-group"><label class="form-label">Scope</label><select class="form-control" id="smart-search-scope"><option value="all">All</option><option value="cases">Cases</option><option value="clients">Clients</option><option value="users">Users</option><option value="tasks">Tasks</option><option value="messages">Messages</option></select></div>
        </div>
        <div style="display:flex;gap:8px"><button class="btn btn-primary btn-sm" onclick="runSmartSearch()">Search</button></div>
      </div>
      <div id="smart-search-results" style="margin-top:16px"></div>
    </div>`;
}

async function runSmartSearch() {
  const query = document.getElementById('smart-search-query').value.trim();
  const scope = document.getElementById('smart-search-scope').value;
  const target = document.getElementById('smart-search-results');
  if (!query) {
    target.innerHTML = '<div class="alert alert-error">Enter a search term.</div>';
    return;
  }
  const data = await API.getSmartSearch(query, scope);
  target.innerHTML = `
    <div class="grid-2">
      <div class="card"><div class="card-title">Cases</div>${(data.cases||[]).map(item => `<div style="margin-bottom:10px"><div class="text-em">${item.CASE_ID}</div><div>${item.PATENT_TITLE || ''}</div></div>`).join('') || '<div class="text-muted">No case results.</div>'}</div>
      <div class="card"><div class="card-title">Clients</div>${(data.clients||[]).map(item => `<div style="margin-bottom:10px"><div class="text-em">${item.CLIENT_CODE || item.CLIENT_ID}</div><div>${item.CLIENT_NAME || ''}</div></div>`).join('') || '<div class="text-muted">No client results.</div>'}</div>
      <div class="card"><div class="card-title">Tasks</div>${(data.tasks||[]).map(item => `<div style="margin-bottom:10px"><div>${item.TITLE || item.TASK_ID}</div><div class="text-muted">${item.STATUS || ''}</div></div>`).join('') || '<div class="text-muted">No task results.</div>'}</div>
      <div class="card"><div class="card-title">Messages</div>${(data.threads||[]).map(item => `<div style="margin-bottom:10px"><div>${item.TITLE || item.THREAD_ID}</div><div class="text-muted">${item.THREAD_TYPE || ''}</div></div>`).join('') || '<div class="text-muted">No message results.</div>'}</div>
    </div>`;
}

async function loadApprovals() {
  destroyCharts();
  setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="openApprovalModal()">+ New Approval</button>`);
  try {
    const items = await API.getApprovalRequests({});
    const canManageAll = sessionHasRoleAtLeast('Admin');
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="card">
          <div class="table-wrap"><table>
            <thead><tr><th>ID</th><th>Title</th><th>Type</th><th>Approver</th><th>Status</th><th></th></tr></thead>
            <tbody>${(items || []).map(item => {
              const canReview = item.STATUS === 'Pending' && (canManageAll || String(item.APPROVER_EMAIL || '').toLowerCase() === String(SESSION.email || '').toLowerCase());
              return `<tr><td class="text-em">${item.APPROVAL_ID}</td><td>${item.TITLE || ''}</td><td>${item.APPROVAL_TYPE || ''}</td><td>${item.APPROVER_EMAIL || ''}</td><td>${statusBadge(item.STATUS || 'Pending')}</td><td>${canReview ? `<div class="action-row"><button class="btn btn-success btn-sm" onclick="reviewApproval('${item.APPROVAL_ID}','Approved')">Approve</button><button class="btn btn-danger btn-sm" onclick="reviewApproval('${item.APPROVAL_ID}','Rejected')">Reject</button></div>` : ''}</td></tr>`;
            }).join('') || '<tr><td colspan="6" class="text-muted" style="padding:20px;text-align:center">No approvals.</td></tr>'}</tbody>
          </table></div>
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadApprovals', e);
  }
}

async function quickCreateApproval() {
  return openApprovalModal();
}

async function openApprovalModal() {
  await ensureUserLookupLoaded();
  await ensureClientLookupLoaded();
  openGenericModal('New Approval Request', `
    <div class="form-group"><label class="form-label">Title *</label><input class="form-control" id="approval-title" /></div>
    <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="approval-description" rows="4"></textarea></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Approver</label><input class="form-control" id="approval-approver" list="approval-user-list" placeholder="Search user" /><datalist id="approval-user-list">${buildUserOptions(USER_LOOKUP_CACHE, '')}</datalist></div>
      <div class="form-group"><label class="form-label">Type</label><select class="form-control" id="approval-type"><option>General</option><option>Case</option><option>Finance</option><option>Document</option></select></div>
    </div>
    <div class="form-group"><label class="form-label">Client Code</label><input class="form-control" id="approval-client" list="approval-client-list" placeholder="Search client code" /><datalist id="approval-client-list">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitApprovalModal()">Create Approval</button>
    </div>`);
}

async function submitApprovalModal() {
  const approverEmail = document.getElementById('approval-approver').value.trim();
  const approver = (USER_LOOKUP_CACHE || []).find(u => String(u.EMAIL || '').toLowerCase() === approverEmail.toLowerCase());
  const client = resolveClientByCode(document.getElementById('approval-client').value.trim());
  await API.saveApprovalRequest({
    title: document.getElementById('approval-title').value.trim(),
    description: document.getElementById('approval-description').value.trim(),
    approverEmail,
    approverRole: approver?.ROLE || '',
    approvalType: document.getElementById('approval-type').value,
    clientId: client?.CLIENT_ID || client?.CLIENT_CODE || ''
  });
  closeModal('modal-generic');
  loadApprovals();
}

async function reviewApproval(approvalId, status) {
  try {
    await API.reviewApprovalRequest(approvalId, { status, notes: '' });
    loadApprovals();
  } catch (e) {
    alert(`Approval update failed: ${e.message}`);
  }
}

async function loadDocumentWorkflow() {
  destroyCharts();
  setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="openDocumentRequestModal()">+ New Request</button>`);
  try {
    const items = await API.getDocumentRequests({});
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="card">
          <div class="table-wrap"><table>
            <thead><tr><th>ID</th><th>Title</th><th>Case</th><th>Due</th><th>Status</th><th>Drive</th></tr></thead>
            <tbody>${(items || []).map(item => `<tr><td class="text-em">${item.REQUEST_ID}</td><td>${item.TITLE || ''}</td><td>${item.CASE_ID || ''}</td><td>${fmt(item.DUE_DATE)}</td><td>${statusBadge(item.STATUS || 'Open')}</td><td>${item.DRIVE_LINK ? `<a href="${item.DRIVE_LINK}" target="_blank">Open</a>` : '—'}</td></tr>`).join('') || '<tr><td colspan="6" class="text-muted" style="padding:20px;text-align:center">No document requests.</td></tr>'}</tbody>
          </table></div>
        </div>
      </div>`;
  } catch (e) {
    showPageError('loadDocumentWorkflow', e);
  }
}

async function quickCreateDocumentRequest() {
  return openDocumentRequestModal();
}

async function openDocumentRequestModal() {
  await ensureClientLookupLoaded();
  openGenericModal('New Document Request', `
    <div class="form-group"><label class="form-label">Title *</label><input class="form-control" id="docreq-title" /></div>
    <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="docreq-description" rows="4"></textarea></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Case ID</label><input class="form-control" id="docreq-case" placeholder="A61M001" /></div>
      <div class="form-group"><label class="form-label">Due Date</label><input type="date" class="form-control" id="docreq-due" /></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Code</label><input class="form-control" id="docreq-client" list="docreq-client-list" placeholder="Search client code" /><datalist id="docreq-client-list">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist></div>
      <div class="form-group"><label class="form-label">Visibility</label><select class="form-control" id="docreq-visible"><option value="Yes">Client Visible</option><option value="No">Internal</option></select></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitDocumentRequestModal()">Save Request</button>
    </div>`);
}

async function submitDocumentRequestModal() {
  const client = resolveClientByCode(document.getElementById('docreq-client').value.trim());
  await API.saveDocumentRequest({
    title: document.getElementById('docreq-title').value.trim(),
    description: document.getElementById('docreq-description').value.trim(),
    caseId: document.getElementById('docreq-case').value.trim(),
    dueDate: document.getElementById('docreq-due').value,
    requestType: 'Document',
    status: 'Open',
    clientVisible: document.getElementById('docreq-visible').value,
    clientId: client?.CLIENT_ID || client?.CLIENT_CODE || ''
  });
  closeModal('modal-generic');
  loadDocumentWorkflow();
}

function renderDocuments(docs) {
  if (docs.error) { document.getElementById('page-content').innerHTML = err(docs.error); return; }
  const selectedClient = (document.getElementById('docs-client-filter')?.value || '').trim().toUpperCase();
  const ICONS = { APPLICATIONS:'📝', OFFICE_ACTIONS:'📨', RESPONSES:'📤', CERTIFICATES:'🏆', INVOICES:'🧾', COMMUNICATION:'💬' };
  const sections = Object.keys(docs).map(cat => {
    const files = (docs[cat] || []).filter(f => !selectedClient || String(f.clientId || '').toUpperCase() === selectedClient || String(f.name || '').toUpperCase().includes(selectedClient));
    const items = files.length ? files.map(f => `
      <div class="doc-item">
        <div class="doc-icon"><svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="var(--primary)" stroke-width="1.5"><path d="M8 1H3a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5L8 1z"/><polyline points="8 1 8 5 12 5"/></svg></div>
        <div>
          <a href="${f.url}" target="_blank" class="doc-name">${f.name}</a>
          <div class="doc-meta">${f.date ? fmt(f.date) : ''} ${f.size > 1048576 ? (f.size/1048576).toFixed(1)+' MB' : Math.round(f.size/1024)+' KB'}</div>
        </div>
      </div>`).join('') :
      '<div class="text-muted" style="font-size:13px;padding:8px 0">No documents in this category.</div>';
    return `<div class="doc-section card" style="margin-bottom:16px">
      <div class="doc-section-title">${ICONS[cat]||'📁'} ${cat.replace(/_/g,' ')}</div>
      ${items}
    </div>`;
  }).join('');
  setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="submitDocumentUpload()">Upload Document</button>`);
  document.getElementById('page-content').innerHTML = `<div class="page">
    <div class="card" style="margin-bottom:16px">
      <div class="filter-grid compact-2">
        <div class="form-group">
          <label class="form-label">Client Code</label>
          <input class="form-control" id="docs-client-filter" list="docs-client-options" value="${selectedClient}" placeholder="Search client code" onchange="loadDocuments()" />
          <datalist id="docs-client-options">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">Category</label>
          <select class="form-control" id="docs-upload-category">
            <option>APPLICATIONS</option>
            <option>OFFICE_ACTIONS</option>
            <option>RESPONSES</option>
            <option>CERTIFICATES</option>
            <option>INVOICES</option>
            <option>COMMUNICATION</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Upload For Client</label>
          <input class="form-control" id="docs-upload-client" list="docs-upload-client-options" placeholder="Search client code" />
          <datalist id="docs-upload-client-options">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
        </div>
        <div class="form-group">
          <label class="form-label">Document File</label>
          <input type="file" class="form-control" id="docs-upload-file" />
        </div>
      </div>
    </div>
    ${sections}</div>`;
}

// ═══════════════════════════════════════════════════════════════════════════════
//  FINANCE
// ═══════════════════════════════════════════════════════════════════════════════
async function loadFinance() {
  destroyCharts();
  if (!canAccessFinance()) {
    document.getElementById('page-content').innerHTML = err('Finance access is restricted for your role.');
    return;
  }
  const canEdit = sessionHasRoleAtLeast('Admin');
  if (canEdit) {
    setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="openInvoiceModal()">+ New Invoice</button>`);
  } else {
    setTopbarActions('');
  }
  try {
    const invoices = await API.getInvoices();
    renderFinance(invoices, canEdit);
  } catch(e) {
    showPageError('loadFinance', e);
  }
}

function renderFinance(invoices, canEdit) {
  let paid=0, unpaid=0, paidSum=0, unpaidSum=0;
  invoices.forEach(inv => {
    const a = parseFloat(inv.TOTAL)||0;
    if (inv.PAYMENT_STATUS === 'Paid') { paid++; paidSum += a; }
    else { unpaid++; unpaidSum += a; }
  });

  const rows = invoices.map(inv => {
    const isPaid = inv.PAYMENT_STATUS === 'Paid';
    return `<tr>
      <td class="text-em" style="font-family:'Syne',sans-serif;font-size:12px">${inv.INVOICE_ID}</td>
      <td>${fmt(inv.INVOICE_DATE)}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${inv.SERVICE_TYPE}</td>
      <td>${money(inv.AMOUNT)}</td>
      <td>${money(inv.GST_AMOUNT)}</td>
      <td style="font-weight:600">${money(inv.TOTAL)}</td>
      <td>${statusBadge(inv.PAYMENT_STATUS)}</td>
      <td>${inv.INVOICE_PDF_LINK ? `<a href="${inv.INVOICE_PDF_LINK}" target="_blank" style="font-size:12px">View PDF</a>` : '—'}</td>
      ${canEdit ? `<td><div class="action-row">
        ${!isPaid ? `<button class="btn btn-success btn-sm" onclick="markPaid('${inv.INVOICE_ID}',this)">Paid</button>` : ''}
        <button class="btn btn-ghost btn-sm" onclick='openInvoiceModal(${JSON.stringify(JSON.stringify(inv))})'>Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteItem('Invoice','${inv.INVOICE_ID}')">Del</button>
      </div></td>` : ''}
    </tr>`;
  }).join('');

  document.getElementById('page-content').innerHTML = `
  <div class="page">
    <div class="stats-grid" style="grid-template-columns:repeat(auto-fill,minmax(150px,1fr))">
      <div class="stat-card"><div class="stat-label">Total Invoices</div><div class="stat-value">${invoices.length}</div></div>
      <div class="stat-card accent-rose"><div class="stat-label">Pending Count</div><div class="stat-value">${unpaid}</div></div>
      <div class="stat-card accent-emerald"><div class="stat-label">Paid Count</div><div class="stat-value">${paid}</div></div>
      <div class="stat-card accent-amber"><div class="stat-label">Pending Amount</div><div class="stat-value" style="font-size:18px">${money(unpaidSum)}</div></div>
      <div class="stat-card accent-emerald"><div class="stat-label">Paid Amount</div><div class="stat-value" style="font-size:18px">${money(paidSum)}</div></div>
    </div>
    <div class="card">
      <div class="card-title">Invoice Ledger</div>
      <div class="table-wrap"><table>
        <thead><tr><th>Invoice #</th><th>Date</th><th>Service</th><th>Amount</th><th>GST</th><th>Total</th><th>Status</th><th>PDF</th>${canEdit?'<th>Actions</th>':''}</tr></thead>
        <tbody>${rows}</tbody>
      </table></div>
    </div>
  </div>`;
}

async function markPaid(invoiceId, btn) {
  if (!confirm(`Mark ${invoiceId} as Paid?`)) return;
  btn.disabled = true; btn.textContent = '…';
  try {
    await API.markInvoicePaid(invoiceId);
    loadFinance();
  } catch(e) {
    btn.disabled = false; btn.textContent = 'Paid';
    alert('Error: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════
let _mgmtTab = 'users';
let ACTIVE_ORG_VIEW_ID = '';

function loadManagement() {
  destroyCharts();
  if (!canAccessManagement()) {
    document.getElementById('page-content').innerHTML = err('Management access is restricted for your role.');
    return;
  }
  document.getElementById('page-content').innerHTML = `
  <div class="page">
    <div class="tab-bar">
      <button class="tab-btn active" id="tab-users" onclick="switchMgmtTab('users')">Users</button>
      <button class="tab-btn" id="tab-clients" onclick="switchMgmtTab('clients')">Clients</button>
      <button class="tab-btn" id="tab-organizations" onclick="switchMgmtTab('organizations')">Organizations</button>
      <button class="tab-btn" id="tab-circles" onclick="switchMgmtTab('circles')">Circles</button>
    </div>
    <div id="mgmt-content"><div class="loading-wrap"><div class="spinner"></div></div></div>
  </div>`;
  switchMgmtTab('users');
}

function switchMgmtTab(tab) {
  _mgmtTab = tab;
  ['users','clients','organizations','circles'].forEach(t => {
    document.getElementById('tab-'+t)?.classList.toggle('active', t===tab);
  });
  const mc = document.getElementById('mgmt-content');
  if (!mc) return;
  mc.innerHTML = '<div class="loading-wrap"><div class="spinner"></div></div>';
  if (tab === 'users') loadMgmtUsers();
  if (tab === 'clients') loadMgmtClients();
  if (tab === 'organizations') loadMgmtOrganizationsV2();
  if (tab === 'circles') loadCircles();
}

async function loadMgmtUsers() {
  const mc = document.getElementById('mgmt-content');
  try {
    await ensureClientLookupLoaded();
    await ensureCircleLookupLoaded();
    const users = await API.getUsers(USER_FILTERS);
    const rows = users.map(u => `<tr>
      <td class="font-mono">${u.USER_ID}</td>
      <td>${u.FULL_NAME}</td>
      <td>${u.EMAIL}</td>
      <td>${statusBadge(u.ROLE)}</td>
      <td>${statusBadge(u.STATUS)}</td>
      <td><div class="action-row">
        ${String(u.EMAIL || '').toLowerCase() === String(SESSION.email || '').toLowerCase() ? '' : `<button class="btn btn-ghost btn-sm" onclick="openDirectMessageForUser('${escapeAttr(u.EMAIL)}','${escapeAttr(u.FULL_NAME)}')">Message</button>`}
        <button class="btn btn-ghost btn-sm" onclick='openUserModal(${JSON.stringify(JSON.stringify(u))})'>Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteItem('User','${u.USER_ID}')">Del</button>
      </div></td>
    </tr>`).join('');
    mc.innerHTML = `
    <div class="card" style="margin-bottom:12px">
      <div class="form-row">
        <div class="form-group"><label class="form-label">Search</label><input class="form-control" id="user-search" value="${USER_FILTERS.query || ''}" placeholder="Name, email, ID" /></div>
        <div class="form-group"><label class="form-label">Role</label><select class="form-control" id="user-role-filter">${['','Super Admin','Admin','Galvanizer','Staff','Attorney','Client Admin','Client Employee','Individual Client'].map(v=>`<option value="${v}"${USER_FILTERS.role===v?' selected':''}>${v||'All'}</option>`).join('')}</select></div>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn btn-primary btn-sm" onclick="applyUserFilters()">Apply</button>
        <button class="btn btn-ghost btn-sm" onclick="resetUserFilters()">Reset</button>
      </div>
    </div>
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px">
      <button class="btn btn-primary btn-sm" onclick="openUserModal()">+ New User</button>
    </div>
    <div class="card"><div class="table-wrap"><table>
      <thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>`;
  } catch(e) { mc.innerHTML = err(e.message); }
}

function applyUserFilters() {
  USER_FILTERS = {
    query: document.getElementById('user-search')?.value.trim() || '',
    role: document.getElementById('user-role-filter')?.value || '',
  };
  loadMgmtUsers();
}

function resetUserFilters() {
  USER_FILTERS = { query: '', role: '' };
  loadMgmtUsers();
}

async function loadMgmtClients() {
  const mc = document.getElementById('mgmt-content');
  try {
    const clients = await API.getClients();
    const rows = clients.map(c => `<tr>
      <td class="font-mono">${c.CLIENT_ID}</td>
      <td>${c.CLIENT_NAME}</td>
      <td style="font-size:12px">${c.EMAIL||'—'}</td>
      <td>${statusBadge(c.STATUS)}</td>
      <td><div class="action-row">
        <button class="btn btn-ghost btn-sm" onclick='openClientModal(${JSON.stringify(JSON.stringify(c))})'>Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteItem('Client','${c.CLIENT_ID}')">Del</button>
      </div></td>
    </tr>`).join('');
    mc.innerHTML = `
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px">
      <button class="btn btn-primary btn-sm" onclick="openClientModal()">+ New Client</button>
    </div>
    <div class="card"><div class="table-wrap"><table>
      <thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Status</th><th>Actions</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>`;
  } catch(e) { mc.innerHTML = err(e.message); }
}

async function loadMgmtCases() {
  const mc = document.getElementById('mgmt-content');
  try {
    await ensureClientLookupLoaded();
    const cases = await API.getCases(CASE_FILTERS);
    const rows = cases.map(c => `<tr>
      <td><input type="checkbox" class="case-bulk-check" value="${c.CASE_ID}" /></td>
      <td class="font-mono" style="font-size:11px">${c.CASE_ID}</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${c.PATENT_TITLE}</td>
      <td>${c.CLIENT_ID}</td>
      <td>${statusBadge(c.CURRENT_STATUS)}</td>
      <td><div class="action-row">
        <button class="btn btn-ghost btn-sm" onclick='openCaseModal(${JSON.stringify(JSON.stringify(c))})'>Edit</button>
        <button class="btn btn-danger btn-sm" onclick="deleteItem('Case','${c.CASE_ID}')">Del</button>
      </div></td>
    </tr>`).join('');
    mc.innerHTML = `
    <div style="display:flex;justify-content:flex-end;gap:8px;margin-bottom:12px">
      <button class="btn btn-ghost btn-sm" onclick="openBulkCaseUpdate()">Bulk Update</button>
      <button class="btn btn-primary btn-sm" onclick="openCaseModal()">+ New Case</button>
    </div>
    <div class="card"><div class="table-wrap"><table>
      <thead><tr><th></th><th>Case ID</th><th>Title</th><th>Client</th><th>Status</th><th>Actions</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div></div>`;
  } catch(e) { mc.innerHTML = err(e.message); }
}

// ── Delete helper ─────────────────────────────────────────────────────────────
async function deleteItem(type, id) {
  if (!confirm(`Delete ${type} ${id}? This cannot be undone.`)) return;
  try {
    const fn = {User: API.deleteUser, Client: API.deleteClient, Case: API.deleteCase, Invoice: API.deleteInvoice}[type];
    const result = await fn(id);
    if (result.success) {
      // Refresh current view
      if (type === 'Invoice') loadFinance();
      else if (type === 'Case') { if (_mgmtTab === 'cases') loadMgmtCases(); else loadCases(); }
      else if (_mgmtTab) switchMgmtTab(_mgmtTab);
    }
  } catch(e) { alert('Error: ' + e.message); }
}

async function loadMgmtOrganizations(selectedOrgId = ACTIVE_ORG_VIEW_ID) {
  const mc = document.getElementById('mgmt-content');
  try {
    await ensureOrgLookupLoaded(true);
    await ensureUserLookupLoaded();
    const orgs = await API.getOrganizations();
    ACTIVE_ORG_VIEW_ID = selectedOrgId || ACTIVE_ORG_VIEW_ID || '';
    const activeOrg = (orgs || []).find(o => o.ORG_ID === ACTIVE_ORG_VIEW_ID) || null;
    const orgUsers = activeOrg ? await API.getOrganizationUsers(activeOrg.ORG_ID) : [];
    const rows = orgs.map(o => `<tr>
      <td class="font-mono">${o.ORG_ID}</td>
      <td>${o.ORG_NAME}</td>
      <td>${o.ASSIGNED_STAFF_EMAIL || '—'}</td>
      <td>${statusBadge(o.STATUS || 'Active')}</td>
    </tr>`).join('');
    mc.innerHTML = `
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px">
      <button class="btn btn-primary btn-sm" onclick="openOrganizationPrompt()">+ New Organization</button>
    </div>
    <div class="card"><div class="table-wrap"><table>
      <thead><tr><th>ID</th><th>Organization</th><th>Assigned Staff</th><th>Status</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="4" class="text-muted" style="padding:20px;text-align:center">No organizations found.</td></tr>'}</tbody>
    </table></div></div>`;
  } catch (e) {
    mc.innerHTML = err(e.message);
  }
}

async function openOrganizationPrompt() {
  const name = prompt('Organization name');
  if (!name) return;
  const email = prompt('Primary email');
  const assignedStaffEmail = prompt('Assigned staff email (optional)') || '';
  try {
    await API.saveOrganization({ ORG_NAME: name, PRIMARY_EMAIL: email || '', ASSIGNED_STAFF_EMAIL: assignedStaffEmail, STATUS: 'Active' });
    loadMgmtOrganizations();
  } catch (e) {
    alert('Error: ' + e.message);
  }
}

async function loadMgmtOrganizationsV2(selectedOrgId = ACTIVE_ORG_VIEW_ID) {
  const mc = document.getElementById('mgmt-content');
  try {
    await ensureOrgLookupLoaded(true);
    await ensureClientLookupLoaded();
    const orgs = await API.getOrganizations();
    ACTIVE_ORG_VIEW_ID = selectedOrgId || ACTIVE_ORG_VIEW_ID || '';
    const activeOrg = (orgs || []).find(o => o.ORG_ID === ACTIVE_ORG_VIEW_ID) || null;
    const orgClients = activeOrg ? (CLIENT_LOOKUP_CACHE || []).filter(client => String(client.ORG_ID || '') === String(activeOrg.ORG_ID)) : [];
    const rows = orgs.map(o => `<tr>
      <td class="font-mono">${o.ORG_ID}</td>
      <td>${o.ORG_NAME}</td>
      <td>${o.ASSIGNED_STAFF_EMAIL || '—'}</td>
      <td>${statusBadge(o.STATUS || 'Active')}</td>
      <td><div class="action-row">
        <button class="btn btn-ghost btn-sm" onclick='openOrganizationModal(${JSON.stringify(JSON.stringify(o))})'>Edit</button>
        <button class="btn btn-ghost btn-sm" onclick="loadMgmtOrganizationsV2('${o.ORG_ID}')">Clients</button>
        <button class="btn btn-primary btn-sm" onclick="openOrganizationClientModal('${o.ORG_ID}','${escapeAttr(o.ORG_NAME || '')}')">+ Add Client</button>
      </div></td>
    </tr>`).join('');
    const userRows = (orgClients || []).map(client => `<tr>
      <td>${client.CLIENT_ID}</td>
      <td>${client.CLIENT_NAME}</td>
      <td>${client.EMAIL || '—'}</td>
      <td>${statusBadge(client.CLIENT_TYPE || 'Client')}</td>
      <td><button class="btn btn-ghost btn-sm" onclick='openClientModal(${JSON.stringify(JSON.stringify(client))})'>Open</button></td>
    </tr>`).join('');
    mc.innerHTML = `
      <div style="display:flex;justify-content:flex-end;margin-bottom:12px">
        <button class="btn btn-primary btn-sm" onclick="openOrganizationModal()">+ New Organization</button>
      </div>
      <div class="grid-2">
        <div class="card"><div class="table-wrap"><table>
          <thead><tr><th>ID</th><th>Organization</th><th>Assigned Staff</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>${rows || '<tr><td colspan="5" class="text-muted" style="padding:20px;text-align:center">No organizations found.</td></tr>'}</tbody>
        </table></div></div>
        <div class="card">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:12px">
            <div class="card-title" style="margin-bottom:0">${activeOrg ? `${activeOrg.ORG_NAME} Clients` : 'Organization Clients'}</div>
            ${activeOrg ? `<button class="btn btn-primary btn-sm" onclick="openOrganizationClientModal('${activeOrg.ORG_ID}','${escapeAttr(activeOrg.ORG_NAME || '')}')">+ Add Client</button>` : ''}
          </div>
          ${activeOrg ? `
            <div class="text-muted" style="margin:10px 0 14px">${activeOrg.ORG_ID} | ${activeOrg.PRIMARY_EMAIL || 'No email'} | Staff: ${activeOrg.ASSIGNED_STAFF_EMAIL || 'Unassigned'}</div>
            <div class="table-wrap"><table>
              <thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Type</th><th></th></tr></thead>
              <tbody>${userRows || '<tr><td colspan="5" class="text-muted" style="padding:20px;text-align:center">No organization clients yet.</td></tr>'}</tbody>
            </table></div>
          ` : '<div class="text-muted">Select an organization to view or attach clients.</div>'}
        </div>
      </div>`;
  } catch (e) {
    mc.innerHTML = err(e.message);
  }
}

async function openOrganizationModal(jsonStr) {
  await ensureUserLookupLoaded();
  const org = jsonStr ? JSON.parse(jsonStr) : null;
  const clientAdminOptions = buildUserIdentityOptions(USER_LOOKUP_CACHE, user => ['Client Admin','Client Employee'].includes(user.ROLE || ''));
  const assignedStaffOptions = buildUserNameEmailOptions(USER_LOOKUP_CACHE, user => getEffectiveRoleList(user).some(role => ['Super Admin','Admin','Galvanizer','Staff'].includes(role)));
  openGenericModal(org ? 'Edit Organization' : 'New Organization', `
    <div class="form-group"><label class="form-label">Organization Name *</label><input class="form-control" id="org-name" value="${org?.ORG_NAME || ''}" /></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Primary Email</label><input type="email" class="form-control" id="org-email" value="${org?.PRIMARY_EMAIL || ''}" /></div>
      <div class="form-group"><label class="form-label">Primary Phone</label><input class="form-control" id="org-phone" value="${org?.PRIMARY_PHONE || ''}" /></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Organization Code</label><input class="form-control" id="org-code" value="${org?.ORG_CODE || ''}" /></div>
      <div class="form-group"><label class="form-label">Assigned Staff</label><input class="form-control" id="org-staff" list="org-staff-list" value="${org?.ASSIGNED_STAFF_EMAIL || ''}" placeholder="Search staff email" /><datalist id="org-staff-list">${assignedStaffOptions}</datalist></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Admin User</label><input class="form-control" id="org-admin-user" list="org-admin-list" value="${org?.CLIENT_ADMIN_USER_ID || ''}" placeholder="Search user id" /><datalist id="org-admin-list">${clientAdminOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Status</label><select class="form-control" id="org-status"><option value="Active"${(org?.STATUS || 'Active') === 'Active' ? ' selected' : ''}>Active</option><option value="Inactive"${org?.STATUS === 'Inactive' ? ' selected' : ''}>Inactive</option></select></div>
    </div>
    <div class="form-group"><label class="form-label">Address</label><textarea class="form-control" id="org-address" rows="3">${org?.ADDRESS || ''}</textarea></div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="org-notes" rows="3">${org?.NOTES || ''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="saveOrganizationModal('${org?.ORG_ID || ''}')">Save Organization</button>
    </div>`);
}

async function saveOrganizationModal(existingOrgId) {
  const adminUser = resolveUserById(document.getElementById('org-admin-user').value.trim());
  await API.saveOrganization({
    ORG_ID: existingOrgId || undefined,
    ORG_NAME: document.getElementById('org-name').value.trim(),
    PRIMARY_EMAIL: document.getElementById('org-email').value.trim(),
    PRIMARY_PHONE: document.getElementById('org-phone').value.trim(),
    ORG_CODE: document.getElementById('org-code').value.trim(),
    CLIENT_ADMIN_USER_ID: adminUser?.USER_ID || document.getElementById('org-admin-user').value.trim(),
    ASSIGNED_STAFF_EMAIL: document.getElementById('org-staff').value.trim(),
    STATUS: document.getElementById('org-status').value,
    ADDRESS: document.getElementById('org-address').value.trim(),
    NOTES: document.getElementById('org-notes').value.trim(),
  });
  closeModal('modal-generic');
  await ensureOrgLookupLoaded(true);
  loadMgmtOrganizationsV2(existingOrgId || ACTIVE_ORG_VIEW_ID);
}

async function openOrganizationClientModal(orgId, orgName) {
  await ensureClientLookupLoaded();
  openGenericModal(`Attach Client to ${orgName}`, `
    <div class="form-group">
      <label class="form-label">Search Existing Client</label>
      <input class="form-control" id="org-client-search" placeholder="Type client code, name, email, contact" oninput="renderSearchableClientAttachList('${orgId}', this.value)" />
    </div>
    <div id="org-client-pick-results" style="max-height:360px;overflow:auto"></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Close</button>
    </div>`);
  renderSearchableClientAttachList(orgId, '');
}

async function attachClientToOrganization(orgId, clientId) {
  const client = (CLIENT_LOOKUP_CACHE || []).find(item => item.CLIENT_ID === clientId);
  if (!client) return;
  await API.saveClient({
    CLIENT_ID: client.CLIENT_ID,
    CLIENT_NAME: client.CLIENT_NAME,
    CONTACT_PERSON: client.CONTACT_PERSON || '',
    EMAIL: client.EMAIL || '',
    PHONE: client.PHONE || '',
    CLIENT_TYPE: 'Organization',
    CLIENT_REGION: client.CLIENT_REGION || 'India',
    CLIENT_CODE: client.CLIENT_CODE || '',
    ORG_ID: orgId,
    CLIENT_ADMIN_USER_ID: client.CLIENT_ADMIN_USER_ID || '',
    ASSIGNED_STAFF_EMAIL: client.ASSIGNED_STAFF_EMAIL || '',
    ADDRESS: client.ADDRESS || '',
    NOTES: client.NOTES || ''
  });
  closeModal('modal-generic');
  CLIENT_LOOKUP_CACHE = [];
  await ensureClientLookupLoaded();
  loadMgmtOrganizationsV2(orgId);
}

function truncateText(value, max = 64) {
  const text = String(value || '').trim();
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function escapeAttr(value) {
  return String(value || '').replace(/"/g, '&quot;');
}

function formatChatTime(value) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  return sameDay
    ? d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    : d.toLocaleDateString([], { day: '2-digit', month: 'short' });
}

async function loadInbox() {
  destroyCharts();
  try {
    await ensureUserLookupLoaded();
    const chats = await API.getDirectInbox();
    const chatMap = new Map((chats || []).map(chat => [String(chat.counterpartEmail || '').toLowerCase(), chat]));
    const availableUsers = (USER_LOOKUP_CACHE || [])
      .filter(user => String(user.EMAIL || '').toLowerCase() !== String(SESSION.email || '').toLowerCase())
      .filter(user => isInternalUserRole(user.ROLE || '') || getEffectiveRoleList(user).some(role => ['Super Admin','Admin','Galvanizer','Staff','Attorney'].includes(role)))
      .map(user => ({ user, chat: chatMap.get(String(user.EMAIL || '').toLowerCase()) || null }));
    setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="createDirectMessage()">+ New Chat</button>`);
    const list = availableUsers.map(({ user, chat }) => `
      <button class="chat-list-item ${chat?.unreadCount ? 'has-unread' : ''}" data-user-email="${user.EMAIL}" data-thread-id="${chat?.THREAD_ID || ''}" onclick="openInboxChatByUser('${user.EMAIL}')">
        <div class="chat-avatar">${(user.FULL_NAME || user.EMAIL || 'U').slice(0, 1).toUpperCase()}</div>
        <div class="chat-meta">
          <div class="chat-head">
            <div class="chat-name">${user.FULL_NAME || 'Unknown User'}</div>
            <div class="chat-time">${formatChatTime(chat?.LAST_MESSAGE_AT)}</div>
          </div>
          <div class="chat-role">${user.ROLE || ''}</div>
          <div class="chat-preview-row">
            <div class="chat-preview">${truncateText(chat?.lastMessageText || 'Start a direct conversation.', 54)}</div>
            ${chat?.unreadCount ? `<span class="chat-unread">${chat.unreadCount}</span>` : ''}
          </div>
        </div>
      </button>`).join('');
    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="inbox-layout">
          <div class="card inbox-sidebar">
            <div class="inbox-sidebar-head">
              <div class="card-title">Direct Chats</div>
              <button class="btn btn-ghost btn-sm" onclick="createDirectMessage()">New</button>
            </div>
            <div class="form-group" style="margin-bottom:12px">
              <input class="form-control" id="inbox-user-search" placeholder="Search user name or email" oninput="filterInboxUsers()" />
            </div>
            <div class="chat-list">
              ${list || '<div class="empty-state"><p>No direct chats yet.</p></div>'}
            </div>
          </div>
          <div class="card inbox-panel" id="inbox-panel">
            <div class="empty-state"><p>Select a chat to start messaging.</p></div>
          </div>
        </div>
      </div>`;
    if (availableUsers.length) {
      openInboxChatByUser(availableUsers[0].user.EMAIL);
    }
  } catch (e) {
    document.getElementById('page-content').innerHTML = err(e.message);
  }
}

function filterInboxUsers() {
  const query = String(document.getElementById('inbox-user-search')?.value || '').trim().toLowerCase();
  document.querySelectorAll('.chat-list-item').forEach(item => {
    item.style.display = !query || item.textContent.toLowerCase().includes(query) ? '' : 'none';
  });
}

async function openInboxChatByUser(email) {
  const result = await API.createDirectThread({ recipientEmail: email, title: 'Direct conversation' });
  if (!(result && result.threadId)) return;
  return openInboxChat(result.threadId, email);
}

async function openInboxChat(threadId, emailHint = '') {
  try {
    const [chats, messages] = await Promise.all([
      API.getDirectInbox(),
      API.getThreadMessages(threadId)
    ]);
    const chat = (chats || []).find(item => item.THREAD_ID === threadId);
    await API.markThreadRead(threadId);
    document.querySelectorAll('.chat-list-item').forEach((item) => {
      const match = item.dataset.threadId === threadId || (emailHint && String(item.dataset.userEmail || '').toLowerCase() === String(emailHint || '').toLowerCase()) || (chat && String(item.dataset.userEmail || '').toLowerCase() === String(chat.counterpartEmail || '').toLowerCase());
      item.classList.toggle('active', match);
      if (match) {
        item.dataset.threadId = threadId;
        item.classList.remove('has-unread');
        const badge = item.querySelector('.chat-unread');
        if (badge) badge.remove();
      }
    });
    const bubbles = (messages || []).map(message => {
      const mine = String(message.SENDER_EMAIL || '').toLowerCase() === String(SESSION.email || '').toLowerCase();
      return `
        <div class="chat-bubble-row ${mine ? 'mine' : 'theirs'}">
          <div class="chat-bubble ${mine ? 'mine' : 'theirs'}">
            <div class="chat-bubble-author">${mine ? 'You' : (message.SENDER_NAME || chat?.counterpartName || 'User')}</div>
            <div class="chat-bubble-text">${message.MESSAGE_TEXT || ''}</div>
            <div class="chat-bubble-time">${fmt(message.CREATED_AT)}</div>
          </div>
        </div>`;
    }).join('');
    document.getElementById('inbox-panel').innerHTML = `
      <div class="inbox-chat-head">
        <div class="chat-avatar large">${((chat && chat.counterpartName) || 'U').slice(0,1).toUpperCase()}</div>
        <div>
          <div class="card-title" style="margin-bottom:4px">${chat?.counterpartName || 'Conversation'}</div>
          <div class="text-muted">${chat?.counterpartRole || chat?.counterpartEmail || ''}</div>
        </div>
      </div>
      <div class="chat-thread">
        ${bubbles || '<div class="text-muted">No messages yet.</div>'}
      </div>
      <div class="chat-composer">
        <textarea class="form-control" id="inbox-reply" rows="3" placeholder="Type a message"></textarea>
        <div class="chat-composer-actions">
          <button class="btn btn-primary btn-sm" onclick="sendInboxReply('${threadId}')">Send</button>
        </div>
      </div>`;
    setTimeout(() => {
      const threadEl = document.querySelector('.chat-thread');
      if (threadEl) threadEl.scrollTop = threadEl.scrollHeight;
    }, 0);
  } catch (e) {
    document.getElementById('inbox-panel').innerHTML = err(e.message);
  }
}

async function sendInboxReply(threadId) {
  const field = document.getElementById('inbox-reply');
  const messageText = field?.value.trim() || '';
  if (!messageText) return;
  try {
    await API.sendThreadMessage({ threadId, messageText, isInternal: 'Yes' });
    await openInboxChat(threadId);
    field.value = '';
  } catch (e) {
    alert('Error: ' + e.message);
  }
}

async function loadMessages() {
  destroyCharts();
  try {
    const threads = (await API.getMessageThreads()).filter(t => String(t.THREAD_TYPE || '') !== 'Direct');
  setTopbarActions(`<button class="btn btn-primary btn-sm" onclick="createThread()">+ New Thread</button>`);
    const rows = threads.map(t => `<tr>
      <td class="text-em">${t.THREAD_ID}</td>
      <td>${t.TITLE || 'Untitled'}</td>
      <td>${t.THREAD_TYPE || 'General'}</td>
      <td>${statusBadge(t.STATUS || 'Open')}</td>
      <td>${fmt(t.LAST_MESSAGE_AT || t.CREATED_AT)}</td>
      <td><div class="action-row"><button class="btn btn-ghost btn-sm" onclick="openThread('${t.THREAD_ID}')">Open</button><button class="btn btn-danger btn-sm" onclick="deleteThread('${t.THREAD_ID}')">Delete</button></div></td>
    </tr>`).join('');
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="grid-2">
        <div class="card"><div class="card-title">Threads</div><div class="table-wrap"><table>
          <thead><tr><th>ID</th><th>Title</th><th>Type</th><th>Status</th><th>Updated</th><th></th></tr></thead>
          <tbody>${rows || '<tr><td colspan="6" class="text-muted" style="padding:20px;text-align:center">No threads found.</td></tr>'}</tbody>
        </table></div></div>
        <div class="card" id="thread-panel"><div class="card-title">Thread Conversation</div><div class="text-muted">Select a thread to read or reply.</div></div>
      </div>
    </div>`;
  } catch (e) {
    document.getElementById('page-content').innerHTML = err(e.message);
  }
}

async function deleteThread(threadId) {
  if (!confirm(`Delete thread ${threadId}?`)) return;
  await API.deleteMessageThread(threadId);
  loadMessages();
}

async function loadCircles(circleJson) {
  destroyCharts();
  let selectedCircle = null;
  if (circleJson) selectedCircle = JSON.parse(circleJson);
  try {
    await ensureUserLookupLoaded();
    const circles = (await API.getCircles() || []).filter(circle => String(circle.STATUS || 'Active') === 'Active');
    CIRCLE_LOOKUP_CACHE = circles || [];
    const activeCircle = selectedCircle || circles[0] || null;
    let members = [];
    if (activeCircle) {
      members = await API.getCircleMembers(activeCircle.CIRCLE_ID);
    }
    const circleCards = circles.map(circle => `
      <div class="card" style="margin-bottom:12px;border-color:${activeCircle && activeCircle.CIRCLE_ID===circle.CIRCLE_ID?'var(--primary)':'var(--border)'}">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:center">
          <div>
            <div style="font-weight:700">${circle.CIRCLE_NAME}</div>
            <div class="text-muted" style="font-size:12px">${circle.DESCRIPTION || 'No description'}</div>
          </div>
          <div class="action-row">
            <button class="btn btn-ghost btn-sm" onclick='loadCircles(${JSON.stringify(JSON.stringify(circle))})'>Open</button>
            ${sessionHasRoleAtLeast('Galvanizer') ? `<button class="btn btn-ghost btn-sm" onclick='prefillCircleForm(${JSON.stringify(JSON.stringify(circle))})'>Edit</button><button class="btn btn-danger btn-sm" onclick="deactivateCircle('${circle.CIRCLE_ID}')">Del</button>` : ''}
          </div>
        </div>
      </div>`).join('');
    const memberRows = members.map(member => `
      <tr>
        <td>${member.USER_EMAIL}</td>
        <td>${member.ROLE_IN_CIRCLE || 'Member'}</td>
        <td><button class="btn btn-danger btn-sm" onclick="removeCircleMember('${member.MEMBERSHIP_ID}','${activeCircle.CIRCLE_ID}')">Remove</button></td>
      </tr>`).join('');
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="grid-2">
        <div>
          ${sessionHasRoleAtLeast('Galvanizer') ? `
            <div class="card" style="margin-bottom:12px">
              <div class="card-title">Circle Editor</div>
              <input type="hidden" id="circle-id" value="${activeCircle ? activeCircle.CIRCLE_ID : ''}" />
              <div class="form-group"><label class="form-label">Circle Name</label><input class="form-control" id="circle-name" value="${activeCircle ? activeCircle.CIRCLE_NAME || '' : ''}" /></div>
              <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="circle-description">${activeCircle ? activeCircle.DESCRIPTION || '' : ''}</textarea></div>
              <div style="display:flex;gap:8px">
                <button class="btn btn-primary btn-sm" onclick="saveCircleForm()">${activeCircle ? 'Update Circle' : 'Create Circle'}</button>
                <button class="btn btn-ghost btn-sm" onclick="resetCircleForm()">Clear</button>
              </div>
            </div>
          ` : ''}
          ${circleCards || '<div class="empty-state"><p>No circles found.</p></div>'}
        </div>
        <div class="card">
          <div class="card-title">${activeCircle ? activeCircle.CIRCLE_NAME : 'Circle Members'}</div>
          ${activeCircle ? `
            <div class="text-muted" style="margin-bottom:12px">${activeCircle.DESCRIPTION || ''}</div>
            ${sessionHasRoleAtLeast('Galvanizer') ? `
              <div class="form-row">
                <div class="form-group"><input class="form-control" id="circle-user-email" list="circle-user-list" placeholder="Select internal user" /><datalist id="circle-user-list">${buildUserOptions(USER_LOOKUP_CACHE.filter(user => isInternalUserRole(user.ROLE || '')), '')}</datalist></div>
                <div class="form-group"><input class="form-control" id="circle-role" placeholder="Lead / Member" /></div>
              </div>
              <button class="btn btn-primary btn-sm" onclick="addCircleMember('${activeCircle.CIRCLE_ID}')">Add Member</button>
            ` : ''}
            <div class="table-wrap" style="margin-top:16px"><table>
              <thead><tr><th>User</th><th>Circle Role</th><th></th></tr></thead>
              <tbody>${memberRows || '<tr><td colspan="3" class="text-muted" style="padding:20px;text-align:center">No members.</td></tr>'}</tbody>
            </table></div>
          ` : '<div class="text-muted">Select a circle.</div>'}
        </div>
      </div>
    </div>`;
  } catch (e) {
    document.getElementById('page-content').innerHTML = err(e.message);
  }
}

function resetCircleForm() {
  const fields = ['circle-id','circle-name','circle-description'];
  fields.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
}

function prefillCircleForm(circleJson) {
  const circle = JSON.parse(circleJson);
  showPage('circles');
  setTimeout(() => {
    document.getElementById('circle-id').value = circle.CIRCLE_ID || '';
    document.getElementById('circle-name').value = circle.CIRCLE_NAME || '';
    document.getElementById('circle-description').value = circle.DESCRIPTION || '';
  }, 50);
}

async function saveCircleForm() {
  const circleId = document.getElementById('circle-id')?.value.trim();
  const name = document.getElementById('circle-name')?.value.trim();
  const description = document.getElementById('circle-description')?.value.trim() || '';
  if (!name) return;
  await API.saveCircle({ CIRCLE_ID: circleId || undefined, CIRCLE_NAME: name, DESCRIPTION: description, STATUS: 'Active' });
  await populateCircleNav();
  loadCircles();
}

async function deactivateCircle(circleId) {
  if (!confirm(`Deactivate circle ${circleId}?`)) return;
  await API.deleteCircle(circleId);
  await populateCircleNav();
  loadCircles();
}

async function addCircleMember(circleId) {
  const userEmail = document.getElementById('circle-user-email').value.trim();
  const roleInCircle = document.getElementById('circle-role').value.trim() || 'Member';
  if (!userEmail) return;
  await API.saveCircleMember({ CIRCLE_ID: circleId, USER_EMAIL: userEmail, ROLE_IN_CIRCLE: roleInCircle });
  await populateCircleNav();
  loadCircles(JSON.stringify({ CIRCLE_ID: circleId, CIRCLE_NAME: '' }));
}

async function removeCircleMember(membershipId, circleId) {
  await API.removeCircleMember(membershipId);
  loadCircles(JSON.stringify({ CIRCLE_ID: circleId, CIRCLE_NAME: '' }));
}

async function createThread() {
  openGenericModal('New Thread', `
    <div class="form-group"><label class="form-label">Thread Title *</label><input class="form-control" id="thread-title" /></div>
    <div class="form-group"><label class="form-label">Thread Type</label><select class="form-control" id="thread-type"><option value="General">General</option><option value="Case">Case</option><option value="Client">Client</option><option value="Internal">Internal</option></select></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitThreadModal()">Create Thread</button>
    </div>`);
}

async function submitThreadModal() {
  const title = document.getElementById('thread-title').value.trim();
  const type = document.getElementById('thread-type').value || 'General';
  if (!title) return;
  try {
    await API.saveMessageThread({ TITLE: title, THREAD_TYPE: type, STATUS: 'Open' });
    closeModal('modal-generic');
    loadMessages();
  } catch (e) {
    alert('Error: ' + e.message);
  }
}

async function createDirectMessage() {
  await ensureUserLookupLoaded();
  openGenericModal('Direct Message', `
    <div class="form-group"><label class="form-label">Recipient *</label><input class="form-control" id="direct-recipient" list="direct-user-list" placeholder="Search user email" /><datalist id="direct-user-list">${buildUserOptions(USER_LOOKUP_CACHE.filter(u => String(u.EMAIL || '').toLowerCase() !== String(SESSION.email || '').toLowerCase()), '')}</datalist></div>
    <div class="form-group"><label class="form-label">Conversation Title</label><input class="form-control" id="direct-title" value="Direct conversation" /></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitDirectMessageModal()">Start Conversation</button>
    </div>`);
}

async function openDirectMessageForUser(email, name = '') {
  await ensureUserLookupLoaded();
  openGenericModal(`Message ${name || email}`, `
    <div class="form-group"><label class="form-label">Recipient</label><input class="form-control" id="direct-recipient" value="${escapeAttr(email)}" readonly /></div>
    <div class="form-group"><label class="form-label">Conversation Title</label><input class="form-control" id="direct-title" value="${escapeAttr(name ? `Direct: ${name}` : 'Direct conversation')}" /></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" onclick="submitDirectMessageModal()">Open Chat</button>
    </div>`);
}

async function submitDirectMessageModal() {
  const recipientEmail = document.getElementById('direct-recipient').value.trim();
  const title = document.getElementById('direct-title').value.trim() || 'Direct conversation';
  const result = await API.createDirectThread({ recipientEmail, title });
  closeModal('modal-generic');
  showPage('inbox');
  setTimeout(() => openInboxChat(result.threadId), 50);
}

async function openThread(threadId) {
  try {
    const messages = await API.getThreadMessages(threadId);
    const items = messages.map(m => `
      <div style="padding:10px 0;border-bottom:1px solid var(--border)">
        <div style="font-weight:600">${m.SENDER_NAME} <span class="text-muted" style="font-size:11px">${m.SENDER_ROLE}</span></div>
        <div style="font-size:13px;margin-top:4px">${m.MESSAGE_TEXT}</div>
        <div class="text-muted" style="font-size:11px;margin-top:4px">${fmt(m.CREATED_AT)}</div>
      </div>`).join('');
    document.getElementById('thread-panel').innerHTML = `
      <div class="card-title">Conversation</div>
      <div style="max-height:420px;overflow:auto">${items || '<div class="text-muted">No messages yet.</div>'}</div>
      <div class="form-group" style="margin-top:16px">
        <textarea class="form-control" id="thread-reply" rows="4" placeholder="Write a reply or internal update"></textarea>
      </div>
      <button class="btn btn-primary btn-sm" onclick="sendThreadReply('${threadId}')">Send</button>`;
  } catch (e) {
    document.getElementById('thread-panel').innerHTML = err(e.message);
  }
}

async function sendThreadReply(threadId) {
  const messageText = document.getElementById('thread-reply').value.trim();
  if (!messageText) return;
  try {
    await API.sendThreadMessage({ threadId, messageText, isInternal: sessionHasRoleAtLeast('Staff') ? 'Yes' : 'No' });
    openThread(threadId);
  } catch (e) {
    alert('Error: ' + e.message);
  }
}

async function loadNotifications() {
  destroyCharts();
  try {
    const items = await API.getNotifications();
    setTopbarActions(`<button class="btn btn-ghost btn-sm" onclick="clearAllNotifications()">Clear All</button>`);
    const rows = items.map(n => `
      <div class="card" style="margin-bottom:12px">
        <div style="display:flex;justify-content:space-between;gap:12px">
          <div>
            <div style="font-weight:600">${n.TITLE}</div>
            <div class="text-muted" style="font-size:13px;margin-top:4px">${n.BODY}</div>
            <div class="text-muted" style="font-size:11px;margin-top:6px">${fmt(n.CREATED_AT)}</div>
          </div>
          <div class="action-row">
            ${n.IS_READ === 'Yes' ? '' : `<button class="btn btn-ghost btn-sm" onclick="markNoticeRead('${n.NOTIFICATION_ID}')">Mark read</button>`}
            <button class="btn btn-danger btn-sm" onclick="deleteNotice('${n.NOTIFICATION_ID}')">Clear</button>
          </div>
        </div>
      </div>`).join('');
    document.getElementById('page-content').innerHTML = `<div class="page">${rows || '<div class="empty-state"><p>No notifications.</p></div>'}</div>`;
  } catch (e) {
    document.getElementById('page-content').innerHTML = err(e.message);
  }
}

async function markNoticeRead(notificationId) {
  await API.markNotificationRead(notificationId);
  loadNotifications();
}

async function deleteNotice(notificationId) {
  await API.deleteNotification(notificationId);
  loadNotifications();
}

async function clearAllNotifications() {
  if (!confirm('Clear all notifications?')) return;
  await API.clearNotifications();
  loadNotifications();
}

async function saveDailyPriorities() {
  await API.saveDailyPriority({
    priority1: document.getElementById('dp-1').value.trim(),
    priority2: document.getElementById('dp-2').value.trim(),
    priority3: document.getElementById('dp-3').value.trim(),
    notes: document.getElementById('dp-notes').value.trim(),
    status: 'Planned',
  });
  loadDailyAudit();
}

async function loadDailyAudit() {
  destroyCharts();
  try {
    await ensureUserLookupLoaded();
    const filters = {
      entryDate: document.getElementById('audit-date')?.value || '',
      userName: document.getElementById('audit-user')?.value || '',
    };
    const audit = await API.getDailyAudit(filters);
    const internalNames = USER_LOOKUP_CACHE.filter(user => isInternalUserRole(user.ROLE)).map(user => ({ FULL_NAME: user.FULL_NAME })).filter(user => user.FULL_NAME);
    const priorities = (audit.priorities || []).map(item => `
      <tr><td>${item.ENTRY_DATE}</td><td>${item.USER_NAME || item.USER_EMAIL}</td><td>${item.PRIORITY_1 || '—'}</td><td>${item.PRIORITY_2 || '—'}</td><td>${item.PRIORITY_3 || '—'}</td></tr>
    `).join('');
    const wrapups = (audit.wrapups || []).map(item => `
      <tr><td>${item.ENTRY_DATE}</td><td>${item.USER_NAME || item.USER_EMAIL}</td><td>${item.HIGH_POINTS || '—'}</td><td>${item.LOW_POINTS || '—'}</td><td>${item.HELP_NEEDED || '—'}</td></tr>
    `).join('');
    setTopbarActions(`<button class="btn btn-ghost btn-sm" onclick="toggleDailyAuditFilters()">${DAILY_AUDIT_FILTER_OPEN ? 'Close Filter' : 'Filter'}</button>`);
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card filter-panel filter-drawer ${DAILY_AUDIT_FILTER_OPEN ? 'open' : ''}" style="margin-bottom:16px">
        <div class="filter-head">
          <div class="card-title">Daily Audit Filters</div>
          <div class="filter-actions">
            <button class="btn btn-primary btn-sm" onclick="applyDailyAuditFilters()">Apply</button>
            <button class="btn btn-ghost btn-sm" onclick="resetDailyAuditFilters()">Reset</button>
          </div>
        </div>
        <div class="filter-grid compact-2">
          <div class="form-group"><label class="form-label">Date</label><input type="date" class="form-control" id="audit-date" value="${filters.entryDate}" /></div>
          <div class="form-group"><label class="form-label">Name</label><input class="form-control" id="audit-user" list="audit-user-list" value="${filters.userName}" placeholder="Select internal user" /><datalist id="audit-user-list">${buildNameOptions(internalNames)}</datalist></div>
        </div>
      </div>
      <div class="grid-2">
        <div class="card">
          <div class="card-title">Daily Priorities</div>
          ${isInternalUser() ? `
          <div class="form-group"><label class="form-label">Priority 1</label><input class="form-control" id="dp-1" /></div>
          <div class="form-group"><label class="form-label">Priority 2</label><input class="form-control" id="dp-2" /></div>
          <div class="form-group"><label class="form-label">Priority 3</label><input class="form-control" id="dp-3" /></div>
          <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="dp-notes"></textarea></div>
          <button class="btn btn-primary btn-sm" onclick="saveDailyPriorities()">Submit Priorities</button>
          ` : ''}
          <div class="table-wrap" style="margin-top:16px"><table>
            <thead><tr><th>Date</th><th>User</th><th>Priority 1</th><th>Priority 2</th><th>Priority 3</th></tr></thead>
            <tbody>${priorities || '<tr><td colspan="5" class="text-muted" style="padding:20px;text-align:center">No priority entries.</td></tr>'}</tbody>
          </table></div>
        </div>
        <div class="card">
          <div class="card-title">Day-End Highlights</div>
          ${isInternalUser() ? `
          <div class="form-group"><label class="form-label">High points</label><textarea class="form-control" id="wu-high" rows="3"></textarea></div>
          <div class="form-group"><label class="form-label">Low points</label><textarea class="form-control" id="wu-low" rows="3"></textarea></div>
          <div class="form-group"><label class="form-label">Help needed</label><textarea class="form-control" id="wu-help" rows="3"></textarea></div>
          <button class="btn btn-primary btn-sm" onclick="saveWrapup()">Submit Day-End Log</button>
          ` : ''}
          <div class="table-wrap" style="margin-top:16px"><table>
            <thead><tr><th>Date</th><th>User</th><th>High</th><th>Low</th><th>Help</th></tr></thead>
            <tbody>${wrapups || '<tr><td colspan="5" class="text-muted" style="padding:20px;text-align:center">No day-end logs.</td></tr>'}</tbody>
          </table></div>
        </div>
      </div>
    </div>`;
  } catch (e) {
    showPageError('loadDailyAudit', e);
  }
}

async function saveWrapup() {
  await API.saveDailyWrapup({
    highPoints: document.getElementById('wu-high').value.trim(),
    lowPoints: document.getElementById('wu-low').value.trim(),
    helpNeeded: document.getElementById('wu-help').value.trim(),
  });
  loadDailyAudit();
}

function resetDailyAuditFilters() {
  DAILY_AUDIT_FILTER_OPEN = false;
  loadDailyAudit();
}

function isInternalUserRole(role) {
  return ['Super Admin','Admin','Galvanizer','Staff','Attorney'].includes(role);
}

function applyDailyAuditFilters() {
  DAILY_AUDIT_FILTER_OPEN = false;
  loadDailyAudit();
}

function toggleDailyAuditFilters() {
  DAILY_AUDIT_FILTER_OPEN = !DAILY_AUDIT_FILTER_OPEN;
  loadDailyAudit();
}

async function loadExpenses() {
  destroyCharts();
  if (!isInternalUser()) {
    document.getElementById('page-content').innerHTML = err('Expense claims are internal only.');
    return;
  }
  try {
    const period = document.getElementById('expense-period')?.value || 'month';
    const fromDate = document.getElementById('expense-from-date')?.value || '';
    const toDate = document.getElementById('expense-to-date')?.value || '';
    const claims = await API.getExpenseClaims({ period, fromDate, toDate });
    const rows = claims.map(c => `
      <tr>
        <td>${c.CLAIM_ID}</td>
        <td>${c.CLAIM_DATE}</td>
        <td>${c.CATEGORY}</td>
        <td>${money(c.AMOUNT)}</td>
        <td>${statusBadge(c.STATUS || 'Submitted')}</td>
        <td>${c.USER_NAME || '—'}</td>
        <td>${c.BILL_LINK ? `<a href="${c.BILL_LINK}" target="_blank" rel="noopener noreferrer" class="btn btn-ghost btn-sm">Open</a>` : '—'}</td>
      </tr>`).join('');
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="grid-2">
        <div class="card">
          <div class="card-title">New Expense Claim</div>
          <div class="form-group"><label class="form-label">Claim Date</label><input type="date" class="form-control" id="ex-date" /></div>
          <div class="form-group"><label class="form-label">Category</label><input class="form-control" id="ex-cat" placeholder="Travel, filing fee, courier, office expense" /></div>
          <div class="form-group"><label class="form-label">Amount</label><input type="number" class="form-control" id="ex-amt" /></div>
          <div class="form-group"><label class="form-label">Bill Image (.jpg, .png)</label><input type="file" accept=".jpg,.jpeg,.png,image/jpeg,image/png" class="form-control" id="ex-bill-file" /></div>
          <div class="form-group"><label class="form-label">Bill link</label><input class="form-control" id="ex-bill" placeholder="Drive link" /></div>
          <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="ex-desc"></textarea></div>
          <button class="btn btn-primary" onclick="saveExpenseClaim()">Submit Claim</button>
        </div>
        <div class="card">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px">
            <div class="card-title" style="margin-bottom:0">Claim History</div>
            <div style="display:flex;align-items:center;gap:8px;flex-wrap:nowrap;justify-content:flex-end;overflow:auto">
              <select class="form-control" id="expense-period" style="max-width:140px" onchange="loadExpenses()">
                <option value="month" ${period === 'month' ? 'selected' : ''}>This Month</option>
                <option value="all" ${period === 'all' ? 'selected' : ''}>All Time</option>
              </select>
              <input type="date" class="form-control" id="expense-from-date" style="max-width:150px" value="${fromDate}" onchange="loadExpenses()" />
              <input type="date" class="form-control" id="expense-to-date" style="max-width:150px" value="${toDate}" onchange="loadExpenses()" />
            </div>
          </div>
          <div class="table-wrap"><table>
            <thead><tr><th>ID</th><th>Date</th><th>Category</th><th>Amount</th><th>Status</th><th>User</th><th>Bill</th></tr></thead>
            <tbody>${rows || '<tr><td colspan="7" class="text-muted" style="padding:20px;text-align:center">No expense claims yet.</td></tr>'}</tbody>
          </table></div>
        </div>
      </div>
    </div>`;
  } catch (e) {
    document.getElementById('page-content').innerHTML = err(e.message);
  }
}

async function loadGalvanizerQueue() {
  destroyCharts();
  if (!sessionHasAnyRole(['Super Admin', 'Admin', 'Galvanizer'])) {
    document.getElementById('page-content').innerHTML = err('Galvanizer queue is restricted.');
    return;
  }
  try {
    const filters = {
      clientCode: CASE_FILTERS.clientCode || '',
      caseId: CASE_FILTERS.caseId || '',
      workflowStage: CASE_FILTERS.workflowStage || '',
      query: CASE_FILTERS.query || '',
    };
    const [cases, commandCenter] = await Promise.all([
      API.getGalvanizerQueue(filters),
      API.getGalvanizerCommandCenter(filters)
    ]);
    const rows = cases.map(c => `
      <tr>
        <td>${c.CASE_ID}</td>
        <td>${c.CLIENT_ID}</td>
        <td>${c.PATENT_TITLE}</td>
        <td>${c.ASSIGNED_STAFF_EMAIL || '—'}</td>
        <td>${c.GALVANIZER_EMAIL || '—'}</td>
        <td>${c.ATTORNEY || '—'}</td>
        <td>${statusBadge(c.WORKFLOW_STAGE || 'Drafting')}</td>
        <td><button class="btn btn-ghost btn-sm" onclick='openCaseModal(${JSON.stringify(JSON.stringify(c))})'>Open</button></td>
      </tr>`).join('');
    document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="stats-grid">
        <div class="stat-card accent-primary"><div class="stat-label">Incoming</div><div class="stat-value">${(commandCenter.incoming || []).length}</div></div>
        <div class="stat-card accent-sky"><div class="stat-label">Ready For Attorney</div><div class="stat-value">${(commandCenter.readyForAttorney || []).length}</div></div>
        <div class="stat-card accent-amber"><div class="stat-label">Under Review</div><div class="stat-value">${(commandCenter.underReview || []).length}</div></div>
        <div class="stat-card accent-rose"><div class="stat-label">Pending Approvals</div><div class="stat-value">${(commandCenter.pendingApprovals || []).length}</div></div>
      </div>
      ${renderCaseFilters()}
      <div class="card">
        <div class="card-title">Galvanizer Work Queue</div>
        <div class="table-wrap"><table>
          <thead><tr><th>Case ID</th><th>Client Code</th><th>Title</th><th>Assigned Staff</th><th>Galvanizer</th><th>Attorney</th><th>Stage</th><th></th></tr></thead>
          <tbody>${rows || '<tr><td colspan="8" class="text-muted" style="padding:20px;text-align:center">No cases in queue.</td></tr>'}</tbody>
        </table></div>
      </div>
    </div>`;
  } catch (e) {
    showPageError('loadGalvanizerQueue', e);
  }
}

async function saveExpenseClaim() {
  let billLink = document.getElementById('ex-bill').value.trim();
  const file = document.getElementById('ex-bill-file')?.files?.[0];
  if (file) {
    const dataUrl = await fileToDataUrl(file);
    const upload = await API.uploadExpenseBill({
      fileName: file.name,
      mimeType: file.type,
      dataUrl
    });
    billLink = upload.fileUrl || billLink;
  }
  await API.submitExpenseClaim({
    claimDate: document.getElementById('ex-date').value,
    category: document.getElementById('ex-cat').value.trim(),
    amount: document.getElementById('ex-amt').value,
    billLink,
    description: document.getElementById('ex-desc').value.trim(),
  });
  loadExpenses();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  CONTACT
// ═══════════════════════════════════════════════════════════════════════════════
function loadContact() {
  document.getElementById('page-content').innerHTML = `
  <div class="page">
    <div class="card" style="max-width:560px">
      <h3 style="margin-bottom:20px">Send a Message</h3>
      <div class="form-group"><label class="form-label">Subject</label><input class="form-control" id="ct-subject" placeholder="What is this regarding?" /></div>
      <div class="form-group"><label class="form-label">Related Case ID (optional)</label><input class="form-control" id="ct-case" placeholder="e.g. 157M001" /></div>
      <div class="form-group"><label class="form-label">Message</label><textarea class="form-control" id="ct-message" rows="5" placeholder="Type your message…"></textarea></div>
      <button class="btn btn-primary" onclick="submitContact()">Send Message</button>
      <div id="ct-status" style="margin-top:12px"></div>
    </div>
  </div>`;
}

async function submitContact() {
  const subject = document.getElementById('ct-subject').value.trim();
  const caseId = document.getElementById('ct-case').value.trim();
  const message = document.getElementById('ct-message').value.trim();
  const status = document.getElementById('ct-status');
  if (!subject || !message) { status.innerHTML = '<div class="alert alert-error">Please fill subject and message.</div>'; return; }
  try {
    const r = await API.submitContact(subject, caseId, message);
    status.innerHTML = `<div class="alert alert-success">${r.message}</div>`;
    document.getElementById('ct-subject').value = '';
    document.getElementById('ct-case').value = '';
    document.getElementById('ct-message').value = '';
  } catch(e) { status.innerHTML = `<div class="alert alert-error">${e.message}</div>`; }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  MODALS — Case
// ═══════════════════════════════════════════════════════════════════════════════
function clientSearchBlock(selectedCode) {
  const listId = 'client-code-options';
  return `
    <div class="form-group">
      <label class="form-label">Client Code *</label>
      <input class="form-control" id="mc-client-code" list="${listId}" value="${selectedCode || ''}" placeholder="Search client code like A61M or 807Y" oninput="syncClientCodeSelection()" ${selectedCode ? 'readonly' : ''} />
      <datalist id="${listId}">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
      <div class="text-muted" id="mc-client-help" style="margin-top:6px;font-size:12px"></div>
    </div>`;
}

function syncClientCodeSelection() {
  const input = document.getElementById('mc-client-code');
  const help = document.getElementById('mc-client-help');
  const client = resolveClientByCode(input.value);
  if (client) {
    help.textContent = `${client.CLIENT_NAME || ''} | ${client.CLIENT_TYPE || 'Client'} | ${client.CLIENT_REGION || ''}`;
  } else {
    help.textContent = 'Select a valid client code from the list.';
  }
}

async function openCaseModal(jsonStr) {
  await ensureOrgLookupLoaded();
  const c = jsonStr ? JSON.parse(jsonStr) : null;
  const staffOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Staff');
  const galvanizerOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Galvanizer');
  const attorneyOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Attorney');
  const orgOptions = buildOrgOptions(ORG_LOOKUP_CACHE);
  document.getElementById('modal-case-title').textContent = c ? 'Edit Case' : 'New Case';
  document.getElementById('modal-case-body').innerHTML = `
    ${clientSearchBlock(c?.CLIENT_ID || '')}
    <div class="form-group"><label class="form-label">Patent Title *</label><input class="form-control" id="mc-title" value="${c?.PATENT_TITLE||''}" /></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Application Number</label><input class="form-control" id="mc-appno" value="${c?.APPLICATION_NUMBER||''}" /></div>
      <div class="form-group"><label class="form-label">Country</label><select class="form-control" id="mc-country">${['India','USA','EPO','PCT','China','Japan','Other'].map(x=>`<option${(c?.COUNTRY||'India')===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Filing Date</label><input type="date" class="form-control" id="mc-filed" value="${toInputDate(c?.FILING_DATE)}" /></div>
      <div class="form-group"><label class="form-label">Next Deadline</label><input type="date" class="form-control" id="mc-deadline" value="${toInputDate(c?.NEXT_DEADLINE)}" /></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Status</label><select class="form-control" id="mc-status">${['Drafted','Filed','Published','Under Examination','Granted','Abandoned','Lapsed','Refused'].map(x=>`<option${c?.CURRENT_STATUS===x?' selected':''}>${x}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Patent Type</label><select class="form-control" id="mc-type">${['Utility','Design','PCT','Provisional'].map(x=>`<option${c?.PATENT_TYPE===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Assigned Staff</label><input class="form-control" id="mc-staff" list="mc-staff-list" value="${c?.ASSIGNED_STAFF_EMAIL||''}" placeholder="Search staff" /><datalist id="mc-staff-list">${staffOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Galvanizer</label><input class="form-control" id="mc-galvanizer" list="mc-galvanizer-list" value="${c?.GALVANIZER_EMAIL||''}" placeholder="Search galvanizer" /><datalist id="mc-galvanizer-list">${galvanizerOptions}</datalist></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Attorney</label><input class="form-control" id="mc-atty" list="mc-atty-list" value="${c?.ATTORNEY||''}" placeholder="Search attorney" /><datalist id="mc-atty-list">${attorneyOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Workflow Stage</label><select class="form-control" id="mc-stage">${['Drafting','Ready for Attorney','Under Attorney Review','Filed'].map(x=>`<option${(c?.WORKFLOW_STAGE||'Drafting')===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Organization ID</label><input class="form-control" id="mc-org" list="mc-org-list" value="${c?.ORG_ID || ''}" placeholder="Search organization" /><datalist id="mc-org-list">${orgOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Priority</label><select class="form-control" id="mc-priority">${['Normal','High','Urgent'].map(x=>`<option${(c?.PRIORITY||'Normal')===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Case ID Preview</label><input class="form-control" id="mc-preview" value="${c?.CASE_ID||''}" readonly /></div>
    </div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="mc-notes">${c?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-case')">Cancel</button>
      <button class="btn btn-primary" id="mc-save" onclick="saveCaseV2('${c?.CASE_ID||''}')">Save Case</button>
    </div>`;
  openModal('modal-case');
  syncClientCodeSelection();
}

async function saveCase(existingId) {
  const btn = document.getElementById('mc-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const clientCode = document.getElementById('mc-client-code').value.trim().toUpperCase();
  const client = resolveClientByCode(clientCode);
  if (!client) {
    btn.disabled = false; btn.textContent = 'Save Case';
    alert('Select a valid client code.');
    return;
  }
  const data = {
    CASE_ID: existingId || undefined,
    CLIENT_ID: client.CLIENT_ID || clientCode,
    PATENT_TITLE: document.getElementById('mc-title').value.trim(),
    APPLICATION_NUMBER: document.getElementById('mc-appno').value.trim(),
    COUNTRY: document.getElementById('mc-country').value,
    FILING_DATE: document.getElementById('mc-filed').value,
    NEXT_DEADLINE: document.getElementById('mc-deadline').value,
    CURRENT_STATUS: document.getElementById('mc-status').value,
    PATENT_TYPE: document.getElementById('mc-type').value,
    ASSIGNED_STAFF_EMAIL: document.getElementById('mc-staff').value.trim(),
    GALVANIZER_EMAIL: document.getElementById('mc-galvanizer').value.trim(),
    ATTORNEY: document.getElementById('mc-atty').value.trim(),
    WORKFLOW_STAGE: document.getElementById('mc-stage').value,
    PRIORITY: document.getElementById('mc-priority').value,
    ORG_ID: client.ORG_ID || '',
    NOTES: document.getElementById('mc-notes').value.trim(),
  };
  if (!existingId) delete data.CASE_ID;
  try {
    const r = await API.saveCase(data);
    if (r.success) { closeModal('modal-case'); loadCases(); loadDashboard(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    showPageError('saveCase', e);
    btn.disabled = false; btn.textContent = 'Save Case';
  }
}

async function saveCaseV2(existingId) {
  const btn = document.getElementById('mc-save');
  btn.disabled = true; btn.textContent = 'Saving...';
  const clientCode = document.getElementById('mc-client-code').value.trim().toUpperCase();
  const client = resolveClientByCode(clientCode);
  const org = resolveOrganization(document.getElementById('mc-org')?.value.trim() || '');
  if (!client) {
    btn.disabled = false; btn.textContent = 'Save Case';
    alert('Select a valid client code.');
    return;
  }
  const data = {
    CASE_ID: existingId || undefined,
    CLIENT_ID: client.CLIENT_ID || clientCode,
    PATENT_TITLE: document.getElementById('mc-title').value.trim(),
    APPLICATION_NUMBER: document.getElementById('mc-appno').value.trim(),
    COUNTRY: document.getElementById('mc-country').value,
    FILING_DATE: document.getElementById('mc-filed').value,
    NEXT_DEADLINE: document.getElementById('mc-deadline').value,
    CURRENT_STATUS: document.getElementById('mc-status').value,
    PATENT_TYPE: document.getElementById('mc-type').value,
    ASSIGNED_STAFF_EMAIL: document.getElementById('mc-staff').value.trim(),
    GALVANIZER_EMAIL: document.getElementById('mc-galvanizer').value.trim(),
    ATTORNEY: document.getElementById('mc-atty').value.trim(),
    WORKFLOW_STAGE: document.getElementById('mc-stage').value,
    PRIORITY: document.getElementById('mc-priority').value,
    ORG_ID: org?.ORG_ID || client.ORG_ID || '',
    NOTES: document.getElementById('mc-notes').value.trim(),
  };
  if (!existingId) delete data.CASE_ID;
  try {
    const r = await API.saveCase(data);
    if (r.success) { closeModal('modal-case'); loadCases(); loadDashboard(); }
    else throw new Error(r.error || r.message);
  } catch (e) {
    showPageError('saveCaseV2', e);
    btn.disabled = false; btn.textContent = 'Save Case';
  }
}

async function openClientModal(jsonStr) {
  await ensureOrgLookupLoaded();
  await ensureUserLookupLoaded();
  const c = jsonStr ? JSON.parse(jsonStr) : null;
  const region = c?.CLIENT_REGION || 'India';
  const rawCode = String(c?.CLIENT_CODE || '').replace(/[MY]$/,'');
  const orgOptions = buildOrgOptions(ORG_LOOKUP_CACHE);
  const clientAdminOptions = buildUserIdentityOptions(USER_LOOKUP_CACHE, user => ['Client Admin','Client Employee'].includes(user.ROLE || ''));
  const assignedStaffOptions = buildUserNameEmailOptions(USER_LOOKUP_CACHE, user => getEffectiveRoleList(user).some(role => ['Super Admin','Admin','Galvanizer','Staff'].includes(role)));
  document.getElementById('modal-client-title').textContent = c ? 'Edit Client' : 'New Client';
  document.getElementById('modal-client-body').innerHTML = `
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Type</label><select class="form-control" id="cl-client-type" onchange="refreshClientTypeFields()">${['Individual','Organization'].map(x=>`<option${(c?.CLIENT_TYPE||'Individual')===x?' selected':''}>${x}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Region</label><select class="form-control" id="cl-region" onchange="updateClientCodePreview()">${['India','Abroad'].map(x=>`<option${region===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Base Code *</label><input class="form-control" id="cl-code-base" value="${rawCode}" placeholder="A61 or 870" oninput="updateClientCodePreview()" /></div>
      <div class="form-group"><label class="form-label">Final Client Code</label><input class="form-control" id="cl-code" value="${c?.CLIENT_CODE||''}" readonly /></div>
    </div>
    <div class="form-row client-org-fields">
      <div class="form-group"><label class="form-label">Organization ID</label><input class="form-control" id="cl-org-id" list="cl-org-list" value="${c?.ORG_ID||''}" placeholder="Search organization" /><datalist id="cl-org-list">${orgOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Assigned Staff Email</label><input class="form-control" id="cl-staff" list="cl-staff-list" value="${c?.ASSIGNED_STAFF_EMAIL||''}" placeholder="Search staff email" /><datalist id="cl-staff-list">${assignedStaffOptions}</datalist></div>
    </div>
    <div class="form-group"><label class="form-label">Client Name *</label><input class="form-control" id="cl-name" value="${c?.CLIENT_NAME||''}" /></div>
    <div class="form-group"><label class="form-label">Contact Person</label><input class="form-control" id="cl-contact" value="${c?.CONTACT_PERSON||''}" /></div>
    <div class="form-group"><label class="form-label">Email</label><input type="email" class="form-control" id="cl-email" value="${c?.EMAIL||''}" /></div>
    <div class="form-group"><label class="form-label">Phone</label><input class="form-control" id="cl-phone" value="${c?.PHONE||''}" /></div>
    <div class="form-group client-org-fields"><label class="form-label">Client Admin User ID</label><input class="form-control" id="cl-admin-user" list="cl-admin-list" value="${c?.CLIENT_ADMIN_USER_ID||''}" placeholder="Search user id" /><datalist id="cl-admin-list">${clientAdminOptions}</datalist></div>
    <div class="form-group"><label class="form-label">Address</label><textarea class="form-control" id="cl-addr">${c?.ADDRESS||''}</textarea></div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="cl-notes">${c?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-client')">Cancel</button>
      <button class="btn btn-primary" id="cl-save" onclick="saveClientV2('${c?.CLIENT_ID||''}')">Save Client</button>
    </div>`;
  openModal('modal-client');
  updateClientCodePreview();
  refreshClientTypeFields();
}

function updateClientCodePreview() {
  const base = String(document.getElementById('cl-code-base')?.value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/[MY]$/,'');
  const region = document.getElementById('cl-region')?.value || 'India';
  const suffix = region === 'Abroad' ? 'Y' : 'M';
  const finalCode = base ? `${base}${suffix}` : '';
  const out = document.getElementById('cl-code');
  if (out) out.value = finalCode;
}

function refreshClientTypeFields() {
  const type = document.getElementById('cl-client-type')?.value || 'Individual';
  document.querySelectorAll('.client-org-fields').forEach(el => {
    el.style.display = type === 'Organization' ? '' : 'none';
  });
}

async function saveClient(existingId) {
  const btn = document.getElementById('cl-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    CLIENT_ID: existingId || undefined,
    CLIENT_NAME: document.getElementById('cl-name').value.trim(),
    CONTACT_PERSON: document.getElementById('cl-contact').value.trim(),
    EMAIL: document.getElementById('cl-email').value.trim(),
    PHONE: document.getElementById('cl-phone').value.trim(),
    CLIENT_TYPE: document.getElementById('cl-client-type').value,
    CLIENT_REGION: document.getElementById('cl-region').value,
    CLIENT_CODE: document.getElementById('cl-code').value.trim().toUpperCase(),
    ORG_ID: document.getElementById('cl-org-id').value.trim(),
    CLIENT_ADMIN_USER_ID: document.getElementById('cl-admin-user').value.trim(),
    ASSIGNED_STAFF_EMAIL: document.getElementById('cl-staff').value.trim(),
    ADDRESS: document.getElementById('cl-addr').value.trim(),
    NOTES: document.getElementById('cl-notes').value.trim(),
  };
  if (!/^[A-Z0-9]{3,4}[MY]$/.test(data.CLIENT_CODE)) {
    btn.disabled = false; btn.textContent = 'Save Client';
    alert('Enter a valid client base code. Final code should look like A61M or 870Y.');
    return;
  }
  if (!existingId) delete data.CLIENT_ID;
  try {
    const r = await API.saveClient(data);
    if (r.success) { closeModal('modal-client'); CLIENT_LOOKUP_CACHE = []; await ensureClientLookupLoaded(); loadMgmtClients(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    showPageError('saveClient', e);
    btn.disabled = false; btn.textContent = 'Save Client';
  }
}

async function saveClientV2(existingId) {
  const btn = document.getElementById('cl-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const clientType = document.getElementById('cl-client-type').value;
  const org = clientType === 'Organization' ? resolveOrganization(document.getElementById('cl-org-id').value.trim()) : null;
  const adminUser = clientType === 'Organization' ? resolveUserById(document.getElementById('cl-admin-user').value.trim()) : null;
  const data = {
    CLIENT_ID: existingId || undefined,
    CLIENT_NAME: document.getElementById('cl-name').value.trim(),
    CONTACT_PERSON: document.getElementById('cl-contact').value.trim(),
    EMAIL: document.getElementById('cl-email').value.trim(),
    PHONE: document.getElementById('cl-phone').value.trim(),
    CLIENT_TYPE: clientType,
    CLIENT_REGION: document.getElementById('cl-region').value,
    CLIENT_CODE: document.getElementById('cl-code').value.trim().toUpperCase(),
    ORG_ID: org?.ORG_ID || '',
    CLIENT_ADMIN_USER_ID: adminUser?.USER_ID || '',
    ASSIGNED_STAFF_EMAIL: document.getElementById('cl-staff').value.trim(),
    ADDRESS: document.getElementById('cl-addr').value.trim(),
    NOTES: document.getElementById('cl-notes').value.trim(),
  };
  if (!/^[A-Z0-9]{3,4}[MY]$/.test(data.CLIENT_CODE)) {
    btn.disabled = false; btn.textContent = 'Save Client';
    alert('Enter a valid client base code. Final code should look like A61M or 870Y.');
    return;
  }
  if (!existingId) delete data.CLIENT_ID;
  try {
    const r = await API.saveClient(data);
    if (r.success) {
      closeModal('modal-client');
      CLIENT_LOOKUP_CACHE = [];
      await ensureClientLookupLoaded();
      await ensureOrgLookupLoaded(true);
      loadMgmtClients();
    } else {
      throw new Error(r.error || r.message);
    }
  } catch (e) {
    showPageError('saveClientV2', e);
    btn.disabled = false; btn.textContent = 'Save Client';
  }
}

// ── Client Modal ──────────────────────────────────────────────────────────────
function openClientModalLegacy(jsonStr) {
  const c = jsonStr ? JSON.parse(jsonStr) : null;
  document.getElementById('modal-client-title').textContent = c ? 'Edit Client' : 'New Client';
  document.getElementById('modal-client-body').innerHTML = `
    ${!c ? `<div class="form-group"><label class="form-label">Client Type *</label>
      <div style="display:flex;gap:8px;margin-top:4px">
        <button class="btn btn-primary btn-sm" id="type-M" onclick="selType('M')">M — Indian</button>
        <button class="btn btn-ghost btn-sm" id="type-Y" onclick="selType('Y')">Y — Abroad</button>
      </div>
      <input type="hidden" id="cl-type" value="false" />
    </div>` : ''}
    <div class="form-group"><label class="form-label">Client Name *</label><input class="form-control" id="cl-name" value="${c?.CLIENT_NAME||''}" /></div>
    <div class="form-group"><label class="form-label">Contact Person</label><input class="form-control" id="cl-contact" value="${c?.CONTACT_PERSON||''}" /></div>
    <div class="form-group"><label class="form-label">Email</label><input type="email" class="form-control" id="cl-email" value="${c?.EMAIL||''}" /></div>
    <div class="form-group"><label class="form-label">Phone</label><input class="form-control" id="cl-phone" value="${c?.PHONE||''}" /></div>
    <div class="form-group"><label class="form-label">Address</label><textarea class="form-control" id="cl-addr">${c?.ADDRESS||''}</textarea></div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="cl-notes">${c?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-client')">Cancel</button>
      <button class="btn btn-primary" id="cl-save" onclick="saveClient('${c?.CLIENT_ID||''}')">Save Client</button>
    </div>`;
  openModal('modal-client');
}

function selTypeLegacy(t) {
  document.getElementById('cl-type').value = t === 'Y' ? 'true' : 'false';
  document.getElementById('type-M')?.classList.toggle('btn-primary', t==='M');
  document.getElementById('type-M')?.classList.toggle('btn-ghost', t!=='M');
  document.getElementById('type-Y')?.classList.toggle('btn-primary', t==='Y');
  document.getElementById('type-Y')?.classList.toggle('btn-ghost', t!=='Y');
}

async function saveClientLegacy(existingId) {
  const btn = document.getElementById('cl-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    CLIENT_ID: existingId || undefined,
    CLIENT_NAME: document.getElementById('cl-name').value.trim(),
    CONTACT_PERSON: document.getElementById('cl-contact').value.trim(),
    EMAIL: document.getElementById('cl-email').value.trim(),
    PHONE: document.getElementById('cl-phone').value.trim(),
    ADDRESS: document.getElementById('cl-addr').value.trim(),
    NOTES: document.getElementById('cl-notes').value.trim(),
    IS_ABROAD: document.getElementById('cl-type')?.value || 'false',
  };
  if (!existingId) delete data.CLIENT_ID;
  try {
    const r = await API.saveClient(data);
    if (r.success) { closeModal('modal-client'); loadMgmtClients(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save Client';
  }
}

// ── User Modal ────────────────────────────────────────────────────────────────
function openUserModal(jsonStr, options = {}) {
  const u = jsonStr ? JSON.parse(jsonStr) : null;
  const prefilledOrg = options.prefilledOrgId ? resolveOrganization(options.prefilledOrgId) : null;
  document.getElementById('modal-user-title').textContent = u ? 'Edit User' : 'New User';
  const clientOptions = buildClientOptions(CLIENT_LOOKUP_CACHE);
  const circleOptions = buildCircleOptions(CIRCLE_LOOKUP_CACHE);
  const orgOptions = buildOrgOptions(ORG_LOOKUP_CACHE);
  document.getElementById('modal-user-body').innerHTML = `
    <div class="form-group"><label class="form-label">Full Name *</label><input class="form-control" id="mu-name" value="${u?.FULL_NAME||''}" /></div>
    <div class="form-group"><label class="form-label">Email *</label><input type="email" class="form-control" id="mu-email" value="${u?.EMAIL||''}" /></div>
    <div class="form-group"><label class="form-label">Password ${u?'(leave blank to keep)':'*'}</label><input type="password" class="form-control" id="mu-pwd" /></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Role</label>
        <select class="form-control" id="mu-role" onchange="syncUserOrgSelection()">${['Super Admin','Admin','Galvanizer','Staff','Attorney','Client Admin','Client Employee','Individual Client'].map(x=>`<option${(u?.ROLE || (prefilledOrg ? 'Client Employee' : ''))===x?' selected':''}>${x}</option>`).join('')}</select>
      </div>
      <div class="form-group"><label class="form-label">Client Code</label><input class="form-control" id="mu-cid" list="mu-client-list" value="${u?.CLIENT_ID||''}" /><datalist id="mu-client-list">${clientOptions}</datalist></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Organization ID</label><input class="form-control" id="mu-org" list="mu-org-list" value="${u?.ORG_ID || prefilledOrg?.ORG_ID || ''}" placeholder="Search organization" onchange="syncUserOrgSelection()" /><datalist id="mu-org-list">${orgOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Can View Finance</label>
        <select class="form-control" id="mu-finance">${['No','Yes'].map(x=>`<option${(u?.CAN_VIEW_FINANCE||'No')===x?' selected':''}>${x}</option>`).join('')}</select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Department</label><select class="form-control" id="mu-dept">${['Management','External'].map(x=>`<option${(u?.DEPARTMENT || (prefilledOrg ? 'External' : 'Management'))===x?' selected':''}>${x}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Reports To</label><input class="form-control" id="mu-reports" value="${u?.REPORTS_TO||''}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Primary Circle</label><input class="form-control" id="mu-circle" list="mu-circle-list" value="" placeholder="Assign circle if needed" /><datalist id="mu-circle-list">${circleOptions}</datalist></div>
    <div class="form-group"><label class="form-label">Additional Roles</label><div>${renderRoleCheckboxes(u?.ADDITIONAL_ROLES||'')}</div></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-user')">Cancel</button>
      <button class="btn btn-primary" id="mu-save" onclick="saveUserV2('${u?.USER_ID||''}')">Save User</button>
    </div>`;
  openModal('modal-user');
  syncUserOrgSelection();
}

function syncUserOrgSelection() {
  const orgInput = document.getElementById('mu-org');
  const roleInput = document.getElementById('mu-role');
  const deptInput = document.getElementById('mu-dept');
  const financeInput = document.getElementById('mu-finance');
  const hasOrg = !!resolveOrganization(orgInput?.value || '');
  if (!hasOrg) return;
  if (roleInput && roleInput.value !== 'Client Admin') roleInput.value = 'Client Employee';
  if (deptInput) deptInput.value = 'External';
  if (financeInput) financeInput.value = 'No';
}

async function saveUser(existingId) {
  const btn = document.getElementById('mu-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    USER_ID: existingId || undefined,
    FULL_NAME: document.getElementById('mu-name').value.trim(),
    EMAIL: document.getElementById('mu-email').value.trim(),
    PASSWORD: document.getElementById('mu-pwd').value,
    ROLE: document.getElementById('mu-role').value,
    CLIENT_ID: document.getElementById('mu-cid').value.trim(),
    ORG_ID: document.getElementById('mu-org').value.trim(),
    DEPARTMENT: document.getElementById('mu-dept').value.trim(),
    CAN_VIEW_FINANCE: document.getElementById('mu-finance').value,
    REPORTS_TO: document.getElementById('mu-reports').value.trim(),
    ADDITIONAL_ROLES: getSelectedAdditionalRoles(),
  };
  const primaryCircle = document.getElementById('mu-circle')?.value.trim();
  if (!existingId) delete data.USER_ID;
  try {
    const r = await API.saveUser(data);
    if (r.success && primaryCircle) {
      await API.saveCircleMember({ CIRCLE_ID: primaryCircle, USER_EMAIL: data.EMAIL, ROLE_IN_CIRCLE: 'Member' });
      await populateCircleNav();
    }
    if (r.success) { closeModal('modal-user'); ensureUserLookupLoaded(true); loadMgmtUsers(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save User';
  }
}

async function saveUserV2(existingId) {
  const btn = document.getElementById('mu-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const org = resolveOrganization(document.getElementById('mu-org').value.trim());
  const selectedRole = document.getElementById('mu-role').value;
  const data = {
    USER_ID: existingId || undefined,
    FULL_NAME: document.getElementById('mu-name').value.trim(),
    EMAIL: document.getElementById('mu-email').value.trim(),
    PASSWORD: document.getElementById('mu-pwd').value,
    ROLE: org ? (selectedRole === 'Client Admin' ? 'Client Admin' : 'Client Employee') : selectedRole,
    CLIENT_ID: document.getElementById('mu-cid').value.trim(),
    ORG_ID: org?.ORG_ID || '',
    DEPARTMENT: org ? 'External' : document.getElementById('mu-dept').value.trim(),
    CAN_VIEW_FINANCE: org ? 'No' : document.getElementById('mu-finance').value,
    REPORTS_TO: document.getElementById('mu-reports').value.trim(),
    ADDITIONAL_ROLES: getSelectedAdditionalRoles(),
  };
  const primaryCircle = document.getElementById('mu-circle')?.value.trim();
  if (!existingId) delete data.USER_ID;
  try {
    const r = await API.saveUser(data);
    if (r.success && primaryCircle) {
      await API.saveCircleMember({ CIRCLE_ID: primaryCircle, USER_EMAIL: data.EMAIL, ROLE_IN_CIRCLE: 'Member' });
      await populateCircleNav();
    }
    if (r.success) {
      closeModal('modal-user');
      ensureUserLookupLoaded(true);
      ensureOrgLookupLoaded(true);
      if (_mgmtTab === 'organizations' && org?.ORG_ID) loadMgmtOrganizationsV2(org.ORG_ID);
      else loadMgmtUsers();
    } else {
      throw new Error(r.error || r.message);
    }
  } catch (e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save User';
  }
}

// ── Invoice Modal ─────────────────────────────────────────────────────────────
function openInvoiceModal(jsonStr) {
  const inv = jsonStr ? JSON.parse(jsonStr) : null;
  document.getElementById('modal-invoice-title').textContent = inv ? 'Edit Invoice' : 'New Invoice';
  document.getElementById('modal-invoice-body').innerHTML = `
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client ID *</label><input class="form-control" id="mi-cid" value="${inv?.CLIENT_ID||''}" /></div>
      <div class="form-group"><label class="form-label">Case ID</label><input class="form-control" id="mi-case" value="${inv?.CASE_ID||''}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Service Type</label>
      <select class="form-control" id="mi-svc">${['Patent Filing','Patent Prosecution','Patent Search','Legal Opinion','Drafting Fees','Government Fees','Consultation','Annual Maintenance','Other'].map(x=>`<option${inv?.SERVICE_TYPE===x?' selected':''}>${x}</option>`).join('')}</select>
    </div>
    <div class="form-group"><label class="form-label">Description</label><textarea class="form-control" id="mi-desc">${inv?.DESCRIPTION||''}</textarea></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Amount (₹) *</label><input type="number" class="form-control" id="mi-amt" value="${inv?.AMOUNT||''}" /></div>
      <div class="form-group"><label class="form-label">GST Rate (%)</label><input type="number" class="form-control" id="mi-gst" value="${inv?.GST_RATE||18}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="mi-notes">${inv?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-invoice')">Cancel</button>
      <button class="btn btn-primary" id="mi-save" onclick="saveInvoice('${inv?.INVOICE_ID||''}')">Save Invoice</button>
    </div>`;
  openModal('modal-invoice');
}

async function saveInvoice(existingId) {
  const btn = document.getElementById('mi-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    INVOICE_ID: existingId || undefined,
    CLIENT_ID: document.getElementById('mi-cid').value.trim(),
    CASE_ID: document.getElementById('mi-case').value.trim(),
    SERVICE_TYPE: document.getElementById('mi-svc').value,
    DESCRIPTION: document.getElementById('mi-desc').value.trim(),
    AMOUNT: document.getElementById('mi-amt').value,
    GST_RATE: document.getElementById('mi-gst').value,
    NOTES: document.getElementById('mi-notes').value.trim(),
  };
  if (!existingId) delete data.INVOICE_ID;
  try {
    const r = await API.saveInvoice(data);
    if (r.success) { closeModal('modal-invoice'); loadFinance(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save Invoice';
  }
}

function openBulkCaseUpdate() {
  const selected = Array.from(document.querySelectorAll('.case-bulk-check:checked')).map(el => el.value);
  if (!selected.length) {
    alert('Select at least one case.');
    return;
  }
  document.getElementById('modal-case-title').textContent = 'Bulk Update Cases';
  document.getElementById('modal-case-body').innerHTML = `
    <div class="alert alert-success">Selected ${selected.length} case(s)</div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Assigned Staff</label><input class="form-control" id="bulk-staff" type="email" /></div>
      <div class="form-group"><label class="form-label">Galvanizer</label><input class="form-control" id="bulk-galvanizer" type="email" /></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Attorney</label><input class="form-control" id="bulk-attorney" type="email" /></div>
      <div class="form-group"><label class="form-label">Workflow Stage</label><select class="form-control" id="bulk-stage">${['','Drafting','Ready for Attorney','Under Attorney Review','Filed'].map(x=>`<option value="${x}">${x || 'No change'}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Status</label><select class="form-control" id="bulk-status">${['','Drafted','Filed','Published','Under Examination','Granted','Abandoned','Lapsed','Refused'].map(x=>`<option value="${x}">${x || 'No change'}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Priority</label><select class="form-control" id="bulk-priority">${['','Normal','High','Urgent'].map(x=>`<option value="${x}">${x || 'No change'}</option>`).join('')}</select></div>
    </div>
    <div class="form-group"><label class="form-label">Next Deadline</label><input type="date" class="form-control" id="bulk-deadline" /></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-case')">Cancel</button>
      <button class="btn btn-primary" onclick='submitBulkCaseUpdate(${JSON.stringify(JSON.stringify(selected))})'>Update Cases</button>
    </div>`;
  openModal('modal-case');
}

async function submitBulkCaseUpdate(caseIdsJson) {
  const caseIds = JSON.parse(caseIdsJson);
  const updates = {
    ASSIGNED_STAFF_EMAIL: document.getElementById('bulk-staff').value.trim(),
    GALVANIZER_EMAIL: document.getElementById('bulk-galvanizer').value.trim(),
    ATTORNEY: document.getElementById('bulk-attorney').value.trim(),
    WORKFLOW_STAGE: document.getElementById('bulk-stage').value,
    CURRENT_STATUS: document.getElementById('bulk-status').value,
    PRIORITY: document.getElementById('bulk-priority').value,
    NEXT_DEADLINE: document.getElementById('bulk-deadline').value,
  };
  try {
    const result = await API.bulkUpdateCases({ caseIds, updates });
    closeModal('modal-case');
    alert(result.message || 'Cases updated.');
    if (_mgmtTab === 'cases') loadMgmtCases();
    else loadCases();
  } catch (e) {
    alert('Error: ' + e.message);
  }
}

async function openBulkDocketTrakImport() {
  await ensureClientLookupLoaded();
  CASE_IMPORT_STATE = { parsedRows: [], importableRows: [], skippedRows: [], fileName: '' };
  openGenericModal('Bulk Import DocketTrak', `
    <div class="alert alert-success">Select an existing client code first. Every imported row from the Excel file will be created under that selected client.</div>
    <div class="form-group">
      <label class="form-label">Client Code *</label>
      <input class="form-control" id="bulk-import-client" list="bulk-import-client-list" placeholder="Search client code" />
      <datalist id="bulk-import-client-list">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist>
    </div>
    <div class="form-group">
      <label class="form-label">DocketTrak Excel File (.xlsx) *</label>
      <input class="form-control" id="bulk-import-file" type="file" accept=".xlsx,.xls" onchange="previewBulkImportFile()" />
    </div>
    <div id="bulk-import-summary" class="text-muted" style="font-size:12px;margin-bottom:12px">Upload a file to preview how many Patent and Trademark rows can be imported.</div>
    <div class="progress-track"><div class="progress-bar" id="bulk-import-progress-bar" style="width:0%"></div></div>
    <div id="bulk-import-progress-label" class="text-muted" style="margin-top:8px;font-size:12px">Waiting to start import.</div>
    <div id="bulk-import-progress-details" class="text-muted" style="margin-top:6px;font-size:12px"></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-generic')">Cancel</button>
      <button class="btn btn-primary" id="bulk-import-submit" onclick="submitBulkDocketTrakImport()">Start Import</button>
    </div>`);
}

async function previewBulkImportFile() {
  const input = document.getElementById('bulk-import-file');
  const file = input?.files?.[0];
  if (!file) return;
  try {
    updateBulkImportProgress(10, 'Reading Excel file...');
    const parsed = await parseDocketTrakWorkbook(file);
    CASE_IMPORT_STATE = parsed;
    const details = `
      <div class="import-summary-grid">
        <span><strong>File:</strong> ${parsed.fileName}</span>
        <span><strong>Rows Found:</strong> ${parsed.parsedRows.length}</span>
        <span><strong>Patent Rows:</strong> ${parsed.patentCount}</span>
        <span><strong>Trademark Rows:</strong> ${parsed.trademarkCount}</span>
        <span><strong>Skipped:</strong> ${parsed.skippedRows.length}</span>
      </div>
      ${parsed.skippedRows.length ? `<div style="margin-top:8px">Skipped rows are missing title or have unsupported record type.</div>` : ''}
    `;
    document.getElementById('bulk-import-summary').innerHTML = details;
    updateBulkImportProgress(25, 'File parsed. Ready to import.', `Importable rows: ${parsed.importableRows.length}`);
  } catch (e) {
    CASE_IMPORT_STATE = { parsedRows: [], importableRows: [], skippedRows: [], fileName: '' };
    document.getElementById('bulk-import-summary').innerHTML = `<span class="text-danger">${e.message}</span>`;
    updateBulkImportProgress(0, 'Import preview failed.', '');
  }
}

async function submitBulkDocketTrakImport() {
  const clientCode = document.getElementById('bulk-import-client')?.value.trim().toUpperCase() || '';
  const submitBtn = document.getElementById('bulk-import-submit');
  const selectedClient = resolveClientByCode(clientCode);
  if (!selectedClient) {
    alert('Select a valid existing client code first.');
    return;
  }
  if (!CASE_IMPORT_STATE.importableRows.length) {
    alert('Upload a valid DocketTrak Excel file first.');
    return;
  }

  const batchSize = 25;
  let imported = 0;
  let skipped = CASE_IMPORT_STATE.skippedRows.length;
  let failed = 0;
  let errorSamples = [];
  submitBtn.disabled = true;
  submitBtn.textContent = 'Importing...';

  try {
    const total = CASE_IMPORT_STATE.importableRows.length;
    for (let offset = 0; offset < total; offset += batchSize) {
      const batch = CASE_IMPORT_STATE.importableRows.slice(offset, offset + batchSize);
      updateBulkImportProgress(
        Math.round((offset / total) * 100),
        `Importing rows ${offset + 1}-${Math.min(offset + batch.length, total)} of ${total}...`,
        `Imported: ${imported} | Skipped: ${skipped} | Failed: ${failed}`
      );
      const result = await API.bulkImportDocketTrakRows({
        clientCode: selectedClient.CLIENT_CODE || selectedClient.CLIENT_ID,
        rows: batch
      });
      imported += Number(result.imported || 0);
      skipped += Number(result.skipped || 0);
      const batchErrors = Array.isArray(result.errors) ? result.errors : [];
      failed += batchErrors.length;
      errorSamples = errorSamples.concat(batchErrors).slice(0, 10);
      updateBulkImportProgress(
        Math.round((Math.min(offset + batch.length, total) / total) * 100),
        `Imported ${Math.min(offset + batch.length, total)} of ${total} rows...`,
        `Imported: ${imported} | Skipped: ${skipped} | Failed: ${failed}`
      );
    }
    const errorHtml = errorSamples.length
      ? `<div style="margin-top:8px"><strong>Sample issues:</strong><br>${errorSamples.map(item => `Row ${item.rowNumber}: ${item.reason}`).join('<br>')}</div>`
      : '';
    updateBulkImportProgress(100, 'Bulk import completed.', `Imported: ${imported} | Skipped: ${skipped} | Failed: ${failed}${errorHtml}`);
    document.getElementById('bulk-import-summary').innerHTML = `
      <div class="alert alert-success">
        Import complete for client <strong>${selectedClient.CLIENT_CODE || selectedClient.CLIENT_ID}</strong>.
        Imported ${imported} row(s), skipped ${skipped}, failed ${failed}.
      </div>`;
    await ensureClientLookupLoaded();
    await loadCases();
    await loadDashboard();
  } catch (e) {
    updateBulkImportProgress(0, 'Bulk import failed.', e.message);
    alert('Bulk import failed: ' + e.message);
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = 'Start Import';
  }
}

function openClientModalLegacyV0(jsonStr) {
  const c = jsonStr ? JSON.parse(jsonStr) : null;
  document.getElementById('modal-client-title').textContent = c ? 'Edit Client' : 'New Client';
  document.getElementById('modal-client-body').innerHTML = `
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Type</label>
        <select class="form-control" id="cl-client-type">${['Individual','Organization'].map(x=>`<option${(c?.CLIENT_TYPE||'Individual')===x?' selected':''}>${x}</option>`).join('')}</select>
      </div>
      <div class="form-group"><label class="form-label">Client Region</label>
        <select class="form-control" id="cl-region">${['India','Abroad'].map(x=>`<option${(c?.CLIENT_REGION||'India')===x?' selected':''}>${x}</option>`).join('')}</select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Code</label><input class="form-control" id="cl-code" value="${c?.CLIENT_CODE||''}" placeholder="A61M or 970Y" /></div>
      <div class="form-group"><label class="form-label">Organization ID</label><input class="form-control" id="cl-org-id" value="${c?.ORG_ID||''}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Client Name *</label><input class="form-control" id="cl-name" value="${c?.CLIENT_NAME||''}" /></div>
    <div class="form-group"><label class="form-label">Contact Person</label><input class="form-control" id="cl-contact" value="${c?.CONTACT_PERSON||''}" /></div>
    <div class="form-group"><label class="form-label">Email</label><input type="email" class="form-control" id="cl-email" value="${c?.EMAIL||''}" /></div>
    <div class="form-group"><label class="form-label">Phone</label><input class="form-control" id="cl-phone" value="${c?.PHONE||''}" /></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Client Admin User ID</label><input class="form-control" id="cl-admin-user" value="${c?.CLIENT_ADMIN_USER_ID||''}" /></div>
      <div class="form-group"><label class="form-label">Assigned Staff Email</label><input class="form-control" id="cl-staff" value="${c?.ASSIGNED_STAFF_EMAIL||''}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Address</label><textarea class="form-control" id="cl-addr">${c?.ADDRESS||''}</textarea></div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="cl-notes">${c?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-client')">Cancel</button>
      <button class="btn btn-primary" id="cl-save" onclick="saveClient('${c?.CLIENT_ID||''}')">Save Client</button>
    </div>`;
  openModal('modal-client');
}

async function saveClientLegacyV0(existingId) {
  const btn = document.getElementById('cl-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    CLIENT_ID: existingId || undefined,
    CLIENT_NAME: document.getElementById('cl-name').value.trim(),
    CONTACT_PERSON: document.getElementById('cl-contact').value.trim(),
    EMAIL: document.getElementById('cl-email').value.trim(),
    PHONE: document.getElementById('cl-phone').value.trim(),
    CLIENT_TYPE: document.getElementById('cl-client-type').value,
    CLIENT_REGION: document.getElementById('cl-region').value,
    CLIENT_CODE: document.getElementById('cl-code').value.trim(),
    ORG_ID: document.getElementById('cl-org-id').value.trim(),
    CLIENT_ADMIN_USER_ID: document.getElementById('cl-admin-user').value.trim(),
    ASSIGNED_STAFF_EMAIL: document.getElementById('cl-staff').value.trim(),
    ADDRESS: document.getElementById('cl-addr').value.trim(),
    NOTES: document.getElementById('cl-notes').value.trim(),
  };
  if (!existingId) delete data.CLIENT_ID;
  try {
    const r = await API.saveClient(data);
    if (r.success) { closeModal('modal-client'); loadMgmtClients(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save Client';
  }
}

function openCaseModalLegacyV0(jsonStr) {
  const c = jsonStr ? JSON.parse(jsonStr) : null;
  const staffOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Staff');
  const galvanizerOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Galvanizer');
  const attorneyOptions = buildUserOptions(USER_LOOKUP_CACHE, 'Attorney');
  document.getElementById('modal-case-title').textContent = c ? 'Edit Case' : 'New Case';
  document.getElementById('modal-case-body').innerHTML = `
    <div class="form-group"><label class="form-label">Client Code *</label><input class="form-control" id="mc-cid" list="mc-client-list" value="${c?.CLIENT_ID||''}" ${c?'readonly':''} placeholder="Search client code" /><datalist id="mc-client-list">${buildClientOptions(CLIENT_LOOKUP_CACHE)}</datalist></div>
    <div class="form-group"><label class="form-label">Patent Title *</label><input class="form-control" id="mc-title" value="${c?.PATENT_TITLE||''}" /></div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">App. No.</label><input class="form-control" id="mc-appno" value="${c?.APPLICATION_NUMBER||''}" /></div>
      <div class="form-group"><label class="form-label">Country</label><select class="form-control" id="mc-country">${['India','USA','EPO','PCT','China','Japan','Other'].map(x=>`<option${(c?.COUNTRY||'India')===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Filing Date</label><input type="date" class="form-control" id="mc-filed" value="${c?.FILING_DATE?new Date(c.FILING_DATE).toISOString().split('T')[0]:''}" /></div>
      <div class="form-group"><label class="form-label">Next Deadline</label><input type="date" class="form-control" id="mc-deadline" value="${c?.NEXT_DEADLINE?new Date(c.NEXT_DEADLINE).toISOString().split('T')[0]:''}" /></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Status</label><select class="form-control" id="mc-status">${['Drafted','Filed','Published','Under Examination','Granted','Abandoned','Lapsed','Refused'].map(x=>`<option${c?.CURRENT_STATUS===x?' selected':''}>${x}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Type</label><select class="form-control" id="mc-type">${['Utility','Design','PCT','Provisional'].map(x=>`<option${c?.PATENT_TYPE===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Assigned Staff</label><input class="form-control" id="mc-staff" list="mc-staff-list" value="${c?.ASSIGNED_STAFF_EMAIL||''}" placeholder="Search staff" /><datalist id="mc-staff-list">${staffOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Galvanizer</label><input class="form-control" id="mc-galvanizer" list="mc-galvanizer-list" value="${c?.GALVANIZER_EMAIL||''}" placeholder="Search galvanizer" /><datalist id="mc-galvanizer-list">${galvanizerOptions}</datalist></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Attorney</label><input class="form-control" id="mc-atty" list="mc-attorney-list" value="${c?.ATTORNEY||''}" placeholder="Search attorney" /><datalist id="mc-attorney-list">${attorneyOptions}</datalist></div>
      <div class="form-group"><label class="form-label">Workflow Stage</label><select class="form-control" id="mc-stage">${['Drafting','Ready for Attorney','Under Attorney Review','Filed'].map(x=>`<option${(c?.WORKFLOW_STAGE||'Drafting')===x?' selected':''}>${x}</option>`).join('')}</select></div>
    </div>
    <div class="form-row">
      <div class="form-group"><label class="form-label">Priority</label><select class="form-control" id="mc-priority">${['Normal','High','Urgent'].map(x=>`<option${(c?.PRIORITY||'Normal')===x?' selected':''}>${x}</option>`).join('')}</select></div>
      <div class="form-group"><label class="form-label">Org ID</label><input class="form-control" id="mc-org" value="${c?.ORG_ID||''}" /></div>
    </div>
    <div class="form-group"><label class="form-label">Notes</label><textarea class="form-control" id="mc-notes">${c?.NOTES||''}</textarea></div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal('modal-case')">Cancel</button>
      <button class="btn btn-primary" id="mc-save" onclick="saveCase('${c?.CASE_ID||''}')">Save Case</button>
    </div>`;
  openModal('modal-case');
}

async function saveCaseLegacyV0(existingId) {
  const btn = document.getElementById('mc-save');
  btn.disabled = true; btn.textContent = 'Saving…';
  const data = {
    CASE_ID: existingId || undefined,
    CLIENT_ID: document.getElementById('mc-cid').value.trim(),
    PATENT_TITLE: document.getElementById('mc-title').value.trim(),
    APPLICATION_NUMBER: document.getElementById('mc-appno').value.trim(),
    COUNTRY: document.getElementById('mc-country').value,
    FILING_DATE: document.getElementById('mc-filed').value,
    NEXT_DEADLINE: document.getElementById('mc-deadline').value,
    CURRENT_STATUS: document.getElementById('mc-status').value,
    PATENT_TYPE: document.getElementById('mc-type').value,
    ASSIGNED_STAFF_EMAIL: document.getElementById('mc-staff').value.trim(),
    GALVANIZER_EMAIL: document.getElementById('mc-galvanizer').value.trim(),
    ATTORNEY: document.getElementById('mc-atty').value.trim(),
    WORKFLOW_STAGE: document.getElementById('mc-stage').value,
    PRIORITY: document.getElementById('mc-priority').value,
    ORG_ID: document.getElementById('mc-org').value.trim(),
    NOTES: document.getElementById('mc-notes').value.trim(),
  };
  if (!existingId) delete data.CASE_ID;
  try {
    const r = await API.saveCase(data);
    if (r.success) { closeModal('modal-case'); loadCases(); loadDashboard(); }
    else throw new Error(r.error || r.message);
  } catch(e) {
    alert('Error: ' + e.message);
    btn.disabled = false; btn.textContent = 'Save Case';
  }
}

