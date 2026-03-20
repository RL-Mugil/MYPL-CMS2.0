let SESSION = null;
let CHARTS = {};
const CLIENT_ROLES = ['Client', 'Client Admin', 'Client Employee', 'Individual Client'];
const THEME_KEY = 'mg_client_theme';
const DOC_CATEGORIES = ['APPLICATIONS', 'OFFICE_ACTIONS', 'RESPONSES', 'CERTIFICATES', 'INVOICES', 'COMMUNICATION'];
const PAGE_TITLES = {
  dashboard: 'Overview',
  cases: 'My Cases',
  documents: 'Documents',
  requests: 'Document Requests',
  finance: 'Invoices',
  threads: 'Threads',
  notifications: 'Notifications',
  search: 'Search',
  contact: 'Contact Firm'
};

const STATE = {
  currentPage: 'dashboard',
  clients: [],
  cases: { all: [], filtered: [], limit: 20, filters: { query: '', client: '', status: '', country: '', fromDate: '', toDate: '' }, filtersOpen: false },
  documents: { all: [], filtered: [], limit: 24, filters: { query: '', client: '', category: '' }, filtersOpen: false },
  requests: { all: [], filtered: [], limit: 20, filters: { query: '', client: '', status: '', caseId: '' }, filtersOpen: false },
  finance: { all: [], filtered: [], limit: 20, filters: { query: '', client: '', status: '', month: '', fromDate: '', toDate: '' }, filtersOpen: false },
  notifications: [],
  threads: { list: [], selectedThreadId: '', messages: [] },
  search: { query: '', scope: 'all', results: null }
};

(async function bootstrap() {
  applyTheme(loadTheme());
  SESSION = await Auth.requireAuth(CLIENT_ROLES);
  if (!SESSION) {
    window.location.href = 'client-login.html';
    return;
  }

  const initials = (SESSION.name || 'C').split(' ').map(function(part) { return part ? part[0] : ''; }).join('').slice(0, 2).toUpperCase();
  document.getElementById('sidebar-avatar').textContent = initials || 'C';
  document.getElementById('sidebar-name').textContent = SESSION.name || SESSION.email || 'Client';
  document.getElementById('sidebar-role').textContent = SESSION.role || 'Client';

  await ensureAccessibleClientsLoaded();
  showPage('dashboard');
})();

function esc(value) {
  return window.Safe && window.Safe.escapeHtml ? window.Safe.escapeHtml(value == null ? '' : String(value)) : String(value == null ? '' : value);
}

function renderLoading(text) {
  return '<div class="loading-wrap"><div class="spinner"></div><span>' + esc(text || 'Loading...') + '</span></div>';
}

function renderError(message) {
  return '<div class="alert alert-error">' + esc(message || 'Something went wrong.') + '</div>';
}

function fmt(value) {
  if (!value) return '-';
  try {
    return new Date(value).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch (e) {
    return value;
  }
}

function fmtDateTime(value) {
  if (!value) return '-';
  try {
    return new Date(value).toLocaleString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  } catch (e) {
    return value;
  }
}

function money(value) {
  var amount = parseFloat(value) || 0;
  return 'Rs ' + amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function statusBadge(status) {
  var key = String(status || 'Drafted');
  var map = {
    'Filed': 'badge-filed',
    'Granted': 'badge-granted',
    'Under Examination': 'badge-exam',
    'Published': 'badge-exam',
    'Drafted': 'badge-drafted',
    'Open': 'badge-drafted',
    'Completed': 'badge-granted',
    'Pending': 'badge-exam',
    'Abandoned': 'badge-abandoned',
    'Lapsed': 'badge-abandoned',
    'Refused': 'badge-abandoned',
    'Paid': 'badge-paid',
    'Unpaid': 'badge-unpaid',
    'Approved': 'badge-granted',
    'Rejected': 'badge-unpaid'
  };
  return '<span class="badge ' + (map[key] || 'badge-drafted') + '">' + esc(key) + '</span>';
}

function readCssVar(name) {
  return getComputedStyle(document.body).getPropertyValue(name).trim();
}

function setTopbarActions(html) {
  document.getElementById('topbar-actions').innerHTML = html || '';
}

function destroyCharts() {
  Object.keys(CHARTS).forEach(function(key) {
    try { CHARTS[key].destroy(); } catch (e) {}
  });
  CHARTS = {};
}

function toggleSidebar(forceOpen) {
  var sidebar = document.querySelector('.sidebar');
  if (!sidebar) return;
  if (typeof forceOpen === 'boolean') {
    sidebar.classList.toggle('open', forceOpen);
    return;
  }
  sidebar.classList.toggle('open');
}

function toggleSidebarDesktop() {
  document.body.classList.toggle('sidebar-hidden');
}

function loadTheme() {
  return localStorage.getItem(THEME_KEY) || 'dark';
}

function applyTheme(theme) {
  document.body.classList.toggle('theme-light', theme === 'light');
  localStorage.setItem(THEME_KEY, theme);
  var btn = document.getElementById('theme-toggle');
  if (btn) btn.textContent = theme === 'light' ? 'Dark Mode' : 'Light Mode';
}

function toggleTheme() {
  applyTheme(document.body.classList.contains('theme-light') ? 'dark' : 'light');
  if (STATE.currentPage) showPage(STATE.currentPage, true);
}

async function handleSignOut() {
  await Auth.signOut();
  window.location.href = 'client-login.html';
}

async function ensureAccessibleClientsLoaded() {
  if (STATE.clients.length) return STATE.clients;
  try {
    var clients = await API.getAccessibleClients();
    STATE.clients = Array.isArray(clients) ? clients : [];
    return STATE.clients;
  } catch (e) {
    STATE.clients = [];
    return [];
  }
}

function getClientLabel(client) {
  var code = client.CLIENT_CODE || client.CLIENT_ID || '';
  var name = client.CLIENT_NAME || client.CONTACT_PERSON || '';
  return code ? (code + ' - ' + name) : name;
}

function getClientById(clientId) {
  var target = String(clientId || '').trim().toUpperCase();
  return STATE.clients.find(function(client) {
    return String(client.CLIENT_ID || '').trim().toUpperCase() === target || String(client.CLIENT_CODE || '').trim().toUpperCase() === target;
  }) || null;
}

function normalizeDateOnly(value) {
  if (!value) return '';
  var date = new Date(value);
  if (isNaN(date.getTime())) return '';
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function matchesDateRange(value, fromDate, toDate) {
  var time = normalizeDateOnly(value);
  if (!time) return !fromDate && !toDate;
  var fromTime = normalizeDateOnly(fromDate);
  var toTime = normalizeDateOnly(toDate);
  if (fromTime && time < fromTime) return false;
  if (toTime && time > toTime) return false;
  return true;
}

function csvContains(values, query) {
  var q = String(query || '').trim().toLowerCase();
  if (!q) return true;
  return values.join(' ').toLowerCase().indexOf(q) > -1;
}

function buildClientOptions(selectedValue, includeBlankLabel) {
  var options = includeBlankLabel ? ['<option value="">' + esc(includeBlankLabel) + '</option>'] : [];
  STATE.clients.forEach(function(client) {
    var id = client.CLIENT_ID || client.CLIENT_CODE || '';
    options.push('<option value="' + esc(id) + '"' + (String(selectedValue || '') === String(id) ? ' selected' : '') + '>' + esc(getClientLabel(client)) + '</option>');
  });
  return options.join('');
}

function showPage(page) {
  STATE.currentPage = page;
  document.querySelectorAll('.nav-item').forEach(function(item) {
    item.classList.toggle('active', item.dataset.page === page);
  });
  document.getElementById('topbar-title').textContent = PAGE_TITLES[page] || page;
  document.getElementById('page-content').innerHTML = renderLoading('Loading ' + (PAGE_TITLES[page] || page) + '...');
  setTopbarActions('');
  destroyCharts();
  toggleSidebar(false);

  var loader = {
    dashboard: loadDashboard,
    cases: loadCases,
    documents: loadDocuments,
    requests: loadRequests,
    finance: loadFinance,
    threads: loadThreads,
    notifications: loadNotifications,
    search: loadSearch,
    contact: loadContact
  }[page];

  if (loader) loader();
}

function fileToDataUrl(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function() { resolve(reader.result); };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function loadDashboard() {
  try {
    const summaryPromise = API.getDashboardSummary({});
    const detailsPromise = API.getDashboardDetails({});
    const notificationsPromise = API.getNotifications();
    const threadsPromise = API.getMessageThreads();
    const requestsPromise = API.getDocumentRequests({});
    const clients = await ensureAccessibleClientsLoaded();
    const results = await Promise.all([summaryPromise, detailsPromise, notificationsPromise, threadsPromise, requestsPromise]);
    const summary = results[0] || {};
    const details = results[1] || {};
    const notifications = Array.isArray(results[2]) ? results[2] : [];
    const threads = Array.isArray(results[3]) ? results[3].filter(function(thread) { return String(thread.THREAD_TYPE || '') !== 'Direct'; }) : [];
    const requests = Array.isArray(results[4]) ? results[4] : [];
    const unreadNotifications = notifications.filter(function(item) { return String(item.IS_READ || 'No') !== 'Yes'; });
    const openRequests = requests.filter(function(item) { return String(item.STATUS || 'Open') !== 'Completed'; });
    const recentThreads = threads.slice(0, 4);
    const recentCases = (details.recentActiveCases || []).slice(0, 5);
    const deadlines = (details.upcomingDeadlines || []).slice(0, 5);

    document.getElementById('page-content').innerHTML = `
      <div class="page">
        <div class="stats-grid">
          <div class="stat-card accent-primary"><div class="stat-label">Granted</div><div class="stat-value">${summary.totalGranted || 0}</div><div class="stat-sub">registered matters</div></div>
          <div class="stat-card accent-sky"><div class="stat-label">Pending</div><div class="stat-value">${summary.totalPending || 0}</div><div class="stat-sub">active matters</div></div>
          <div class="stat-card accent-amber"><div class="stat-label">Upcoming Deadlines</div><div class="stat-value">${summary.upcomingDeadlineCount || (details.upcomingDeadlines || []).length || 0}</div><div class="stat-sub">need attention</div></div>
          <div class="stat-card accent-rose"><div class="stat-label">Unread Alerts</div><div class="stat-value">${unreadNotifications.length}</div><div class="stat-sub">notifications</div></div>
          <div class="stat-card accent-emerald"><div class="stat-label">Open Threads</div><div class="stat-value">${summary.openThreads || threads.length || 0}</div><div class="stat-sub">shared conversations</div></div>
          <div class="stat-card accent-primary"><div class="stat-label">Clients</div><div class="stat-value">${clients.length}</div><div class="stat-sub">accessible records</div></div>
          <div class="stat-card accent-sky"><div class="stat-label">Open Requests</div><div class="stat-value">${openRequests.length}</div><div class="stat-sub">document workflows</div></div>
        </div>

        <div class="grid-7-5" style="margin-bottom:20px">
          <div class="card">
            <div style="display:flex;justify-content:space-around;margin-bottom:16px">
              <span class="card-title" style="margin:0">Granted by Country</span>
              <span class="card-title" style="margin:0">Pending by Country</span>
            </div>
            <div style="display:flex;justify-content:space-around;align-items:center;gap:16px;flex-wrap:wrap">
              <div class="donut-wrap"><canvas id="client-granted-chart"></canvas><div class="donut-center"><div class="donut-num">${summary.totalGranted || 0}</div><div class="donut-lbl">Granted</div></div></div>
              <div class="donut-wrap"><canvas id="client-pending-chart"></canvas><div class="donut-center"><div class="donut-num">${summary.totalPending || 0}</div><div class="donut-lbl">Pending</div></div></div>
            </div>
            <div class="chart-legend" id="dashboard-legend"></div>
          </div>

          <div style="display:flex;flex-direction:column;gap:16px">
            <div class="card">
              <div class="card-title">Upcoming Deadlines</div>
              ${deadlines.map(function(item) {
                return '<div class="list-entry"><div><div class="list-entry-label">' + esc(item.country || 'Country') + '</div><div class="list-entry-val">' + esc(item.docket || item.id || '') + '</div></div><div style="text-align:right"><div class="list-entry-label">Due</div><div class="list-entry-val" style="color:var(--amber)">' + esc(item.dateStr || fmt(item.rawDate)) + '</div></div></div>';
              }).join('') || '<div class="text-muted" style="padding:12px 0">No upcoming deadlines.</div>'}
            </div>

            <div class="card">
              <div class="card-title">Recent Shared Threads</div>
              ${recentThreads.map(function(item) {
                return '<div class="list-entry"><div><div class="list-entry-val">' + esc(item.TITLE || item.THREAD_ID) + '</div><div class="list-entry-label">' + esc(item.THREAD_TYPE || 'General') + '</div></div><button class="btn btn-ghost btn-sm" onclick="openThreadFromDashboard(\'' + esc(item.THREAD_ID) + '\')">Open</button></div>';
              }).join('') || '<div class="text-muted" style="padding:12px 0">No shared threads.</div>'}
            </div>
          </div>
        </div>

        <div class="grid-2">
          <div class="card">
            <div class="card-title">Pending by Status</div>
            <div style="height:260px"><canvas id="client-status-chart"></canvas></div>
          </div>
          <div class="card">
            <div class="card-title">Recent Active Cases</div>
            <div class="table-wrap">
              <table>
                <thead><tr><th>Case</th><th>Filing Date</th><th>Status</th></tr></thead>
                <tbody>
                  ${recentCases.map(function(item) {
                    return '<tr><td><div class="text-em">' + esc(item.docket || item.CASE_ID || '') + '</div><div class="text-muted" style="font-size:11px">' + esc(item.title || '') + '</div></td><td>' + esc(item.date || '-') + '</td><td>' + statusBadge(item.status || 'Drafted') + '</td></tr>';
                  }).join('') || '<tr><td colspan="3" class="text-muted" style="padding:20px;text-align:center">No active cases.</td></tr>'}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>`;

    renderDashboardCharts(details);
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

function renderDashboardCharts(details) {
  var grantedByCountry = details.grantedByCountry || {};
  var pendingByCountry = details.pendingByCountry || {};
  var pendingByStatus = details.pendingByStatus || {};
  var palette = {
    India: '#818cf8',
    USA: '#38bdf8',
    EPO: '#fb7185',
    PCT: '#fbbf24',
    Japan: '#34d399',
    China: '#22c55e',
    'United States': '#38bdf8',
    'European Union': '#fb7185',
    Other: '#6b7097'
  };

  var legendCountries = Object.keys(Object.assign({}, grantedByCountry, pendingByCountry));
  document.getElementById('dashboard-legend').innerHTML = legendCountries.map(function(country) {
    return '<div class="legend-item"><div class="legend-dot" style="background:' + (palette[country] || '#6b7097') + '"></div><span>' + esc(country) + '</span></div>';
  }).join('');

  Chart.defaults.color = readCssVar('--chart-text') || '#9191b4';
  Chart.defaults.font.family = "'DM Sans', sans-serif";

  function createDonut(id, dataObj) {
    var labels = Object.keys(dataObj);
    var data = labels.map(function(label) { return dataObj[label]; });
    if (!data.length) {
      labels = ['None'];
      data = [1];
    }
    CHARTS[id] = new Chart(document.getElementById(id), {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: labels.map(function(label) { return palette[label] || readCssVar('--chart-empty') || '#1e1e30'; }),
          borderWidth: 0,
          cutout: '78%'
        }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
  }

  createDonut('client-granted-chart', grantedByCountry);
  createDonut('client-pending-chart', pendingByCountry);

  CHARTS.status = new Chart(document.getElementById('client-status-chart'), {
    type: 'bar',
    data: {
      labels: Object.keys(pendingByStatus),
      datasets: [{
        data: Object.keys(pendingByStatus).map(function(label) { return pendingByStatus[label]; }),
        backgroundColor: readCssVar('--primary'),
        borderRadius: 8,
        barThickness: 18
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: readCssVar('--chart-grid') }, ticks: { color: readCssVar('--chart-text') } },
        y: { grid: { display: false }, ticks: { color: readCssVar('--chart-text') } }
      }
    }
  });
}

async function loadCases() {
  setTopbarActions('<button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer(\'cases\')">Filters</button>');
  try {
    if (!STATE.cases.all.length) STATE.cases.all = await API.getCases({});
    applyCaseFilters();
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

function toggleFilterDrawer(section) {
  if (!STATE[section]) return;
  STATE[section].filtersOpen = !STATE[section].filtersOpen;
  var renderer = { cases: applyCaseFilters, documents: applyDocumentFilters, requests: applyRequestFilters, finance: applyFinanceFilters }[section];
  if (renderer) renderer();
}

function handleCaseFilterChange(key, value) {
  STATE.cases.filters[key] = value;
  applyCaseFilters(true);
}

function resetCaseFilters() {
  STATE.cases.filters = { query: '', client: '', status: '', country: '', fromDate: '', toDate: '' };
  applyCaseFilters(true);
}

function loadMoreCases() {
  STATE.cases.limit += 20;
  applyCaseFilters();
}

function applyCaseFilters(resetLimit) {
  if (resetLimit) STATE.cases.limit = 20;
  var filters = STATE.cases.filters;
  var filtered = STATE.cases.all.filter(function(item) {
    if (filters.client && String(item.CLIENT_ID || '').trim().toUpperCase() !== String(filters.client || '').trim().toUpperCase()) return false;
    if (filters.status && String(item.CURRENT_STATUS || '') !== filters.status) return false;
    if (filters.country && String(item.COUNTRY || '') !== filters.country) return false;
    if (!matchesDateRange(item.NEXT_DEADLINE || item.FILING_DATE, filters.fromDate, filters.toDate)) return false;
    return csvContains([item.CASE_ID, item.CLIENT_ID, item.PATENT_TITLE, item.APPLICATION_NUMBER, item.COUNTRY, item.CURRENT_STATUS], filters.query);
  });

  filtered.sort(function(a, b) {
    return new Date(b.LAST_UPDATED || b.CREATED_DATE || 0) - new Date(a.LAST_UPDATED || a.CREATED_DATE || 0);
  });

  STATE.cases.filtered = filtered;
  var visible = filtered.slice(0, STATE.cases.limit);
  var countries = Array.from(new Set(STATE.cases.all.map(function(item) { return item.COUNTRY || ''; }).filter(Boolean))).sort();
  var statuses = Array.from(new Set(STATE.cases.all.map(function(item) { return item.CURRENT_STATUS || ''; }).filter(Boolean))).sort();

  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card filter-panel">
        <div class="filter-toggle-row"><button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer('cases')">${STATE.cases.filtersOpen ? 'Hide Filters' : 'Show Filters'}</button></div>
        <div class="filter-drawer ${STATE.cases.filtersOpen ? 'open' : ''}">
          <div class="filter-head"><div class="card-title">Case Filters</div><div class="filter-actions"><button class="btn btn-ghost btn-sm" onclick="resetCaseFilters()">Reset</button></div></div>
          <div class="filter-grid">
            <div><label class="form-label">Search</label><input class="form-control" value="${esc(filters.query)}" placeholder="Case ID, title, application no." oninput="handleCaseFilterChange('query', this.value)" /></div>
            <div><label class="form-label">Client</label><select class="form-control" onchange="handleCaseFilterChange('client', this.value)">${buildClientOptions(filters.client, 'All Clients')}</select></div>
            <div><label class="form-label">Status</label><select class="form-control" onchange="handleCaseFilterChange('status', this.value)"><option value="">All Statuses</option>${statuses.map(function(status) { return '<option value="' + esc(status) + '"' + (filters.status === status ? ' selected' : '') + '>' + esc(status) + '</option>'; }).join('')}</select></div>
            <div><label class="form-label">Country</label><select class="form-control" onchange="handleCaseFilterChange('country', this.value)"><option value="">All Countries</option>${countries.map(function(country) { return '<option value="' + esc(country) + '"' + (filters.country === country ? ' selected' : '') + '>' + esc(country) + '</option>'; }).join('')}</select></div>
            <div><label class="form-label">From Date</label><input class="form-control" type="date" value="${esc(filters.fromDate)}" onchange="handleCaseFilterChange('fromDate', this.value)" /></div>
            <div><label class="form-label">To Date</label><input class="form-control" type="date" value="${esc(filters.toDate)}" onchange="handleCaseFilterChange('toDate', this.value)" /></div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Case Portfolio</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Case ID</th><th>Title</th><th>Client</th><th>Country</th><th>Status</th><th>Deadline</th></tr></thead>
            <tbody>
              ${visible.map(function(item) {
                return '<tr><td class="text-em">' + esc(item.CASE_ID) + '</td><td><div class="text-em">' + esc(item.PATENT_TITLE || '') + '</div><div class="text-muted" style="font-size:11px">' + esc(item.APPLICATION_NUMBER || '') + '</div></td><td>' + esc((getClientById(item.CLIENT_ID) || {}).CLIENT_NAME || item.CLIENT_ID || '-') + '</td><td>' + esc(item.COUNTRY || '-') + '</td><td>' + statusBadge(item.CURRENT_STATUS || 'Drafted') + '</td><td>' + esc(fmt(item.NEXT_DEADLINE)) + '</td></tr>';
              }).join('') || '<tr><td colspan="6" class="text-muted" style="padding:20px;text-align:center">No cases found for the selected filters.</td></tr>'}
            </tbody>
          </table>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
          <div class="text-muted">${visible.length} of ${filtered.length} cases</div>
          ${visible.length < filtered.length ? '<button class="btn btn-primary btn-sm" onclick="loadMoreCases()">Load More</button>' : ''}
        </div>
      </div>
    </div>`;
}

async function loadDocuments() {
  setTopbarActions('<button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer(\'documents\')">Filters</button>');
  try {
    if (!STATE.documents.all.length) {
      const docs = await API.getDocuments();
      var flattened = [];
      Object.keys(docs || {}).forEach(function(category) {
        (docs[category] || []).forEach(function(file) {
          flattened.push(Object.assign({ category: category }, file));
        });
      });
      STATE.documents.all = flattened;
    }
    applyDocumentFilters();
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

function handleDocumentFilterChange(key, value) {
  STATE.documents.filters[key] = value;
  applyDocumentFilters(true);
}

function resetDocumentFilters() {
  STATE.documents.filters = { query: '', client: '', category: '' };
  applyDocumentFilters(true);
}

function loadMoreDocuments() {
  STATE.documents.limit += 24;
  applyDocumentFilters();
}

function applyDocumentFilters(resetLimit) {
  if (resetLimit) STATE.documents.limit = 24;
  var filters = STATE.documents.filters;
  var filtered = STATE.documents.all.filter(function(item) {
    if (filters.client && String(item.clientId || '').trim().toUpperCase() !== String(filters.client || '').trim().toUpperCase()) return false;
    if (filters.category && String(item.category || '') !== filters.category) return false;
    return csvContains([item.name, item.clientId, item.category, item.type], filters.query);
  });

  filtered.sort(function(a, b) {
    return new Date(b.date || 0) - new Date(a.date || 0);
  });

  STATE.documents.filtered = filtered;
  var visible = filtered.slice(0, STATE.documents.limit);
  var uploadClientValue = STATE.documents.filters.client || ((STATE.clients[0] && (STATE.clients[0].CLIENT_ID || STATE.clients[0].CLIENT_CODE)) || '');

  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="grid-2" style="margin-bottom:20px">
        <div class="card filter-panel">
          <div class="filter-toggle-row"><button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer('documents')">${STATE.documents.filtersOpen ? 'Hide Filters' : 'Show Filters'}</button></div>
          <div class="filter-drawer ${STATE.documents.filtersOpen ? 'open' : ''}">
            <div class="filter-head"><div class="card-title">Document Filters</div><div class="filter-actions"><button class="btn btn-ghost btn-sm" onclick="resetDocumentFilters()">Reset</button></div></div>
            <div class="filter-grid compact-2">
              <div><label class="form-label">Search</label><input class="form-control" value="${esc(filters.query)}" placeholder="Document name or type" oninput="handleDocumentFilterChange('query', this.value)" /></div>
              <div><label class="form-label">Client</label><select class="form-control" onchange="handleDocumentFilterChange('client', this.value)">${buildClientOptions(filters.client, 'All Clients')}</select></div>
              <div><label class="form-label">Category</label><select class="form-control" onchange="handleDocumentFilterChange('category', this.value)"><option value="">All Categories</option>${DOC_CATEGORIES.map(function(category) { return '<option value="' + esc(category) + '"' + (filters.category === category ? ' selected' : '') + '>' + esc(category.replace(/_/g, ' ')) + '</option>'; }).join('')}</select></div>
            </div>
          </div>
        </div>

        <div class="card">
          <div class="card-title">Upload a Document</div>
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Client</label>
              <select class="form-control" id="docs-upload-client">${buildClientOptions(uploadClientValue, 'Select Client')}</select>
            </div>
            <div class="form-group">
              <label class="form-label">Category</label>
              <select class="form-control" id="docs-upload-category">${DOC_CATEGORIES.map(function(category) { return '<option value="' + esc(category) + '">' + esc(category.replace(/_/g, ' ')) + '</option>'; }).join('')}</select>
            </div>
          </div>
          <div class="form-group">
            <label class="form-label">File</label>
            <input class="form-control" id="docs-upload-file" type="file" />
          </div>
          <button class="btn btn-primary" id="docs-upload-btn" onclick="uploadClientDocument()">Upload Document</button>
          <div id="docs-upload-status" style="margin-top:12px"></div>
        </div>
      </div>

      <div class="card">
        <div class="card-title">Document Library</div>
        ${visible.map(function(item) {
          return '<div class="doc-item"><div class="doc-icon">DOC</div><div style="flex:1;min-width:0"><a class="doc-name" href="' + esc(item.url) + '" target="_blank">' + esc(item.name || 'Document') + '</a><div class="doc-meta">' + esc((item.category || '').replace(/_/g, ' ')) + ' | ' + esc(item.clientId || '-') + ' | ' + esc(fmt(item.date)) + '</div></div></div>';
        }).join('') || '<div class="empty-state"><p>No documents available.</p></div>'}
        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
          <div class="text-muted">${visible.length} of ${filtered.length} documents</div>
          ${visible.length < filtered.length ? '<button class="btn btn-primary btn-sm" onclick="loadMoreDocuments()">Load More</button>' : ''}
        </div>
      </div>
    </div>`;
}

async function uploadClientDocument() {
  var clientCode = document.getElementById('docs-upload-client').value;
  var category = document.getElementById('docs-upload-category').value;
  var fileInput = document.getElementById('docs-upload-file');
  var status = document.getElementById('docs-upload-status');
  var file = fileInput.files && fileInput.files[0];
  if (!clientCode || !file) {
    status.innerHTML = '<div class="alert alert-error">Select a client and file first.</div>';
    return;
  }

  var button = document.getElementById('docs-upload-btn');
  button.disabled = true;
  button.textContent = 'Uploading...';
  status.innerHTML = '<div class="alert alert-info">Uploading document...</div>';
  try {
    var dataUrl = await fileToDataUrl(file);
    await API.uploadPortalDocument({
      clientCode: clientCode,
      subfolderName: category,
      fileName: file.name,
      mimeType: file.type || 'application/octet-stream',
      dataUrl: dataUrl
    });
    STATE.documents.all = [];
    await loadDocuments();
  } catch (e) {
    status.innerHTML = '<div class="alert alert-error">' + esc(e.message) + '</div>';
  } finally {
    button.disabled = false;
    button.textContent = 'Upload Document';
  }
}

async function loadRequests() {
  setTopbarActions('<button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer(\'requests\')">Filters</button>');
  try {
    STATE.requests.all = await API.getDocumentRequests({});
    applyRequestFilters();
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

function handleRequestFilterChange(key, value) {
  STATE.requests.filters[key] = value;
  applyRequestFilters(true);
}

function resetRequestFilters() {
  STATE.requests.filters = { query: '', client: '', status: '', caseId: '' };
  applyRequestFilters(true);
}

function loadMoreRequests() {
  STATE.requests.limit += 20;
  applyRequestFilters();
}

function applyRequestFilters(resetLimit) {
  if (resetLimit) STATE.requests.limit = 20;
  var filters = STATE.requests.filters;
  var filtered = STATE.requests.all.filter(function(item) {
    if (filters.client && String(item.CLIENT_ID || '').trim().toUpperCase() !== String(filters.client || '').trim().toUpperCase()) return false;
    if (filters.status && String(item.STATUS || '') !== filters.status) return false;
    if (filters.caseId && String(item.CASE_ID || '').trim().toUpperCase().indexOf(String(filters.caseId || '').trim().toUpperCase()) === -1) return false;
    return csvContains([item.REQUEST_ID, item.TITLE, item.DESCRIPTION, item.CLIENT_ID, item.CASE_ID, item.STATUS], filters.query);
  });

  filtered.sort(function(a, b) {
    return new Date(b.CREATED_AT || 0) - new Date(a.CREATED_AT || 0);
  });

  STATE.requests.filtered = filtered;
  var visible = filtered.slice(0, STATE.requests.limit);
  var statuses = Array.from(new Set(STATE.requests.all.map(function(item) { return item.STATUS || ''; }).filter(Boolean))).sort();

  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card filter-panel" style="margin-bottom:16px">
        <div class="filter-toggle-row"><button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer('requests')">${STATE.requests.filtersOpen ? 'Hide Filters' : 'Show Filters'}</button></div>
        <div class="filter-drawer ${STATE.requests.filtersOpen ? 'open' : ''}">
          <div class="filter-head"><div class="card-title">Request Filters</div><div class="filter-actions"><button class="btn btn-ghost btn-sm" onclick="resetRequestFilters()">Reset</button></div></div>
          <div class="filter-grid compact-2">
            <div><label class="form-label">Search</label><input class="form-control" value="${esc(filters.query)}" placeholder="Request title or ID" oninput="handleRequestFilterChange('query', this.value)" /></div>
            <div><label class="form-label">Client</label><select class="form-control" onchange="handleRequestFilterChange('client', this.value)">${buildClientOptions(filters.client, 'All Clients')}</select></div>
            <div><label class="form-label">Status</label><select class="form-control" onchange="handleRequestFilterChange('status', this.value)"><option value="">All Statuses</option>${statuses.map(function(status) { return '<option value="' + esc(status) + '"' + (filters.status === status ? ' selected' : '') + '>' + esc(status) + '</option>'; }).join('')}</select></div>
            <div><label class="form-label">Case ID</label><input class="form-control" value="${esc(filters.caseId)}" placeholder="Case ID" oninput="handleRequestFilterChange('caseId', this.value)" /></div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Document Request Workflow</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>ID</th><th>Title</th><th>Client</th><th>Case</th><th>Due</th><th>Status</th><th>Drive</th></tr></thead>
            <tbody>
              ${visible.map(function(item) {
                return '<tr><td class="text-em">' + esc(item.REQUEST_ID || '') + '</td><td><div class="text-em">' + esc(item.TITLE || '') + '</div><div class="text-muted" style="font-size:11px">' + esc(item.DESCRIPTION || '') + '</div></td><td>' + esc(item.CLIENT_ID || '-') + '</td><td>' + esc(item.CASE_ID || '-') + '</td><td>' + esc(fmt(item.DUE_DATE)) + '</td><td>' + statusBadge(item.STATUS || 'Open') + '</td><td>' + (item.DRIVE_LINK ? '<a href="' + esc(item.DRIVE_LINK) + '" target="_blank">Open</a>' : '-') + '</td></tr>';
              }).join('') || '<tr><td colspan="7" class="text-muted" style="padding:20px;text-align:center">No document requests found.</td></tr>'}
            </tbody>
          </table>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
          <div class="text-muted">${visible.length} of ${filtered.length} requests</div>
          ${visible.length < filtered.length ? '<button class="btn btn-primary btn-sm" onclick="loadMoreRequests()">Load More</button>' : ''}
        </div>
      </div>
    </div>`;
}

async function loadFinance() {
  setTopbarActions('<button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer(\'finance\')">Filters</button>');
  try {
    STATE.finance.all = await API.getInvoices();
    applyFinanceFilters();
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

function handleFinanceFilterChange(key, value) {
  STATE.finance.filters[key] = value;
  applyFinanceFilters(true);
}

function resetFinanceFilters() {
  STATE.finance.filters = { query: '', client: '', status: '', month: '', fromDate: '', toDate: '' };
  applyFinanceFilters(true);
}

function loadMoreFinance() {
  STATE.finance.limit += 20;
  applyFinanceFilters();
}

function applyFinanceFilters(resetLimit) {
  if (resetLimit) STATE.finance.limit = 20;
  var filters = STATE.finance.filters;
  var filtered = STATE.finance.all.filter(function(item) {
    var invoiceDate = item.INVOICE_DATE || item.CREATED_AT;
    var invoiceMonth = invoiceDate ? new Date(invoiceDate).toISOString().slice(0, 7) : '';
    if (filters.client && String(item.CLIENT_ID || '').trim().toUpperCase() !== String(filters.client || '').trim().toUpperCase()) return false;
    if (filters.status && String(item.PAYMENT_STATUS || '') !== filters.status) return false;
    if (filters.month && invoiceMonth !== filters.month) return false;
    if (!matchesDateRange(invoiceDate, filters.fromDate, filters.toDate)) return false;
    return csvContains([item.INVOICE_ID, item.CLIENT_ID, item.SERVICE_TYPE, item.DESCRIPTION, item.PAYMENT_STATUS], filters.query);
  });

  filtered.sort(function(a, b) {
    return new Date(b.INVOICE_DATE || b.CREATED_AT || 0) - new Date(a.INVOICE_DATE || a.CREATED_AT || 0);
  });

  STATE.finance.filtered = filtered;
  var visible = filtered.slice(0, STATE.finance.limit);
  var pendingAmount = filtered.filter(function(item) { return String(item.PAYMENT_STATUS || '') !== 'Paid'; }).reduce(function(sum, item) { return sum + (parseFloat(item.TOTAL) || 0); }, 0);
  var paidAmount = filtered.filter(function(item) { return String(item.PAYMENT_STATUS || '') === 'Paid'; }).reduce(function(sum, item) { return sum + (parseFloat(item.TOTAL) || 0); }, 0);

  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="stats-grid">
        <div class="stat-card accent-rose"><div class="stat-label">Pending Amount</div><div class="stat-value">${money(pendingAmount)}</div><div class="stat-sub">outstanding</div></div>
        <div class="stat-card accent-emerald"><div class="stat-label">Paid Amount</div><div class="stat-value">${money(paidAmount)}</div><div class="stat-sub">received</div></div>
        <div class="stat-card accent-primary"><div class="stat-label">Invoices</div><div class="stat-value">${filtered.length}</div><div class="stat-sub">matching filters</div></div>
      </div>
      <div class="card filter-panel" style="margin-bottom:16px">
        <div class="filter-toggle-row"><button class="btn btn-ghost btn-sm" onclick="toggleFilterDrawer('finance')">${STATE.finance.filtersOpen ? 'Hide Filters' : 'Show Filters'}</button></div>
        <div class="filter-drawer ${STATE.finance.filtersOpen ? 'open' : ''}">
          <div class="filter-head"><div class="card-title">Invoice Filters</div><div class="filter-actions"><button class="btn btn-ghost btn-sm" onclick="resetFinanceFilters()">Reset</button></div></div>
          <div class="filter-grid">
            <div><label class="form-label">Search</label><input class="form-control" value="${esc(filters.query)}" placeholder="Invoice ID or service type" oninput="handleFinanceFilterChange('query', this.value)" /></div>
            <div><label class="form-label">Client</label><select class="form-control" onchange="handleFinanceFilterChange('client', this.value)">${buildClientOptions(filters.client, 'All Clients')}</select></div>
            <div><label class="form-label">Status</label><select class="form-control" onchange="handleFinanceFilterChange('status', this.value)"><option value="">All Statuses</option><option value="Paid"${filters.status === 'Paid' ? ' selected' : ''}>Paid</option><option value="Unpaid"${filters.status === 'Unpaid' ? ' selected' : ''}>Unpaid</option></select></div>
            <div><label class="form-label">Month</label><input class="form-control" type="month" value="${esc(filters.month)}" onchange="handleFinanceFilterChange('month', this.value)" /></div>
            <div><label class="form-label">From Date</label><input class="form-control" type="date" value="${esc(filters.fromDate)}" onchange="handleFinanceFilterChange('fromDate', this.value)" /></div>
            <div><label class="form-label">To Date</label><input class="form-control" type="date" value="${esc(filters.toDate)}" onchange="handleFinanceFilterChange('toDate', this.value)" /></div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Invoice Ledger</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Invoice ID</th><th>Date</th><th>Service</th><th>Total</th><th>Status</th><th>PDF</th></tr></thead>
            <tbody>
              ${visible.map(function(item) {
                return '<tr><td class="text-em">' + esc(item.INVOICE_ID || '') + '</td><td>' + esc(fmt(item.INVOICE_DATE)) + '</td><td><div class="text-em">' + esc(item.SERVICE_TYPE || '') + '</div><div class="text-muted" style="font-size:11px">' + esc(item.DESCRIPTION || '') + '</div></td><td>' + esc(money(item.TOTAL)) + '</td><td>' + statusBadge(item.PAYMENT_STATUS || 'Unpaid') + '</td><td>' + (item.INVOICE_PDF_LINK ? '<a href="' + esc(item.INVOICE_PDF_LINK) + '" target="_blank">Open</a>' : '-') + '</td></tr>';
              }).join('') || '<tr><td colspan="6" class="text-muted" style="padding:20px;text-align:center">No invoices found.</td></tr>'}
            </tbody>
          </table>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
          <div class="text-muted">${visible.length} of ${filtered.length} invoices</div>
          ${visible.length < filtered.length ? '<button class="btn btn-primary btn-sm" onclick="loadMoreFinance()">Load More</button>' : ''}
        </div>
      </div>
    </div>`;
}

async function loadNotifications() {
  setTopbarActions('<button class="btn btn-ghost btn-sm" onclick="clearAllNotifications()">Clear All</button>');
  try {
    STATE.notifications = await API.getNotifications();
    document.getElementById('page-content').innerHTML = '<div class="page">' + ((STATE.notifications || []).map(function(item) {
      return '<div class="card" style="margin-bottom:12px"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start"><div><div class="text-em">' + esc(item.TITLE || 'Notification') + '</div><div class="text-muted" style="font-size:13px;margin-top:6px">' + esc(item.BODY || '') + '</div><div class="text-muted" style="font-size:11px;margin-top:8px">' + esc(fmtDateTime(item.CREATED_AT)) + '</div></div><div class="action-row">' + (String(item.IS_READ || 'No') === 'Yes' ? '' : '<button class="btn btn-ghost btn-sm" onclick="markNotificationRead(\'' + esc(item.NOTIFICATION_ID) + '\')">Mark Read</button>') + '<button class="btn btn-danger btn-sm" onclick="clearNotification(\'' + esc(item.NOTIFICATION_ID) + '\')">Clear</button></div></div></div>';
    }).join('') || '<div class="empty-state"><p>No notifications.</p></div>') + '</div>';
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

async function markNotificationRead(notificationId) {
  await API.markNotificationRead(notificationId);
  loadNotifications();
}

async function clearNotification(notificationId) {
  await API.deleteNotification(notificationId);
  loadNotifications();
}

async function clearAllNotifications() {
  await API.clearNotifications();
  loadNotifications();
}

async function loadThreads() {
  setTopbarActions('<button class="btn btn-primary btn-sm" onclick="openThreadModal()">+ New Thread</button>');
  try {
    STATE.threads.list = (await API.getMessageThreads()).filter(function(thread) {
      return String(thread.THREAD_TYPE || '') !== 'Direct';
    });
    if (!STATE.threads.selectedThreadId && STATE.threads.list.length) {
      STATE.threads.selectedThreadId = STATE.threads.list[0].THREAD_ID;
    }
    await renderThreadsPage();
  } catch (e) {
    document.getElementById('page-content').innerHTML = renderError(e.message);
  }
}

async function renderThreadsPage() {
  var threads = STATE.threads.list || [];
  var selectedThread = threads.find(function(thread) { return thread.THREAD_ID === STATE.threads.selectedThreadId; }) || null;
  if (selectedThread) {
    var messages = await API.getThreadMessages(selectedThread.THREAD_ID);
    STATE.threads.messages = Array.isArray(messages) ? messages : [];
    await API.markThreadRead(selectedThread.THREAD_ID);
  } else {
    STATE.threads.messages = [];
  }

  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="inbox-layout">
        <div class="card inbox-sidebar">
          <div class="inbox-sidebar-head"><div class="card-title" style="margin:0">Visible Threads</div></div>
          <div class="chat-list">
            ${threads.map(function(thread) {
              return '<button class="chat-list-item ' + (selectedThread && selectedThread.THREAD_ID === thread.THREAD_ID ? 'active' : '') + '" onclick="selectThread(\'' + esc(thread.THREAD_ID) + '\')"><div class="chat-avatar">' + esc((thread.TITLE || 'T').slice(0, 1).toUpperCase()) + '</div><div class="chat-meta"><div class="chat-head"><div class="chat-name">' + esc(thread.TITLE || thread.THREAD_ID) + '</div><div class="chat-time">' + esc(fmt(thread.LAST_MESSAGE_AT || thread.CREATED_AT)) + '</div></div><div class="chat-preview-row"><div class="chat-preview">' + esc(thread.THREAD_TYPE || 'General') + '</div><div class="chat-role">' + esc(thread.STATUS || 'Open') + '</div></div></div></button>';
            }).join('') || '<div class="empty-state" style="padding:24px 8px"><p>No shared threads.</p></div>'}
          </div>
        </div>
        <div class="card inbox-panel">
          ${selectedThread ? `
            <div class="inbox-chat-head">
              <div>
                <div class="chat-name" style="font-size:18px">${esc(selectedThread.TITLE || selectedThread.THREAD_ID)}</div>
                <div class="chat-role">${esc(selectedThread.THREAD_TYPE || 'General')} | ${esc(selectedThread.STATUS || 'Open')}</div>
              </div>
            </div>
            <div class="chat-thread" id="client-thread-scroll">
              ${STATE.threads.messages.map(function(message) {
                var mine = String(message.SENDER_EMAIL || '').toLowerCase() === String(SESSION.email || '').toLowerCase();
                return '<div class="chat-bubble-row ' + (mine ? 'mine' : '') + '"><div class="chat-bubble ' + (mine ? 'mine' : 'theirs') + '"><div class="chat-bubble-author">' + esc(message.SENDER_NAME || message.SENDER_EMAIL || 'User') + '</div><div class="chat-bubble-text">' + esc(message.MESSAGE_TEXT || '') + '</div><div class="chat-bubble-time">' + esc(fmtDateTime(message.CREATED_AT)) + '</div></div></div>';
              }).join('') || '<div class="empty-state"><p>No messages in this thread yet.</p></div>'}
            </div>
            <div class="chat-composer">
              <div class="form-group" style="margin-bottom:10px"><textarea class="form-control" id="thread-reply" rows="3" placeholder="Reply to this thread..."></textarea></div>
              <div class="chat-composer-actions"><button class="btn btn-primary" onclick="sendThreadReply()">Send Reply</button></div>
            </div>` : `
            <div class="empty-state" style="margin:auto"><p>Select a thread to read or reply.</p></div>`}
        </div>
      </div>
    </div>`;

  setTimeout(function() {
    var scroller = document.getElementById('client-thread-scroll');
    if (scroller) scroller.scrollTop = scroller.scrollHeight;
  }, 0);
}

async function selectThread(threadId) {
  STATE.threads.selectedThreadId = threadId;
  await renderThreadsPage();
}

async function openThreadFromDashboard(threadId) {
  STATE.threads.selectedThreadId = threadId;
  showPage('threads');
}

function openThreadModal() {
  document.getElementById('thread-title').value = '';
  document.getElementById('thread-case-id').value = '';
  document.getElementById('thread-message').value = '';
  document.getElementById('thread-client').innerHTML = buildClientOptions('', 'Select Client');
  if (STATE.clients.length === 1) {
    document.getElementById('thread-client').value = STATE.clients[0].CLIENT_ID || STATE.clients[0].CLIENT_CODE || '';
  }
  document.getElementById('thread-modal').classList.add('open');
}

function closeThreadModal() {
  document.getElementById('thread-modal').classList.remove('open');
}

async function createClientThread() {
  var title = document.getElementById('thread-title').value.trim();
  var clientId = document.getElementById('thread-client').value;
  var threadType = document.getElementById('thread-type').value;
  var caseId = document.getElementById('thread-case-id').value.trim();
  var message = document.getElementById('thread-message').value.trim();
  if (!title || !clientId || !message) {
    alert('Title, client, and opening message are required.');
    return;
  }
  if (threadType === 'Case' && !caseId) {
    alert('Case ID is required for case threads.');
    return;
  }

  var button = document.getElementById('thread-create-btn');
  button.disabled = true;
  button.textContent = 'Creating...';
  try {
    var client = getClientById(clientId) || {};
    var relatedEntityType = threadType === 'Case' ? 'CASE' : (threadType === 'Client' ? 'CLIENT' : '');
    var relatedEntityId = threadType === 'Case' ? caseId : (threadType === 'Client' ? clientId : '');
    var result = await API.saveMessageThread({
      TITLE: title,
      THREAD_TYPE: threadType,
      RELATED_ENTITY_TYPE: relatedEntityType,
      RELATED_ENTITY_ID: relatedEntityId,
      CLIENT_ID: client.CLIENT_ID || clientId,
      ORG_ID: client.ORG_ID || SESSION.orgId || '',
      VISIBLE_TO_CLIENT: 'Yes',
      STATUS: 'Open'
    });
    await API.sendThreadMessage({
      threadId: result.threadId,
      messageText: message,
      isInternal: 'No',
      threadStatus: 'Open'
    });
    closeThreadModal();
    STATE.threads.selectedThreadId = result.threadId;
    await loadThreads();
  } catch (e) {
    alert('Thread creation failed: ' + e.message);
  } finally {
    button.disabled = false;
    button.textContent = 'Create Thread';
  }
}

async function sendThreadReply() {
  var selectedThreadId = STATE.threads.selectedThreadId;
  var text = document.getElementById('thread-reply').value.trim();
  if (!selectedThreadId || !text) return;
  try {
    await API.sendThreadMessage({
      threadId: selectedThreadId,
      messageText: text,
      isInternal: 'No',
      threadStatus: 'Open'
    });
    await renderThreadsPage();
  } catch (e) {
    alert('Failed to send reply: ' + e.message);
  }
}

function loadSearch() {
  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card" style="margin-bottom:16px">
        <div class="card-title">Smart Search</div>
        <div class="form-row">
          <div class="form-group" style="margin-bottom:0">
            <label class="form-label">Query</label>
            <input class="form-control" id="smart-query" value="${esc(STATE.search.query)}" placeholder="Case ID, title, document, invoice..." onkeydown="if(event.key==='Enter')runSmartSearch()" />
          </div>
          <div class="form-group" style="margin-bottom:0">
            <label class="form-label">Scope</label>
            <select class="form-control" id="smart-scope">
              <option value="all"${STATE.search.scope === 'all' ? ' selected' : ''}>All</option>
              <option value="cases"${STATE.search.scope === 'cases' ? ' selected' : ''}>Cases</option>
              <option value="clients"${STATE.search.scope === 'clients' ? ' selected' : ''}>Clients</option>
              <option value="messages"${STATE.search.scope === 'messages' ? ' selected' : ''}>Threads</option>
            </select>
          </div>
        </div>
        <button class="btn btn-primary" onclick="runSmartSearch()">Search</button>
      </div>
      <div id="smart-results">${renderSmartSearchResults()}</div>
    </div>`;
}

function renderSmartSearchResults() {
  var data = STATE.search.results;
  if (!data) {
    return '<div class="card"><div class="text-muted">Run a search to see client-safe results across cases, threads, and clients.</div></div>';
  }
  return `
    <div class="grid-2">
      <div class="card"><div class="card-title">Cases</div>${(data.cases || []).map(function(item) { return '<div style="margin-bottom:12px"><div class="text-em">' + esc(item.CASE_ID || '') + '</div><div>' + esc(item.PATENT_TITLE || '') + '</div></div>'; }).join('') || '<div class="text-muted">No case results.</div>'}</div>
      <div class="card"><div class="card-title">Clients</div>${(data.clients || []).map(function(item) { return '<div style="margin-bottom:12px"><div class="text-em">' + esc(item.CLIENT_CODE || item.CLIENT_ID || '') + '</div><div>' + esc(item.CLIENT_NAME || '') + '</div></div>'; }).join('') || '<div class="text-muted">No client results.</div>'}</div>
      <div class="card"><div class="card-title">Threads</div>${(data.threads || []).map(function(item) { return '<div style="margin-bottom:12px"><div class="text-em">' + esc(item.TITLE || item.THREAD_ID || '') + '</div><div class="text-muted">' + esc(item.THREAD_TYPE || '') + '</div></div>'; }).join('') || '<div class="text-muted">No thread results.</div>'}</div>
      <div class="card"><div class="card-title">Tasks</div>${(data.tasks || []).map(function(item) { return '<div style="margin-bottom:12px"><div class="text-em">' + esc(item.TITLE || item.TASK_ID || '') + '</div><div class="text-muted">' + esc(item.STATUS || '') + '</div></div>'; }).join('') || '<div class="text-muted">No task results.</div>'}</div>
    </div>`;
}

async function runSmartSearch() {
  var query = document.getElementById('smart-query').value.trim();
  var scope = document.getElementById('smart-scope').value;
  if (!query) {
    STATE.search.query = '';
    STATE.search.scope = scope;
    STATE.search.results = null;
    loadSearch();
    return;
  }
  try {
    STATE.search.query = query;
    STATE.search.scope = scope;
    STATE.search.results = await API.getSmartSearch(query, scope);
    document.getElementById('smart-results').innerHTML = renderSmartSearchResults();
  } catch (e) {
    document.getElementById('smart-results').innerHTML = renderError(e.message);
  }
}

function loadContact() {
  document.getElementById('page-content').innerHTML = `
    <div class="page">
      <div class="card" style="max-width:620px">
        <div class="card-title">Message Your Firm</div>
        <div class="form-group"><label class="form-label">Subject</label><input class="form-control" id="contact-subject" placeholder="What is this regarding?" /></div>
        <div class="form-group"><label class="form-label">Related Case ID (optional)</label><input class="form-control" id="contact-case-id" placeholder="Example: 397M001" /></div>
        <div class="form-group"><label class="form-label">Message</label><textarea class="form-control" id="contact-message" rows="5" placeholder="Write your message..."></textarea></div>
        <button class="btn btn-primary" onclick="submitContact()">Send Message</button>
        <div id="contact-status" style="margin-top:12px"></div>
      </div>
    </div>`;
}

async function submitContact() {
  var subject = document.getElementById('contact-subject').value.trim();
  var caseId = document.getElementById('contact-case-id').value.trim();
  var message = document.getElementById('contact-message').value.trim();
  var status = document.getElementById('contact-status');
  if (!subject || !message) {
    status.innerHTML = '<div class="alert alert-error">Subject and message are required.</div>';
    return;
  }
  try {
    var result = await API.submitContact(subject, caseId, message);
    status.innerHTML = '<div class="alert alert-success">' + esc(result.message || 'Message sent.') + '</div>';
    document.getElementById('contact-subject').value = '';
    document.getElementById('contact-case-id').value = '';
    document.getElementById('contact-message').value = '';
  } catch (e) {
    status.innerHTML = '<div class="alert alert-error">' + esc(e.message) + '</div>';
  }
}

window.addEventListener('click', function(event) {
  if (event.target && event.target.id === 'thread-modal') {
    closeThreadModal();
  }
});
