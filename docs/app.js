// ============================================================
// app.js — talks to the Apps Script Web App at window.API_URL.
//
// IMPORTANT: every POST uses Content-Type: text/plain;charset=utf-8
// with a JSON string body. Apps Script Web Apps can't handle a
// CORS preflight (OPTIONS) request, so a normal
// "application/json" POST from the browser would be blocked
// before it ever reached the script. Sending the same JSON as
// plain text skips the preflight — the server still parses the
// body as JSON (see WebApp.gs doPost). Don't change this without
// changing WebApp.gs to match.
// ============================================================

const SESSION_KEY = 'mcd_token';

const loginView = document.getElementById('loginView');
const appView = document.getElementById('appView');
const loginForm = document.getElementById('loginForm');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');

// ---------- API helper ----------

async function apiPost(action, payload) {
  const body = Object.assign({ action }, payload);
  const res = await fetch(window.API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error('Network error (' + res.status + ')');
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

function getToken() {
  return sessionStorage.getItem(SESSION_KEY);
}

function escapeHtml(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

function formatDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return d.toLocaleDateString();
}

// ---------- view switching ----------

function showApp() {
  loginView.hidden = true;
  appView.hidden = false;
  logoutBtn.hidden = false;
  initTabs();
  runSearch({ resetAgeBounds: true });
}

function showLogin() {
  loginView.hidden = false;
  appView.hidden = true;
  logoutBtn.hidden = true;
}

if (getToken()) showApp(); else showLogin();

// ---------- login ----------

loginForm.addEventListener('submit', async (ev) => {
  ev.preventDefault();
  loginError.hidden = true;
  const password = document.getElementById('passwordInput').value;
  try {
    const data = await apiPost('login', { password });
    sessionStorage.setItem(SESSION_KEY, data.token);
    loginForm.reset();
    showApp();
  } catch (err) {
    loginError.textContent = err.message;
    loginError.hidden = false;
  }
});

logoutBtn.addEventListener('click', () => {
  sessionStorage.removeItem(SESSION_KEY);
  showLogin();
});

function handleSessionError(err) {
  if (/session/i.test(err.message)) {
    sessionStorage.removeItem(SESSION_KEY);
    showLogin();
    loginError.textContent = err.message;
    loginError.hidden = false;
    return true;
  }
  return false;
}

// ============================================================
// Tabs
// ============================================================

function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.hidden = true);
      btn.classList.add('active');
      const panel = document.getElementById(btn.dataset.tab);
      panel.hidden = false;
      if (btn.dataset.tab === 'overviewTab') loadStats();
      if (btn.dataset.tab === 'logTab') loadReservationLog();
    });
  });
}

// ============================================================
// Search & Reserve tab
// ============================================================

const searchInput = document.getElementById('searchInput');
const colonyFilter = document.getElementById('colonyFilter');
const refreshBtn = document.getElementById('refreshBtn');
const statusMsg = document.getElementById('statusMsg');
const resultsBody = document.getElementById('resultsBody');
const selectAllCheckbox = document.getElementById('selectAllCheckbox');
const selectionBar = document.getElementById('selectionBar');
const selectionCount = document.getElementById('selectionCount');
const openReserveBtn = document.getElementById('openReserveBtn');

const ageMin = document.getElementById('ageMin');
const ageMax = document.getElementById('ageMax');
const ageSliderFill = document.getElementById('ageSliderFill');
const ageRangeLabel = document.getElementById('ageRangeLabel');

let rawResults = [];
let selectedRows = new Set();
let currentSex = '';
let ageBoundsInitialized = false;

// --- sex filter ---
document.getElementById('sexFilter').addEventListener('click', (ev) => {
  const btn = ev.target.closest('.seg-btn');
  if (!btn) return;
  document.querySelectorAll('#sexFilter .seg-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  currentSex = btn.dataset.sex;
  applyFiltersAndRender();
});

// --- age dual slider ---
function updateSliderFill() {
  const min = Number(ageMin.min), max = Number(ageMin.max);
  const lo = Number(ageMin.value), hi = Number(ageMax.value);
  const pct = v => ((v - min) / (max - min || 1)) * 100;
  ageSliderFill.style.left = pct(lo) + '%';
  ageSliderFill.style.right = (100 - pct(hi)) + '%';
  ageRangeLabel.textContent = `${lo} – ${hi} w`;
}

function setAgeBounds(min, max) {
  [ageMin, ageMax].forEach(el => { el.min = min; el.max = max; });
  ageMin.value = min;
  ageMax.value = max;
  updateSliderFill();
}

ageMin.addEventListener('input', () => {
  if (Number(ageMin.value) > Number(ageMax.value)) ageMin.value = ageMax.value;
  updateSliderFill();
  applyFiltersAndRender();
});
ageMax.addEventListener('input', () => {
  if (Number(ageMax.value) < Number(ageMin.value)) ageMax.value = ageMin.value;
  updateSliderFill();
  applyFiltersAndRender();
});
updateSliderFill();

// --- fetching ---

async function runSearch(opts) {
  opts = opts || {};
  statusMsg.textContent = 'Loading…';
  try {
    const data = await apiPost('search', {
      token: getToken(),
      query: searchInput.value.trim(),
      colony: colonyFilter.value
    });
    rawResults = data.results || [];
    populateColonySelect(colonyFilter, data.colonies || []);
    populateColonySelect(document.getElementById('overviewColonyFilter'), data.colonies || []);

    if (!ageBoundsInitialized || opts.resetAgeBounds) {
      const ages = rawResults.map(r => Number(r.ageWeeks)).filter(n => !isNaN(n) && isFinite(n));
      const lo = ages.length ? Math.floor(Math.min(...ages)) : 0;
      const hi = ages.length ? Math.ceil(Math.max(...ages)) : 150;
      setAgeBounds(lo, Math.max(hi, lo + 1));
      ageBoundsInitialized = true;
    }

    selectedRows.clear();
    applyFiltersAndRender();
  } catch (err) {
    if (handleSessionError(err)) return;
    statusMsg.textContent = 'Error: ' + err.message;
  }
}

function populateColonySelect(select, colonies) {
  const current = select.value;
  select.innerHTML = '<option value="">All shared colonies</option>' +
    colonies.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  if (colonies.includes(current)) select.value = current;
}

function applyFiltersAndRender() {
  const lo = Number(ageMin.value), hi = Number(ageMax.value);
  const filtered = rawResults.filter(r => {
    if (currentSex && r.sex !== currentSex) return false;
    const age = Number(r.ageWeeks);
    if (!isNaN(age) && isFinite(age)) {
      if (age < lo || age > hi) return false;
    }
    return true;
  });
  renderResults(filtered);
  statusMsg.textContent = filtered.length + ' animal(s) found.';
  updateSelectionBar();
}

function renderResults(rows) {
  resultsBody.innerHTML = rows.map(r => {
    const checked = selectedRows.has(r.row) ? 'checked' : '';
    return `
      <tr>
        <td><input type="checkbox" class="rowCheckbox" data-row="${r.row}" ${checked} /></td>
        <td>${escapeHtml(r.colony)}</td>
        <td>${escapeHtml(r.sex)}</td>
        <td>${escapeHtml(r.ageWeeks)}</td>
        <td>${escapeHtml(r.mouseline)}</td>
        <td>${escapeHtml(r.genotype)}</td>
      </tr>`;
  }).join('');

  resultsBody.querySelectorAll('.rowCheckbox').forEach(cb => {
    cb.addEventListener('change', () => {
      const row = Number(cb.dataset.row);
      if (cb.checked) selectedRows.add(row); else selectedRows.delete(row);
      updateSelectionBar();
    });
  });

  selectAllCheckbox.checked = rows.length > 0 && rows.every(r => selectedRows.has(r.row));
}

selectAllCheckbox.addEventListener('change', () => {
  const checkboxes = resultsBody.querySelectorAll('.rowCheckbox');
  checkboxes.forEach(cb => {
    const row = Number(cb.dataset.row);
    cb.checked = selectAllCheckbox.checked;
    if (selectAllCheckbox.checked) selectedRows.add(row); else selectedRows.delete(row);
  });
  updateSelectionBar();
});

function updateSelectionBar() {
  const n = selectedRows.size;
  selectionBar.hidden = n === 0;
  selectionCount.textContent = n === 1 ? '1 animal selected' : `${n} animals selected`;
}

searchInput.addEventListener('input', debounce(() => runSearch(), 350));
colonyFilter.addEventListener('change', () => runSearch({ resetAgeBounds: true }));
refreshBtn.addEventListener('click', () => runSearch());

function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

// ============================================================
// Reservation dialog
// ============================================================

const reserveDialog = document.getElementById('reserveDialog');
const reserveForm = document.getElementById('reserveForm');
const reserveSummary = document.getElementById('reserveSummary');
const reserveError = document.getElementById('reserveError');
const cancelReserve = document.getElementById('cancelReserve');

openReserveBtn.addEventListener('click', () => {
  const selected = rawResults.filter(r => selectedRows.has(r.row));
  if (selected.length === 0) return;

  const colonies = new Set(selected.map(r => r.colony));
  if (colonies.size > 1) {
    reserveError.textContent = 'Selected animals span more than one colony — please request one colony at a time.';
    reserveError.hidden = false;
  } else {
    reserveError.hidden = true;
  }

  reserveSummary.innerHTML = selected.map(r =>
    `${escapeHtml(r.colony)} — ${escapeHtml(r.sex)}, ${escapeHtml(r.genotype)}, ${escapeHtml(r.ageWeeks)}w`
  ).join('<br>');

  reserveForm.reset();
  reserveDialog.showModal();
});

cancelReserve.addEventListener('click', () => reserveDialog.close());

reserveForm.addEventListener('submit', async (ev) => {
  ev.preventDefault();
  reserveError.hidden = true;

  const selected = rawResults.filter(r => selectedRows.has(r.row));
  const colonies = new Set(selected.map(r => r.colony));
  if (colonies.size > 1) {
    reserveError.textContent = 'Selected animals span more than one colony — please request one colony at a time.';
    reserveError.hidden = false;
    return;
  }

  try {
    await apiPost('reserve', {
      token: getToken(),
      rows: selected.map(r => r.row),
      requesterName: document.getElementById('reqName').value.trim(),
      requesterEmail: document.getElementById('reqEmail').value.trim(),
      startDate: document.getElementById('reqStartDate').value,
      anticipatedAgeMonths: document.getElementById('reqAge').value,
      experimentType: document.getElementById('reqType').value.trim(),
      conditions: document.getElementById('reqConditions').value.trim(),
      animalsPerCondition: document.getElementById('reqPerCondition').value.trim(),
      tissues: document.getElementById('reqTissues').value.trim()
    });
    reserveDialog.close();
    statusMsg.textContent = 'Request sent — the colony owner has been emailed for approval.';
    runSearch();
  } catch (err) {
    reserveError.textContent = err.message;
    reserveError.hidden = false;
  }
});

// ============================================================
// Overview tab
// ============================================================

const overviewColonyFilter = document.getElementById('overviewColonyFilter');
let statusChartInstance = null;
let genotypeChartInstance = null;

const CHART_PALETTE = ['#3c6e47', '#e8a33d', '#c25b5b', '#4f7cac', '#8a5fbf', '#3d9a8c', '#d4845c', '#6b7280'];

overviewColonyFilter.addEventListener('change', loadStats);

async function loadStats() {
  try {
    const data = await apiPost('stats', { token: getToken(), colony: overviewColonyFilter.value });
    document.getElementById('totalAnimals').textContent = data.total;

    renderDonut('statusChart', 'statusLegend', data.byState, {
      Available: '#bfe3c8', Requested: '#ffe6a3', Reserved: '#f6c2c2'
    }, (chart) => { statusChartInstance = chart; }, statusChartInstance);

    renderDonut('genotypeChart', 'genotypeLegend', data.genotypeOfAvailable, null,
      (chart) => { genotypeChartInstance = chart; }, genotypeChartInstance);
  } catch (err) {
    if (handleSessionError(err)) return;
  }
}

function renderDonut(canvasId, legendId, dataObj, colorMap, setInstance, existingInstance) {
  const labels = Object.keys(dataObj).filter(k => dataObj[k] > 0);
  const values = labels.map(l => dataObj[l]);
  const colors = labels.map((l, i) => (colorMap && colorMap[l]) || CHART_PALETTE[i % CHART_PALETTE.length]);

  if (existingInstance) existingInstance.destroy();

  const ctx = document.getElementById(canvasId).getContext('2d');
  const chart = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] },
    options: {
      plugins: { legend: { display: false } },
      cutout: '62%'
    }
  });
  setInstance(chart);

  const legend = document.getElementById(legendId);
  legend.innerHTML = labels.map((l, i) =>
    `<span class="legend-item"><span class="legend-swatch" style="background:${colors[i]}"></span>${escapeHtml(l)} (${values[i]})</span>`
  ).join('');
}

// ============================================================
// Reservation Log tab
// ============================================================

const logBody = document.getElementById('logBody');

async function loadReservationLog() {
  try {
    const data = await apiPost('reservations', { token: getToken() });
    const reservations = data.reservations || [];
    logBody.innerHTML = reservations.map(r => `
      <tr>
        <td>${formatDate(r.requestedDate)}</td>
        <td>${escapeHtml(r.requesterName)}<br><span class="muted">${escapeHtml(r.requesterEmail)}</span></td>
        <td>${escapeHtml(r.colony)}</td>
        <td>${escapeHtml(r.physicalTags)} (${escapeHtml(r.numberOfAnimals)})</td>
        <td>${escapeHtml(r.experimentType)}</td>
        <td>${escapeHtml(r.conditions)}</td>
        <td>${escapeHtml(r.animalsPerCondition)}</td>
        <td>${escapeHtml(r.tissues)}</td>
        <td>${escapeHtml(r.anticipatedAgeMonths)}</td>
        <td>${formatDate(r.startDate)}</td>
        <td><span class="badge ${String(r.status || '').toLowerCase()}">${escapeHtml(r.status)}</span></td>
        <td>${formatDate(r.decidedDate)}</td>
      </tr>`).join('');
  } catch (err) {
    if (handleSessionError(err)) return;
  }
}