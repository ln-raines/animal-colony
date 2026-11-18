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
const dashboardView = document.getElementById('dashboardView');
const loginForm = document.getElementById('loginForm');
const loginError = document.getElementById('loginError');
const logoutBtn = document.getElementById('logoutBtn');
const searchInput = document.getElementById('searchInput');
const colonyFilter = document.getElementById('colonyFilter');
const refreshBtn = document.getElementById('refreshBtn');
const statusMsg = document.getElementById('statusMsg');
const resultsBody = document.getElementById('resultsBody');

const reserveDialog = document.getElementById('reserveDialog');
const reserveForm = document.getElementById('reserveForm');
const reserveSummary = document.getElementById('reserveSummary');
const reserveError = document.getElementById('reserveError');
const cancelReserve = document.getElementById('cancelReserve');

let currentRows = [];
let reserveTargetRow = null;

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

// ---------- view switching ----------

function showDashboard() {
  loginView.hidden = true;
  dashboardView.hidden = false;
  logoutBtn.hidden = false;
  runSearch();
}

function showLogin() {
  loginView.hidden = false;
  dashboardView.hidden = true;
  logoutBtn.hidden = true;
}

if (getToken()) showDashboard(); else showLogin();

// ---------- login ----------

loginForm.addEventListener('submit', async (ev) => {
  ev.preventDefault();
  loginError.hidden = true;
  const password = document.getElementById('passwordInput').value;
  try {
    const data = await apiPost('login', { password });
    sessionStorage.setItem(SESSION_KEY, data.token);
    loginForm.reset();
    showDashboard();
  } catch (err) {
    loginError.textContent = err.message;
    loginError.hidden = false;
  }
});

logoutBtn.addEventListener('click', () => {
  sessionStorage.removeItem(SESSION_KEY);
  showLogin();
});

// ---------- search ----------

async function runSearch() {
  statusMsg.textContent = 'Loading…';
  resultsBody.innerHTML = '';
  try {
    const data = await apiPost('search', {
      token: getToken(),
      query: searchInput.value.trim(),
      colony: colonyFilter.value
    });
    currentRows = data.results || [];
    populateColonyFilter(data.colonies || []);
    renderResults(currentRows);
    statusMsg.textContent = currentRows.length + ' animal(s) found.';
  } catch (err) {
    if (/session/i.test(err.message)) {
      sessionStorage.removeItem(SESSION_KEY);
      showLogin();
      loginError.textContent = err.message;
      loginError.hidden = false;
      return;
    }
    statusMsg.textContent = 'Error: ' + err.message;
  }
}

function populateColonyFilter(colonies) {
  const current = colonyFilter.value;
  colonyFilter.innerHTML = '<option value="">All shared colonies</option>' +
    colonies.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  if (colonies.includes(current)) colonyFilter.value = current;
}

function renderResults(rows) {
  resultsBody.innerHTML = rows.map(r => {
    const stateClass = (r.reservationState || 'available').toLowerCase();
    const canRequest = stateClass === 'available';
    return `
      <tr>
        <td>${escapeHtml(r.colony)}</td>
        <td>${escapeHtml(r.physicalTag || '—')}</td>
        <td>${escapeHtml(r.sex)}</td>
        <td>${escapeHtml(r.ageWeeks)}</td>
        <td>${escapeHtml(r.mouseline)}</td>
        <td>${escapeHtml(r.genotype)}</td>
        <td>${escapeHtml(r.cageTag)}</td>
        <td>${escapeHtml(r.status)}</td>
        <td><span class="badge ${stateClass}">${escapeHtml(r.reservationState || 'Available')}</span></td>
        <td>${canRequest ? `<button data-row="${r.row}" class="reserveBtn">Request</button>` : ''}</td>
      </tr>`;
  }).join('');

  resultsBody.querySelectorAll('.reserveBtn').forEach(btn => {
    btn.addEventListener('click', () => openReserveDialog(Number(btn.dataset.row)));
  });
}

function escapeHtml(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

searchInput.addEventListener('input', debounce(runSearch, 350));
colonyFilter.addEventListener('change', runSearch);
refreshBtn.addEventListener('click', runSearch);

function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

// ---------- reservation ----------

function openReserveDialog(row) {
  reserveTargetRow = row;
  const record = currentRows.find(r => r.row === row);
  reserveSummary.textContent = record
    ? `${record.colony} — ${record.physicalTag || record.cageTag} (${record.genotype || ''})`
    : '';
  reserveError.hidden = true;
  reserveForm.reset();
  reserveDialog.showModal();
}

cancelReserve.addEventListener('click', () => reserveDialog.close());

reserveForm.addEventListener('submit', async (ev) => {
  ev.preventDefault();
  reserveError.hidden = true;
  try {
    await apiPost('reserve', {
      token: getToken(),
      row: reserveTargetRow,
      requesterName: document.getElementById('reqName').value.trim(),
      requesterEmail: document.getElementById('reqEmail').value.trim()
    });
    reserveDialog.close();
    statusMsg.textContent = 'Request sent — the colony owner has been emailed for approval.';
    runSearch();
  } catch (err) {
    reserveError.textContent = err.message;
    reserveError.hidden = false;
  }
});
