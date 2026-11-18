// ============================================================
// WebApp.gs — the dashboard's backend. Deploy this project as a
// Web App (Deploy > New deployment > Web app; execute as "Me",
// access "Anyone"). The GitHub Pages frontend calls the
// deployment URL; nothing here needs a separate database or
// hosting — it reads/writes the live Animals sheet directly.
//
// ONE-TIME SETUP (run each of these once from the script editor):
//   setDashboardPassword('choose-a-password')
//   ensureApprovalSecret_()
//
// Security note: this is a shared-password gate for an internal
// lab tool, not bank-grade auth. It keeps casual visitors out;
// it is not a substitute for not exposing anything truly
// sensitive on the dashboard.
// ============================================================

const PROP_PASSWORD_HASH = 'DASHBOARD_PASSWORD_HASH';
const PROP_APPROVAL_SECRET = 'APPROVAL_SECRET';
const SESSION_TTL_SECONDS = 6 * 60 * 60; // 6 hours

// ---------- one-time setup helpers ----------

function setDashboardPassword(plainPassword) {
  const hash = hashString_(plainPassword);
  PropertiesService.getScriptProperties().setProperty(PROP_PASSWORD_HASH, hash);
  Logger.log('Dashboard password set.');
}

function ensureApprovalSecret_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_APPROVAL_SECRET)) {
    props.setProperty(PROP_APPROVAL_SECRET, Utilities.getUuid() + Utilities.getUuid());
  }
}

function hashString_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ---------- entry points ----------

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    if (action === 'login') return jsonOutput_(handleLogin_(body));
    if (action === 'search') return jsonOutput_(handleSearch_(body));
    if (action === 'reserve') return jsonOutput_(handleReserve_(body));

    return jsonOutput_({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput_({ error: String(err) });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'approve' || action === 'deny') {
      return htmlOutput_(handleApproveOrDeny_(e.parameter, action));
    }
    return htmlOutput_('Mouse Colony Dashboard API is running.');
  } catch (err) {
    return htmlOutput_('Error: ' + String(err));
  }
}

// ---------- login / session ----------

function handleLogin_(body) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty(PROP_PASSWORD_HASH);
  if (!expected) return { error: 'No dashboard password has been set yet. Run setDashboardPassword() in the script editor.' };

  if (hashString_(body.password || '') !== expected) {
    return { error: 'Incorrect password.' };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session:' + token, 'ok', SESSION_TTL_SECONDS);
  return { token };
}

function requireSession_(token) {
  if (!token) return false;
  return CacheService.getScriptCache().get('session:' + token) === 'ok';
}

// ---------- search ----------

function handleSearch_(body) {
  if (!requireSession_(body.token)) return { error: 'Session expired, please log in again.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sharedColonies = loadSharedColonyNames_(ss);
  const sheet = ss.getSheetByName(SHEETS.ANIMALS);
  const data = sheet.getDataRange().getValues();

  const query = (body.query || '').toString().toLowerCase();
  const colonyFilter = body.colony || '';

  const results = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const colony = row[ANIMALS_COL.COLONY - 1];
    if (!sharedColonies.includes(colony)) continue;
    if (colonyFilter && colony !== colonyFilter) continue;
    if (['Used', 'Archived', 'Dead'].includes(row[ANIMALS_COL.STATUS - 1])) continue;

    const record = {
      row: r + 1,
      colony,
      physicalTag: row[ANIMALS_COL.PHYSICAL_TAG - 1],
      sex: row[ANIMALS_COL.SEX - 1],
      ageWeeks: row[ANIMALS_COL.AGE_WEEKS - 1],
      mouseline: row[ANIMALS_COL.MOUSELINE - 1],
      genotype: row[ANIMALS_COL.GENOTYPE - 1],
      cageTag: row[ANIMALS_COL.CAGE_TAG - 1],
      status: row[ANIMALS_COL.STATUS - 1],
      reservationState: row[ANIMALS_COL.RESERVATION_STATE - 1] || 'Available',
      requestBy: row[ANIMALS_COL.REQUEST_BY - 1]
    };

    if (query) {
      const haystack = Object.values(record).join(' ').toLowerCase();
      if (!haystack.includes(query)) continue;
    }
    results.push(record);
  }
  return { colonies: sharedColonies, results };
}

function loadSharedColonyNames_(ss) {
  const sheet = ss.getSheetByName(SHEETS.COLONIES);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => (r[COLONIES_COL.SHARED - 1] || '').toString().toUpperCase() === 'Y')
    .map(r => r[COLONIES_COL.NAME - 1]);
}

// ---------- reservation request ----------

function handleReserve_(body) {
  if (!requireSession_(body.token)) return { error: 'Session expired, please log in again.' };

  const row = Number(body.row);
  const requesterName = (body.requesterName || '').toString().trim();
  const requesterEmail = (body.requesterEmail || '').toString().trim();
  if (!row || !requesterName || !requesterEmail) {
    return { error: 'Missing row, requesterName, or requesterEmail.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ANIMALS);
  const colony = sheet.getRange(row, ANIMALS_COL.COLONY).getValue();
  const currentState = sheet.getRange(row, ANIMALS_COL.RESERVATION_STATE).getValue();
  if (currentState === 'Reserved' || currentState === 'Requested') {
    return { error: 'That animal is already ' + currentState.toLowerCase() + '.' };
  }

  sheet.getRange(row, ANIMALS_COL.REQUEST_BY).setValue(requesterName);
  sheet.getRange(row, ANIMALS_COL.REQUEST_DATE).setValue(new Date());
  updateReservationStateForRow_(sheet, row);

  sendApprovalEmail_(ss, sheet, row, colony, requesterName, requesterEmail);
  return { ok: true };
}

function sendApprovalEmail_(ss, sheet, row, colony, requesterName, requesterEmail) {
  ensureApprovalSecret_();
  const approverEmail = lookupApproverEmail_(ss, colony);
  if (!approverEmail) {
    Logger.log('No approver email configured for colony ' + colony + ' — request left Requested with no email sent.');
    return;
  }

  const physicalTag = sheet.getRange(row, ANIMALS_COL.PHYSICAL_TAG).getValue();
  const cageTag = sheet.getRange(row, ANIMALS_COL.CAGE_TAG).getValue();
  const scriptUrl = ScriptApp.getService().getUrl();

  const approveUrl = buildDecisionUrl_(scriptUrl, row, colony, 'approve', requesterEmail, requesterName);
  const denyUrl = buildDecisionUrl_(scriptUrl, row, colony, 'deny', requesterEmail, requesterName);

  const body = [
    `${requesterName} (${requesterEmail}) requested a mouse from ${colony}.`,
    ``,
    `Physical Tag: ${physicalTag || '(untagged)'}`,
    `Cage Tag: ${cageTag}`,
    ``,
    `Approve: ${approveUrl}`,
    `Deny: ${denyUrl}`
  ].join('\n');

  MailApp.sendEmail({ to: approverEmail, subject: `Reservation request: ${colony} — ${physicalTag || cageTag}`, body });
}

function lookupApproverEmail_(ss, colony) {
  const sheet = ss.getSheetByName(SHEETS.COLONIES);
  if (!sheet) return '';
  const rows = sheet.getDataRange().getValues().slice(1);
  const match = rows.find(r => r[COLONIES_COL.NAME - 1] === colony);
  return match ? match[COLONIES_COL.APPROVER_EMAIL - 1] : '';
}

// ---------- approve / deny (clicked from the email) ----------

function buildDecisionUrl_(scriptUrl, row, colony, decision, requesterEmail, requesterName) {
  const secret = PropertiesService.getScriptProperties().getProperty(PROP_APPROVAL_SECRET);
  const payload = `${row}|${colony}|${decision}`;
  const token = Utilities.computeHmacSha256Signature(payload, secret)
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

  const params = {
    action: decision,
    row: String(row),
    colony: encodeURIComponent(colony),
    token,
    requesterEmail: encodeURIComponent(requesterEmail),
    requesterName: encodeURIComponent(requesterName)
  };
  const query = Object.entries(params).map(([k, v]) => `${k}=${v}`).join('&');
  return `${scriptUrl}?${query}`;
}

function handleApproveOrDeny_(params, decision) {
  const row = Number(params.row);
  const colony = decodeURIComponent(params.colony || '');
  const requesterEmail = decodeURIComponent(params.requesterEmail || '');
  const requesterName = decodeURIComponent(params.requesterName || '');

  const secret = PropertiesService.getScriptProperties().getProperty(PROP_APPROVAL_SECRET);
  const payload = `${row}|${colony}|${decision}`;
  const expected = Utilities.computeHmacSha256Signature(payload, secret)
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  if (expected !== params.token) return 'This link is invalid or has already been used.';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ANIMALS);
  const rowColony = sheet.getRange(row, ANIMALS_COL.COLONY).getValue();
  if (rowColony !== colony) return 'This reservation no longer matches that row — it may have changed since the email was sent.';

  if (decision === 'approve') {
    sheet.getRange(row, ANIMALS_COL.USER).setValue(requesterName || sheet.getRange(row, ANIMALS_COL.REQUEST_BY).getValue());
  } else {
    sheet.getRange(row, ANIMALS_COL.REQUEST_BY).clearContent();
    sheet.getRange(row, ANIMALS_COL.REQUEST_DATE).clearContent();
  }
  updateReservationStateForRow_(sheet, row);

  if (requesterEmail) {
    const physicalTag = sheet.getRange(row, ANIMALS_COL.PHYSICAL_TAG).getValue();
    MailApp.sendEmail({
      to: requesterEmail,
      subject: `Your ${colony} reservation was ${decision === 'approve' ? 'approved' : 'denied'}`,
      body: `Your request for ${physicalTag || 'a mouse in ' + colony} was ${decision === 'approve' ? 'approved.' : 'denied.'}`
    });
  }

  return decision === 'approve' ? 'Approved. The requester has been emailed.' : 'Denied. The requester has been emailed.';
}

// ---------- output helpers ----------

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function htmlOutput_(message) {
  return HtmlService.createHtmlOutput(`<p style="font-family:sans-serif">${message}</p>`);
}
