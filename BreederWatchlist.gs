// ============================================================
// BreederWatchlist.gs — flags breeders that have stalled
// (no litter in BREEDER_STALL_DAYS) or aged out
// (older than BREEDER_MAX_AGE_WEEKS), without anyone having
// to check dates by hand.
// ============================================================

function refreshBreederWatchlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const animalsSheet = ss.getSheetByName(SHEETS.ANIMALS);
  const watchlistSheet = ss.getSheetByName(SHEETS.BREEDER_WATCHLIST);
  if (!animalsSheet || !watchlistSheet) return;

  const hits = findBreederWatchlistHits_(animalsSheet);

  // clearContents() (not clear()) so this rebuild only wipes values, not
  // the sheet's font/background formatting — clear() would strip those
  // on every refresh (every edit, plus the nightly sweep).
  const headerRow = watchlistSheet.getRange(1, 1, 1, 8).getValues();
  const lastRow = watchlistSheet.getMaxRows();
  watchlistSheet.getRange(1, 1, lastRow, 8).clearContent();
  watchlistSheet.getRange(1, 1, 1, 8).setValues(headerRow).setFontWeight('bold');
  watchlistSheet.setFrozenRows(1);

  if (hits.length > 0) {
    watchlistSheet.getRange(2, 1, hits.length, 8).setValues(hits.map(h => [
      h.colony, h.physicalTag, h.cageTag, h.pairedDate, h.lastLitterDate,
      h.daysSinceLastLitter, h.ageWeeks, h.reason
    ]));
    const range = watchlistSheet.getRange(2, 1, hits.length, 8);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=TRUE').setBackground('#fff2cc').setRanges([range]).build()
    ];
    watchlistSheet.setConditionalFormatRules(rules);
  }
  watchlistSheet.autoResizeColumns(1, 8);
}

function findBreederWatchlistHits_(animalsSheet) {
  const data = animalsSheet.getDataRange().getValues();
  const hits = [];
  const now = new Date();

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[ANIMALS_COL.STATUS - 1] !== 'Breeder') continue;

    const lastLitterDate = row[ANIMALS_COL.DATE_OF_LAST_LITTER - 1];
    const ageWeeks = Number(row[ANIMALS_COL.AGE_WEEKS - 1]) || null;
    const pairedDate = row[ANIMALS_COL.BREEDER_PAIRED_DATE - 1];

    let daysSinceLastLitter = null;
    if (lastLitterDate) {
      daysSinceLastLitter = Math.floor((now - new Date(lastLitterDate)) / (1000 * 60 * 60 * 24));
    } else if (pairedDate) {
      daysSinceLastLitter = Math.floor((now - new Date(pairedDate)) / (1000 * 60 * 60 * 24));
    }

    const reasons = [];
    if (daysSinceLastLitter !== null && daysSinceLastLitter > BREEDER_STALL_DAYS) {
      reasons.push(`No litter in ${daysSinceLastLitter} days`);
    } else if (daysSinceLastLitter === null) {
      reasons.push('No litter on record yet');
    }
    if (ageWeeks !== null && ageWeeks > BREEDER_MAX_AGE_WEEKS) {
      reasons.push(`Age ${ageWeeks}w > ${BREEDER_MAX_AGE_WEEKS}w`);
    }
    if (reasons.length === 0) continue;

    hits.push({
      colony: row[ANIMALS_COL.COLONY - 1],
      physicalTag: row[ANIMALS_COL.PHYSICAL_TAG - 1],
      cageTag: row[ANIMALS_COL.CAGE_TAG - 1],
      pairedDate,
      lastLitterDate,
      daysSinceLastLitter: daysSinceLastLitter === null ? '' : daysSinceLastLitter,
      ageWeeks: ageWeeks === null ? '' : ageWeeks,
      reason: reasons.join('; ')
    });
  }
  return hits;
}

/** Weekly digest email — installed by Setup.gs to run Monday mornings. */
function weeklyBreederDigest() {
  refreshBreederWatchlist();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const watchlistSheet = ss.getSheetByName(SHEETS.BREEDER_WATCHLIST);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (!watchlistSheet || !usersSheet) return;

  const lastRow = watchlistSheet.getLastRow();
  if (lastRow < 2) return; // nothing to report

  const rows = watchlistSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const lines = rows.map(r => `- [${r[0]}] ${r[1] || '(no tag)'} in ${r[2]}: ${r[7]}`);
  const body = `Breeder watchlist (${rows.length} flagged):\n\n${lines.join('\n')}\n\nFull sheet: ${ss.getUrl()}`;

  const recipients = usersSheet.getRange(2, 1, Math.max(usersSheet.getLastRow() - 1, 0), 2).getValues()
    .map(([, email]) => email).filter(Boolean);
  if (recipients.length === 0) return;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: `Breeder watchlist: ${rows.length} flagged`,
    body
  });
}
