// ============================================================
// Triggers.gs — the one onEdit handler for the whole workbook.
// Replaces the old per-colony processEdit(): one Animals sheet
// now carries every colony, so there's one state machine instead
// of three near-duplicate ones.
// ============================================================

const COLOR_RESERVED = '#f4cccc';
const COLOR_AVAILABLE = '#d9ead3';
const COLOR_REQUESTED = '#fff2cc';
const COLOR_DEFAULT = '#ffffff';

function onEditInstallable(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (![SHEETS.ANIMALS, SHEETS.LITTERS, SHEETS.USERS].includes(sheetName)) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    if (sheetName === SHEETS.USERS) {
      // Nothing to do — Animals' User column reads Users!A2:A200 live via
      // data validation, same spreadsheet, no sync required.
      return;
    }
    if (sheetName === SHEETS.LITTERS) {
      handleLittersEdit_(e, sheet);
      return;
    }
    if (sheetName === SHEETS.ANIMALS) {
      handleAnimalsEdit_(e, sheet);
      return;
    }
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// ANIMALS — archiving + reservation state machine
// ------------------------------------------------------------

function handleAnimalsEdit_(e, sheet) {
  const startCol = e.range.getColumn();
  const numCols = e.range.getNumColumns();
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  let touchedCoreColumns = startCol <= ANIMALS_COL.STATUS; // edit landed in columns 1..Status
  let madeChanges = false;

  const colsThatDriveReservation = [ANIMALS_COL.CAGE_TAG, ANIMALS_COL.STATUS, ANIMALS_COL.REQUEST_BY, ANIMALS_COL.USER, ANIMALS_COL.RESERVATION_STATE];
  const editTouches = col => startCol <= col && startCol + numCols - 1 >= col;
  const touchesReservationInputs = colsThatDriveReservation.some(editTouches);

  if (touchesReservationInputs) {
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row === 1) continue;
      if (updateReservationStateForRow_(sheet, row)) madeChanges = true;
    }
  }

  if (archiveFinishedRows_(sheet)) madeChanges = true;

  if (touchedCoreColumns || madeChanges) {
    refreshCageChartAll();
  }
  const touchesBreederFields = [ANIMALS_COL.STATUS, ANIMALS_COL.BREEDER_PAIRED_DATE, ANIMALS_COL.NUMBER_OF_LITTERS, ANIMALS_COL.DATE_OF_LAST_LITTER].some(editTouches);
  if (touchesBreederFields) {
    refreshBreederWatchlist();
  }
}

/** Returns true if this row's reservation state/background changed. */
function updateReservationStateForRow_(sheet, row) {
  const cageTag = sheet.getRange(row, ANIMALS_COL.CAGE_TAG).getValue();
  const stateCell = sheet.getRange(row, ANIMALS_COL.RESERVATION_STATE);
  let currentState = stateCell.getValue();

  if (cageTag === '') {
    if (currentState !== '') {
      stateCell.clearContent();
      stateCell.setBackground(COLOR_DEFAULT);
      return true;
    }
    return false;
  }

  if (currentState === '') {
    stateCell.setValue('Available');
    stateCell.setBackground(COLOR_AVAILABLE);
    currentState = 'Available';
  }

  const status = sheet.getRange(row, ANIMALS_COL.STATUS).getValue();
  if (['Used', 'Archived', 'Dead'].includes(status)) return false; // leaving for archiveFinishedRows_

  const requestDateCell = sheet.getRange(row, ANIMALS_COL.REQUEST_DATE);
  const requestByCell = sheet.getRange(row, ANIMALS_COL.REQUEST_BY);
  const userCell = sheet.getRange(row, ANIMALS_COL.USER);
  const approvedCell = sheet.getRange(row, ANIMALS_COL.APPROVED);
  const approvedDateCell = sheet.getRange(row, ANIMALS_COL.APPROVED_DATE);

  const requestByVal = requestByCell.getValue();
  const userVal = userCell.getValue();
  const before = [currentState, stateCell.getBackground()].join('|');

  if (userVal) {
    stateCell.setValue('Reserved');
    stateCell.setBackground(COLOR_RESERVED);
    if (!approvedDateCell.getValue()) approvedDateCell.setValue(new Date());
    if (!approvedCell.getValue()) approvedCell.setValue('Approved');
  } else if (requestByVal) {
    stateCell.setValue('Requested');
    stateCell.setBackground(COLOR_REQUESTED);
    if (!requestDateCell.getValue()) requestDateCell.setValue(new Date());
    approvedDateCell.clearContent();
    approvedCell.clearContent();
  } else {
    stateCell.setValue('Available');
    stateCell.setBackground(COLOR_AVAILABLE);
    requestDateCell.clearContent();
    approvedDateCell.clearContent();
    approvedCell.clearContent();
  }

  const after = [stateCell.getValue(), stateCell.getBackground()].join('|');
  return before !== after;
}

/** Moves every Used/Archived/Dead row to Archive, bottom-up so row numbers stay valid. Returns true if anything moved. */
function archiveFinishedRows_(sheet) {
  const allData = sheet.getDataRange().getValues();
  const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ARCHIVE);
  let moved = false;

  for (let r = allData.length - 1; r > 0; r--) {
    const rowState = allData[r][ANIMALS_COL.STATUS - 1];
    if (['Used', 'Archived', 'Dead'].includes(rowState)) {
      const rowData = allData[r].slice();
      rowData.push(new Date());
      archiveSheet.appendRow(rowData);
      sheet.deleteRow(r + 1);
      moved = true;
    }
  }
  return moved;
}

// ------------------------------------------------------------
// LITTERS — wean-date default, breeder stat rollup
// ------------------------------------------------------------

function handleLittersEdit_(e, sheet) {
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  if (startRow === 1 && numRows === 1) return; // header-only edit

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue;
    fillWeanDateDefault_(sheet, row);
  }

  updateBreederStatsFromLitters_();
  refreshCageChartAll();
  refreshBreederWatchlist();
}

function fillWeanDateDefault_(sheet, row) {
  const dob = sheet.getRange(row, LITTERS_COL.DOB).getValue();
  if (!dob) return;
  const weanCell = sheet.getRange(row, LITTERS_COL.WEAN_DATE);
  if (!weanCell.getValue()) {
    const wean = new Date(dob);
    wean.setDate(wean.getDate() + WEAN_DAYS_DEFAULT);
    weanCell.setValue(wean);
  }
  const statusCell = sheet.getRange(row, LITTERS_COL.STATUS);
  if (!statusCell.getValue()) statusCell.setValue('Active');

  // Auto-flip to Weaned once the wean date has passed, so the Active/Weaned
  // flag doesn't rely on someone remembering to update it by hand.
  const weanDate = weanCell.getValue();
  if (weanDate && new Date() >= new Date(weanDate) && statusCell.getValue() !== 'Weaned') {
    statusCell.setValue('Weaned');
  }
}

/**
 * Recomputes Number of Litters / Date of Last Litter on Animals from the
 * Litters sheet. Fully recomputed (not incremented) each time, so edits
 * and deletions in Litters can never double-count.
 *
 * Matching rule: if a litter row has a Dam Tag, it credits that specific
 * animal (harem cages — the cage tag alone doesn't say which female).
 * Otherwise it credits every Breeder-status animal whose Cage Tag matches
 * the litter's Breeding Cage Tag (a fixed single pair).
 */
function updateBreederStatsFromLitters_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const littersSheet = ss.getSheetByName(SHEETS.LITTERS);
  const animalsSheet = ss.getSheetByName(SHEETS.ANIMALS);
  if (!littersSheet || !animalsSheet) return;

  const litters = littersSheet.getDataRange().getValues().slice(1)
    .filter(r => r[LITTERS_COL.COLONY - 1] && r[LITTERS_COL.BREEDING_CAGE_TAG - 1] && r[LITTERS_COL.DOB - 1]);

  const animalsData = animalsSheet.getDataRange().getValues();

  for (let r = 1; r < animalsData.length; r++) {
    const row = animalsData[r];
    if (row[ANIMALS_COL.STATUS - 1] !== 'Breeder') continue;

    const colony = row[ANIMALS_COL.COLONY - 1];
    const cageTag = row[ANIMALS_COL.CAGE_TAG - 1];
    const physicalTag = row[ANIMALS_COL.PHYSICAL_TAG - 1];

    const matching = litters.filter(l => {
      if (l[LITTERS_COL.COLONY - 1] !== colony || l[LITTERS_COL.BREEDING_CAGE_TAG - 1] !== cageTag) return false;
      const damTag = l[LITTERS_COL.DAM_TAG - 1];
      return damTag ? damTag === physicalTag : true;
    });

    if (matching.length === 0) continue;

    const dates = matching.map(l => new Date(l[LITTERS_COL.DOB - 1]).getTime());
    const lastLitterDate = new Date(Math.max.apply(null, dates));

    const sheetRow = r + 1;
    animalsSheet.getRange(sheetRow, ANIMALS_COL.NUMBER_OF_LITTERS).setValue(matching.length);
    animalsSheet.getRange(sheetRow, ANIMALS_COL.DATE_OF_LAST_LITTER).setValue(lastLitterDate);
  }
}

// ------------------------------------------------------------
// Daily / weekly maintenance (installed by Setup.gs)
// ------------------------------------------------------------

/** Safety net for bulk edits (paste, import) that skip onEdit entirely. */
function dailyMaintenance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const animalsSheet = ss.getSheetByName(SHEETS.ANIMALS);
  const littersSheet = ss.getSheetByName(SHEETS.LITTERS);

  if (animalsSheet) {
    const lastRow = animalsSheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) {
      updateReservationStateForRow_(animalsSheet, row);
    }
    archiveFinishedRows_(animalsSheet);
  }

  if (littersSheet) {
    const lastRow = littersSheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) {
      fillWeanDateDefault_(littersSheet, row);
    }
    updateBreederStatsFromLitters_();
  }

  refreshCageChartAll();
  refreshBreederWatchlist();
}
