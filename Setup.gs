// ============================================================
// Setup.gs — run ONCE (or whenever you want to re-check/repair
// the sheet structure). Safe to re-run: it only creates what's
// missing and fixes headers/validation, it never touches your
// animal rows.
//
// From the Apps Script editor: select setupAll in the function
// dropdown and click Run. Approve the permissions prompt.
// ============================================================

function setupAll() {
  setupAnimals_();
  setupArchive_();
  setupLitters_();
  setupUsers_();
  setupColonies_();
  setupGenotypeColors_();
  setupBreederWatchlist_();
  ensureCageChartSheetExists_();
  installTriggers_();
  applyDefaultFontToAllManagedSheets_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Setup complete. See the README for next steps (Colonies/Genotype_Colors rows, Web App deploy).', 'Mouse Colony System', 10);
}

/**
 * Hardcodes Lato 14pt across every sheet this system manages (Google's
 * theme picker only offers a handful of preset font/size combos, no
 * custom size, so this is set directly on the cells instead). Safe to
 * re-run any time — e.g. after adding more rows/columns by hand — since
 * it just re-applies the same font, it doesn't touch bold/background/
 * borders.
 */
function applyDefaultFontToAllManagedSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) applyDefaultFont_(sheet);
  });
}

function applyDefaultFont_(sheet) {
  const maxRows = Math.max(sheet.getMaxRows(), 1);
  const maxCols = Math.max(sheet.getMaxColumns(), 1);
  sheet.getRange(1, 1, maxRows, maxCols)
    .setFontFamily(DEFAULT_FONT_FAMILY)
    .setFontSize(DEFAULT_FONT_SIZE);
}

// ---------- individual sheet builders ----------

function setupAnimals_() {
  const sheet = getOrCreateSheet_(SHEETS.ANIMALS);
  writeHeaderRowIfEmpty_(sheet, ANIMALS_HEADERS);

  // Colony dropdown reads from Colonies!A2:A (so the list of colonies lives in one place)
  const coloniesSheet = getOrCreateSheet_(SHEETS.COLONIES);
  applyListValidation_(sheet, ANIMALS_COL.COLONY, 2, 2000,
    SpreadsheetApp.newDataValidation().requireValueInRange(coloniesSheet.getRange('A2:A200'), true).setAllowInvalid(false).build());

  applyListValidation_(sheet, ANIMALS_COL.SEX, 2, 2000,
    SpreadsheetApp.newDataValidation().requireValueInList(['M', 'F'], true).setAllowInvalid(false).build());

  applyListValidation_(sheet, ANIMALS_COL.STATUS, 2, 2000,
    SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Breeder', 'Used', 'Archived', 'Dead'], true).setAllowInvalid(false).build());

  applyListValidation_(sheet, ANIMALS_COL.RESERVATION_STATE, 2, 2000,
    SpreadsheetApp.newDataValidation().requireValueInList(['Available', 'Requested', 'Reserved'], true).setAllowInvalid(true).build());

  const usersSheet = getOrCreateSheet_(SHEETS.USERS);
  applyListValidation_(sheet, ANIMALS_COL.USER, 2, 2000,
    SpreadsheetApp.newDataValidation().requireValueInRange(usersSheet.getRange('A2:A200'), true).setAllowInvalid(true).build());

  // Age (weeks)/(months) are formulas off DOB — seed them down a generous range
  // so they're live without anyone having to copy the formula down by hand.
  seedAgeFormulas_(sheet, 2, 2000);

  // Reservation-state background colors, same palette the old Aging sheet used.
  applyReservationStateFormatting_(sheet);

  sheet.setFrozenRows(1);
}

function seedAgeFormulas_(sheet, firstRow, lastRow) {
  const dobCol = colLetter_(ANIMALS_COL.DOB);
  const weeksCol = ANIMALS_COL.AGE_WEEKS;
  const monthsCol = ANIMALS_COL.AGE_MONTHS;
  const weeksRange = sheet.getRange(firstRow, weeksCol, lastRow - firstRow + 1, 1);
  const monthsRange = sheet.getRange(firstRow, monthsCol, lastRow - firstRow + 1, 1);
  const weeksFormulas = [];
  const monthsFormulas = [];
  for (let r = firstRow; r <= lastRow; r++) {
    weeksFormulas.push([`=IF(${dobCol}${r}="","",ROUND((TODAY()-${dobCol}${r})/7,1))`]);
    monthsFormulas.push([`=IF(${dobCol}${r}="","",ROUND((TODAY()-${dobCol}${r})/30.44,1))`]);
  }
  weeksRange.setFormulas(weeksFormulas);
  monthsRange.setFormulas(monthsFormulas);
}

function setupArchive_() {
  const sheet = getOrCreateSheet_(SHEETS.ARCHIVE);
  writeHeaderRowIfEmpty_(sheet, ANIMALS_HEADERS.concat(['Archived Date']));
  sheet.setFrozenRows(1);
}

function setupLitters_() {
  const sheet = getOrCreateSheet_(SHEETS.LITTERS);
  writeHeaderRowIfEmpty_(sheet, LITTERS_HEADERS);

  const coloniesSheet = getOrCreateSheet_(SHEETS.COLONIES);
  applyListValidation_(sheet, LITTERS_COL.COLONY, 2, 500,
    SpreadsheetApp.newDataValidation().requireValueInRange(coloniesSheet.getRange('A2:A200'), true).setAllowInvalid(false).build());

  applyListValidation_(sheet, LITTERS_COL.STATUS, 2, 500,
    SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Weaned'], true).setAllowInvalid(false).build());

  sheet.setFrozenRows(1);
}

function setupUsers_() {
  const sheet = getOrCreateSheet_(SHEETS.USERS);
  writeHeaderRowIfEmpty_(sheet, USERS_HEADERS);
  sheet.setFrozenRows(1);
}

function setupColonies_() {
  const sheet = getOrCreateSheet_(SHEETS.COLONIES);
  const isNew = writeHeaderRowIfEmpty_(sheet, COLONIES_HEADERS);
  if (isNew) {
    // Seed with your six colonies so you're not typing dorm names from scratch.
    // Edit freely — this is just a starting point. Shared? is Y/N (plain text),
    // kept as text rather than a checkbox so it reads cleanly from a script.
    const seedRows = [
      ['Aging', '1R', 'Y', ''],
      ['Pham', '1R', 'Y', ''],
      ['Cas9-Cre', '4R + 7R', 'N', ''],
      ['Cas9 whole body', '6R', 'N', ''],
      ['Aldh1b1 fl/fl', '6R', 'N', ''],
      ['LysMCre', '6R', 'N', ''],
      ['Cd63 fl/fl', '6R', 'N', '']
    ];
    sheet.getRange(2, 1, seedRows.length, 4).setValues(seedRows);
    applyListValidation_(sheet, COLONIES_COL.SHARED, 2, 200,
      SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).setAllowInvalid(false).build());
  }
  sheet.setFrozenRows(1);
}

function setupGenotypeColors_() {
  const sheet = getOrCreateSheet_(SHEETS.GENOTYPE_COLORS);
  const isNew = writeHeaderRowIfEmpty_(sheet, GENOTYPE_COLORS_HEADERS);
  if (isNew) {
    // Carried over from the old refreshCageChart script's genotypeColors map.
    // Add a row per genotype you use — anything not listed here falls back to white.
    sheet.getRange(2, 1, 2, 2).setValues([
      ['fl/fl', '#cfe2f3'],
      ['fl/fl; cre/+', '#ead1dc']
    ]);
  }
  sheet.setFrozenRows(1);
}

function setupBreederWatchlist_() {
  const sheet = getOrCreateSheet_(SHEETS.BREEDER_WATCHLIST);
  writeHeaderRowIfEmpty_(sheet, [
    'Colony', 'Physical Tag', 'Cage Tag', 'Breeder Paired Date',
    'Date of Last Litter', 'Days Since Last Litter', 'Age (weeks)', 'Reason'
  ]);
  sheet.setFrozenRows(1);
}

function ensureCageChartSheetExists_() {
  // refreshCageChartAll() builds the actual content; this just makes sure
  // the tab exists so the first run of setupAll() has somewhere to draw into.
  getOrCreateSheet_(SHEETS.CAGE_CHART);
}

function installTriggers_() {
  // Remove any previously-installed triggers for these functions so re-running
  // setupAll() never leaves duplicates behind.
  const handlerNames = ['onEditInstallable', 'weeklyBreederDigest', 'dailyMaintenance'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  // Catches bulk pastes/imports that skip onEdit entirely.
  ScriptApp.newTrigger('dailyMaintenance').timeBased().everyDays(1).atHour(5).create();
  // Monday mornings — adjust the hour/day to taste.
  ScriptApp.newTrigger('weeklyBreederDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

// ---------- small shared helpers ----------

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Writes the header row only if row 1 is currently empty. Returns true if it wrote (i.e. the sheet was new). */
function writeHeaderRowIfEmpty_(sheet, headers) {
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '') return false;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function applyListValidation_(sheet, col, firstRow, lastRow, rule) {
  sheet.getRange(firstRow, col, lastRow - firstRow + 1, 1).setDataValidation(rule);
}

function applyReservationStateFormatting_(sheet) {
  const range = sheet.getRange(2, ANIMALS_COL.RESERVATION_STATE, 1999, 1);
  const rules = sheet.getConditionalFormatRules().filter(r =>
    !r.getRanges().some(rr => rr.getColumn() === ANIMALS_COL.RESERVATION_STATE)
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Available').setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Requested').setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Reserved').setBackground('#f4cccc').setRanges([range]).build()
  );
  sheet.setConditionalFormatRules(rules);
}

/** 1 -> 'A', 2 -> 'B', ... 27 -> 'AA' */
function colLetter_(col) {
  let letter = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
