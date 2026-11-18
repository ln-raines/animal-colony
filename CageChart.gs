// ============================================================
// CageChart.gs — one combined Cage_Chart tab across all four
// dorms. Colony "zones" are computed live from Animals.Colony,
// not a fixed range, so colonies can grow or shrink into cages
// without touching this script or any config.
//
// Grid convention: the letter names the row (A-I, 9 rows down),
// the number names the column (1-7, 7 columns across). Cage Tag
// format is "<Dorm>-<RowLetter><ColNumber>", e.g. "1R-A3".
// ============================================================

const GAP_COLS = 1; // blank columns between dorm blocks
const TITLE_ROW = 1;
const COL_NUMBER_ROW = 2;
const FIRST_BLOCK_ROW = 3;

const PALETTE = ['#cfe2f3', '#ead1dc', '#d9ead3', '#fff2cc', '#f4cccc', '#d0e0e3', '#e6cff2', '#fce5cd'];
const LITTER_NOTE_WINDOW_DAYS = 45; // how long a litter keeps showing on the chart after its DOB

function refreshCageChartAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cageSheet = ss.getSheetByName(SHEETS.CAGE_CHART);
  if (!cageSheet) return;

  const animalsByCageTag = loadActiveAnimalsByCageTag_();
  const littersByCageKey = loadRecentLittersByCageKey_();
  const colonyColor = buildColonyColorMap_();
  const genotypeColor = buildGenotypeColorMap_();

  DORM_NAMES.forEach((dorm, dormIdx) => {
    const blockStartCol = 1 + dormIdx * (COLS_PER_DORM + GAP_COLS);
    writeDormTitle_(cageSheet, dorm, blockStartCol);
    writeColumnNumberRow_(cageSheet, blockStartCol);

    for (let r = 0; r < ROWS_PER_DORM; r++) {
      const rowLetter = String.fromCharCode(65 + r);
      const blockTop = FIRST_BLOCK_ROW + r * ROWS_PER_CAGE;

      for (let c = 0; c < COLS_PER_DORM; c++) {
        const colNumber = c + 1;
        const sheetCol = blockStartCol + c;
        const cageTag = `${dorm}-${rowLetter}${colNumber}`;
        const occupants = animalsByCageTag[cageTag] || [];
        const colony = occupants.length > 0 ? occupants[0].colony : '';

        writeCageHeader_(cageSheet, blockTop, sheetCol, rowLetter + colNumber, colony, colonyColor);
        writeCageNotes_(cageSheet, blockTop + 1, sheetCol, colony, cageTag, littersByCageKey);
        writeCageAnimalSlots_(cageSheet, blockTop + 2, sheetCol, occupants, genotypeColor);

        cageSheet.getRange(blockTop, sheetCol, ROWS_PER_CAGE, 1)
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment('center');
      }
    }
  });

  const totalCols = DORM_NAMES.length * COLS_PER_DORM + (DORM_NAMES.length - 1) * GAP_COLS;
  cageSheet.autoResizeColumns(1, totalCols);
}

// ---------- data loading ----------

function loadActiveAnimalsByCageTag_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ANIMALS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const byCage = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const status = row[ANIMALS_COL.STATUS - 1];
    const cageTag = row[ANIMALS_COL.CAGE_TAG - 1];
    if (!cageTag || ['Used', 'Archived', 'Dead'].includes(status)) continue;

    const occupant = {
      colony: row[ANIMALS_COL.COLONY - 1],
      physicalTag: row[ANIMALS_COL.PHYSICAL_TAG - 1],
      sex: row[ANIMALS_COL.SEX - 1],
      genotype: row[ANIMALS_COL.GENOTYPE - 1],
      ageWeeks: row[ANIMALS_COL.AGE_WEEKS - 1],
      ageMonths: row[ANIMALS_COL.AGE_MONTHS - 1]
    };
    (byCage[cageTag] = byCage[cageTag] || []).push(occupant);
  }
  return byCage;
}

/** Most recent litter per colony+cageTag, whether Active or Weaned, within the display window. */
function loadRecentLittersByCageKey_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LITTERS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const byKey = {};
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - LITTER_NOTE_WINDOW_DAYS);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const colony = row[LITTERS_COL.COLONY - 1];
    const cageTag = row[LITTERS_COL.BREEDING_CAGE_TAG - 1];
    const dob = row[LITTERS_COL.DOB - 1];
    if (!colony || !cageTag || !dob) continue;
    if (new Date(dob) < cutoff) continue;

    const key = colony + '|' + cageTag;
    const entry = { dob, weanDate: row[LITTERS_COL.WEAN_DATE - 1], status: row[LITTERS_COL.STATUS - 1] };
    if (!byKey[key] || new Date(entry.dob) > new Date(byKey[key].dob)) byKey[key] = entry;
  }
  return byKey;
}

function buildColonyColorMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.COLONIES);
  const map = {};
  if (!sheet) return map;
  const names = sheet.getRange(2, COLONIES_COL.NAME, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  names.forEach((row, i) => {
    const name = row[0];
    if (name) map[name] = PALETTE[i % PALETTE.length];
  });
  return map;
}

function buildGenotypeColorMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.GENOTYPE_COLORS);
  const map = {};
  if (!sheet) return map;
  const rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  rows.forEach(([genotype, color]) => {
    if (genotype) map[genotype] = color || '#ffffff';
  });
  return map;
}

// ---------- drawing ----------

function writeDormTitle_(sheet, dorm, blockStartCol) {
  const range = sheet.getRange(TITLE_ROW, blockStartCol, 1, COLS_PER_DORM);
  range.breakApart();
  range.merge()
    .setValue('Dorm ' + dorm)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setHorizontalAlignment('center');
}

function writeColumnNumberRow_(sheet, blockStartCol) {
  for (let c = 0; c < COLS_PER_DORM; c++) {
    sheet.getRange(COL_NUMBER_ROW, blockStartCol + c)
      .setValue(c + 1)
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground('#f3f3f3')
      .setHorizontalAlignment('center');
  }
}

function writeCageHeader_(sheet, row, col, label, colony, colonyColorMap) {
  sheet.getRange(row, col)
    .setValue(colony ? `${label} — ${colony}` : label)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(colony ? (colonyColorMap[colony] || '#eeeeee') : '#ffffff');
}

/** Leaves hand-typed notes alone unless there's a litter to report on this cage. */
function writeCageNotes_(sheet, row, col, colony, cageTag, littersByCageKey) {
  const litter = littersByCageKey[colony + '|' + cageTag];
  if (!litter) return;

  const tz = Session.getScriptTimeZone();
  const dobStr = Utilities.formatDate(new Date(litter.dob), tz, 'MM/dd');
  const notesCell = sheet.getRange(row, col);

  if (litter.weanDate && (litter.status === 'Weaned' || new Date() >= new Date(litter.weanDate))) {
    const weanStr = Utilities.formatDate(new Date(litter.weanDate), tz, 'MM/dd');
    notesCell.setValue(`Weaned ${weanStr} — pups moved out`);
    notesCell.setBackground('#d9ead3');
  } else {
    const weanStr = litter.weanDate ? Utilities.formatDate(new Date(litter.weanDate), tz, 'MM/dd') : '?';
    notesCell.setValue(`Pups born ${dobStr} — wean by ${weanStr}`);
    notesCell.setBackground('#fff2cc');
  }
}

function writeCageAnimalSlots_(sheet, firstRow, col, occupants, genotypeColorMap) {
  for (let i = 0; i < ANIMALS_PER_CAGE; i++) {
    const cell = sheet.getRange(firstRow + i, col);
    const occupant = occupants[i];
    if (!occupant) {
      cell.setValue('').setBackground('#ffffff');
      continue;
    }
    const ageStr = formatAge_(occupant.ageWeeks, occupant.ageMonths);
    cell.setValue(`${occupant.sex || ''} | ${occupant.physicalTag || '—'} | ${occupant.genotype || ''}${ageStr}`);
    cell.setBackground(genotypeColorMap[occupant.genotype] || '#ffffff');
  }
}

/** Whole-number weeks and months, e.g. " | 76w (18mo)". Blank DOB -> no age shown at all. */
function formatAge_(weeksVal, monthsVal) {
  if (weeksVal === '' || weeksVal === null || weeksVal === undefined) return '';
  const weeksInt = Math.round(Number(weeksVal));
  const hasMonths = !(monthsVal === '' || monthsVal === null || monthsVal === undefined);
  const monthsInt = hasMonths ? Math.round(Number(monthsVal)) : null;
  return monthsInt === null ? ` | ${weeksInt}w` : ` | ${weeksInt}w (${monthsInt}mo)`;
}
