// ============================================================
// CONFIG.gs — sheet names, column layout, dorm geometry.
// Edit this file first when colonies, dorms or columns change;
// everything else reads from these constants.
// ============================================================

const SHEETS = {
  ANIMALS: 'Animals',
  ARCHIVE: 'Archive',
  LITTERS: 'Litters',
  USERS: 'Users',
  COLONIES: 'Colonies',
  GENOTYPE_COLORS: 'Genotype_Colors',
  CAGE_CHART: 'Cage_Chart',
  BREEDER_WATCHLIST: 'Breeder_Watchlist'
};

// --- Animals sheet — 1-indexed columns ---
const ANIMALS_COL = {
  COLONY: 1,
  PHYSICAL_TAG: 2,
  SEX: 3,
  DOB: 4,
  AGE_WEEKS: 5,
  AGE_MONTHS: 6,
  MOUSELINE: 7,
  GENOTYPE: 8,
  CAGE_TAG: 9,
  FATHER_TAG: 10,
  MOTHER_TAG: 11,
  COMMENT: 12,
  STATUS: 13,             // Active / Breeder / Used / Archived / Dead
  RESERVATION_STATE: 14,  // Available / Requested / Reserved — shared colonies only
  REQUEST_DATE: 15,
  REQUEST_BY: 16,
  USER: 17,
  APPROVED: 18,
  APPROVED_DATE: 19,
  BREEDER_PAIRED_DATE: 20,
  NUMBER_OF_LITTERS: 21,
  DATE_OF_LAST_LITTER: 22
};

const ANIMALS_HEADERS = [
  'Colony', 'Physical Tag', 'Sex', 'DOB', 'Age (weeks)', 'Age (months)',
  'Mouseline', 'Genotype', 'Cage Tag', 'Father Tag', 'Mother Tag', 'Comment',
  'Status', 'Reservation State', 'Request Date', 'Request By', 'User',
  'Approved', 'Approved Date', 'Breeder Paired Date', 'Number of Litters',
  'Date of Last Litter'
];

// --- Litters sheet ---
const LITTERS_COL = {
  LITTER_ID: 1,
  COLONY: 2,
  BREEDING_CAGE_TAG: 3,
  DAM_TAG: 4,   // optional — only needed for harem cages (one male, several females)
  DOB: 5,
  WEAN_DATE: 6, // auto-filled DOB+21 if left blank; you can overwrite per litter
  STATUS: 7,    // Active / Weaned
  NOTES: 8
};

const LITTERS_HEADERS = [
  'Litter ID', 'Colony', 'Breeding Cage Tag', 'Dam Tag', 'DOB',
  'Expected Wean Date', 'Status', 'Notes'
];

const WEAN_DAYS_DEFAULT = 21;

// --- Users sheet (reservation dropdown + email for notifications) ---
const USERS_HEADERS = ['Name', 'Email'];

// --- Colonies config sheet ---
const COLONIES_COL = { NAME: 1, PRIMARY_DORM: 2, SHARED: 3, APPROVER_EMAIL: 4 };
const COLONIES_HEADERS = ['Colony Name', 'Primary Dorm (reference)', 'Shared? (Y/N)', 'Approver Email'];

// --- Genotype_Colors config sheet ---
const GENOTYPE_COLORS_COL = { GENOTYPE: 1, COLOR: 2 };
const GENOTYPE_COLORS_HEADERS = ['Genotype', 'Color (hex)'];

// --- Breeder alert thresholds ---
const BREEDER_STALL_DAYS = 90;
const BREEDER_MAX_AGE_WEEKS = 65;

// --- Default font applied to every managed sheet ---
const DEFAULT_FONT_FAMILY = 'Lato';
const DEFAULT_FONT_SIZE = 14;

// --- Dorm geometry ---
// Letter = row (A–I, 9 rows down). Number = column (1–7, 7 columns across).
// Cage Tag format: "<Dorm>-<RowLetter><ColNumber>", e.g. "1R-A3".
const DORM_NAMES = ['1R', '4R', '6R', '7R'];
const ROWS_PER_DORM = 9;       // A–I
const COLS_PER_DORM = 7;       // 1–7
const ANIMALS_PER_CAGE = 5;
const ROWS_PER_CAGE = 1 + 1 + ANIMALS_PER_CAGE; // header row + notes row + 5 animal slots
