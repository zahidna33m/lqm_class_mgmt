/**
 * Book Column Migration — LQM Al-Falah 26
 *
 * One-time script that adds the "Book" column (column E, before "Lesson #") to
 * every group HW_Submissions sheet and to each aggregate sheet, and backfills
 * the Book value on existing rows.
 *
 * Lives in the same Apps Script project as HW_Submission_Aggregator.gs and uses
 * its constants (MASTER_SHEET_ID, AGG_TAB_NAME, SRC_TAB_NAME) and extractSheetId(),
 * and Books_Config.gs for the Book_HW tab. Run it manually from the editor; it has
 * no trigger.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 *   1. Disable the intake and aggregator triggers.
 *   2. Run migrateBookColumnDryRun() and read View → Logs. Nothing is changed.
 *   3. Run migrateBookColumn() to apply. Re-running is safe: sheets that already
 *      have the column are only checked for blank Book cells.
 *
 * ── What it does to each tab ─────────────────────────────────────────────────
 *   Old layout (header starts Submission ID … Student Name, Lesson #, …):
 *     inserts a column before Lesson #, titles it "Book", backfills every row.
 *   Already has "Book" in column E:
 *     fills only blank Book cells; existing values are never overwritten.
 *   Any other header:
 *     skipped and reported, never modified.
 *
 * ── Backfill rule ────────────────────────────────────────────────────────────
 *   Each row gets the book that resolveBook() gives its Timestamp (column B), using
 *   the Book_HW tab; a missing or non-date Timestamp gets Book 1.
 *   Rows with nothing in columns A–D are left blank.
 */

const BOOK_COL = 5; // column E, 1-based

// Header cells A–J of the layout before the Book column existed.
const PRE_BOOK_HEADERS = [
  'Submission ID', 'Timestamp', 'Student ID', 'Student Name', 'Lesson #',
  'HW File Link', 'Grader', 'Grading Status', 'Grade', 'Comments',
];

function migrateBookColumnDryRun() {
  migrateBookColumn_(true);
}

function migrateBookColumn() {
  migrateBookColumn_(false);
}

// ─── Core ──────────────────────────────────────────────────────────────────────

function migrateBookColumn_(dryRun) {
  Logger.log(dryRun ? '=== DRY RUN — nothing will be changed ===' : '=== MIGRATING ===');

  const books = loadBooks(SpreadsheetApp.openById(MASTER_SHEET_ID));
  const targets = []; // { label, spreadsheetId, tabName }

  // Group sheets, de-duplicated by spreadsheet ID.
  const seen = new Set();
  const groupRows = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('Groups').getDataRange().getValues();
  for (let i = 1; i < groupRows.length; i++) {
    const groupId   = String(groupRows[i][0] || '').trim();
    const sheetLink = String(groupRows[i][6] || '').trim();
    if (!groupId || !sheetLink) continue;
    const id = extractSheetId(sheetLink);
    if (!id) {
      Logger.log(`Group ${groupId}: cannot parse sheet ID from "${sheetLink}". Skipped.`);
      continue;
    }
    if (seen.has(id)) continue;
    seen.add(id);
    targets.push({ label: `Group ${groupId}`, spreadsheetId: id, tabName: SRC_TAB_NAME });
  }

  // Aggregates.
  for (const b of books) {
    if (!b.aggSheetId) {
      Logger.log(`Book ${b.book}: no Aggregate Sheet Link in ${BOOKS_TAB}. Skipped.`);
      continue;
    }
    targets.push({ label: `Book ${b.book} aggregate`, spreadsheetId: b.aggSheetId, tabName: AGG_TAB_NAME });
  }

  const summary = { inserted: 0, backfilled: 0, unchanged: 0, skipped: 0 };
  for (const t of targets) {
    let result;
    try {
      const tab = SpreadsheetApp.openById(t.spreadsheetId).getSheetByName(t.tabName);
      result = tab ? migrateTab_(tab, dryRun, books) : { status: 'skipped', note: `tab "${t.tabName}" not found` };
    } catch (e) {
      result = { status: 'skipped', note: `cannot open: ${e}` };
    }
    summary[result.status]++;
    Logger.log(`${t.label}: ${result.status.toUpperCase()} — ${result.note}`);
  }

  Logger.log(
    `Done${dryRun ? ' (dry run)' : ''}. ` +
    `Column inserted: ${summary.inserted}, blanks backfilled: ${summary.backfilled}, ` +
    `already complete: ${summary.unchanged}, skipped: ${summary.skipped}.`
  );
}

/**
 * Migrates one tab. Returns { status: 'inserted' | 'backfilled' | 'unchanged' | 'skipped', note }.
 */
function migrateTab_(tab, dryRun, books) {
  const width  = Math.max(tab.getLastColumn(), PRE_BOOK_HEADERS.length + 1);
  const header = tab.getRange(1, 1, 1, width).getValues()[0].map(v => String(v || '').trim());

  if (header[BOOK_COL - 1] === 'Book') {
    if (header[BOOK_COL] !== 'Lesson #') {
      return { status: 'skipped', note: `has "Book" in column E but column F is "${header[BOOK_COL]}", not "Lesson #"` };
    }
    const n = backfillBook_(tab, dryRun, books);
    return n === 0
      ? { status: 'unchanged', note: 'Book column present, no blank cells' }
      : { status: 'backfilled', note: `Book column present, ${n} blank cell(s) ${dryRun ? 'would be filled' : 'filled'}` };
  }

  const matchesOld = PRE_BOOK_HEADERS.every((h, i) => header[i].toLowerCase() === h.toLowerCase());
  if (!matchesOld) {
    return { status: 'skipped', note: `unexpected header: ${JSON.stringify(header.slice(0, PRE_BOOK_HEADERS.length))}` };
  }

  if (dryRun) {
    const n = backfillBook_(tab, true, books, true);
    return { status: 'inserted', note: `would insert Book column and fill ${n} row(s)` };
  }

  tab.insertColumnBefore(BOOK_COL);
  tab.getRange(1, BOOK_COL).setValue('Book');
  tab.setColumnWidth(BOOK_COL, 55);
  const n = backfillBook_(tab, false, books);
  SpreadsheetApp.flush();
  return { status: 'inserted', note: `inserted Book column, filled ${n} row(s)` };
}

/**
 * Fills blank Book cells (rows 2..last) with the book resolveBook() gives each row's
 * Timestamp. Existing values are left alone. With `preInsert` true (dry run on an
 * old-layout tab) the Book column does not exist yet, so every row counts as blank.
 * Returns the number of cells filled (or that would be filled).
 */
function backfillBook_(tab, dryRun, books, preInsert) {
  const lastRow = tab.getLastRow();
  if (lastRow < 2) return 0;

  const rows   = lastRow - 1;
  const ident  = tab.getRange(2, 1, rows, 4).getValues();                        // cols A–D (Timestamp is B)
  const cells  = preInsert ? null : tab.getRange(2, BOOK_COL, rows, 1).getValues();

  let filled = 0;
  const out = [];
  for (let i = 0; i < rows; i++) {
    const existing = preInsert ? '' : String(cells[i][0] || '').trim();
    if (existing !== '') { out.push([cells[i][0]]); continue; }

    const hasContent = ident[i].some(v => String(v || '').trim() !== '');
    if (!hasContent) { out.push(['']); continue; }

    const ts = ident[i][1];
    out.push([ts instanceof Date ? resolveBook(ts, books) : books[0].book]);
    filled++;
  }

  if (!dryRun && filled > 0) tab.getRange(2, BOOK_COL, rows, 1).setValues(out);
  return filled;
}
