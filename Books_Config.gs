/**
 * Books Config — LQM Al-Falah 26
 *
 * Shared by the intake, the aggregator and the migration. Add this same file to
 * BOTH Apps Script projects (the intake project and the aggregator project) —
 * it declares no spreadsheet IDs of its own; callers pass the Master spreadsheet.
 *
 * ── Master Sheet — Book_HW tab ───────────────────────────────────────────────
 *   Book #  |  Start Date  |  Aggregate Sheet Link
 *
 *   One row per book. An email belongs to the latest book whose Start Date has
 *   passed (by the email's own date). Before the first start date it is the first book.
 *
 *   Book #                 1, 2, 3 … consecutive, no gaps.
 *   Start Date             When the book begins accepting submissions, Eastern time.
 *                          Required for every book except Book 1 (blank = from the
 *                          beginning). A date-time cell, or text such as
 *                          "2026-11-14 00:00". Dates must increase with the book number.
 *   Aggregate Sheet Link   The book's "All HW Submissions" spreadsheet. Leave blank to
 *                          stop aggregating a finished book; the book's rows are then
 *                          skipped by the aggregator.
 *
 * Starting a new book = add one row to Book_HW. No script changes.
 *
 * If the tab is missing or invalid, loadBooks() throws and the calling script
 * stops rather than guess a book.
 */

const BOOKS_TAB       = 'Book_HW';
// Toronto and New York share the same Eastern rules. The Sheets UI's "Eastern Time - Toronto"
// entry is stored as America/Nassau, so we require America/New_York ("Eastern Time" in the UI).
const MASTER_TIMEZONE = 'America/New_York';

/**
 * Reads and validates the Book_HW tab.
 * Returns [{ book: number, start: Date|null, aggSheetId: string|null }] ordered by book.
 */
function loadBooks(masterSS) {
  const tz = masterSS.getSpreadsheetTimeZone();
  if (tz !== MASTER_TIMEZONE) {
    throw new Error(`Master sheet timezone is "${tz}", expected "${MASTER_TIMEZONE}". Fix it under File → Settings.`);
  }

  const tab = masterSS.getSheetByName(BOOKS_TAB);
  if (!tab) throw new Error(`Tab "${BOOKS_TAB}" not found in the Master sheet.`);

  const rows = tab.getDataRange().getValues();
  const header = (rows[0] || []).map(h => String(h || '').trim().toLowerCase());
  const iBook  = header.indexOf('book #');
  const iStart = header.indexOf('start date');
  const iLink  = header.indexOf('aggregate sheet link');
  if (iBook < 0 || iStart < 0 || iLink < 0) {
    throw new Error(`${BOOKS_TAB} must have headers "Book #", "Start Date", "Aggregate Sheet Link". Found: ${JSON.stringify(rows[0])}`);
  }

  const books = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const rawBook = r[iBook];
    if (String(rawBook).trim() === '' && String(r[iStart]).trim() === '' && String(r[iLink]).trim() === '') continue;

    const book = Number(rawBook);
    if (!Number.isInteger(book) || book < 1) throw new Error(`${BOOKS_TAB} row ${i + 1}: Book # "${rawBook}" is not a whole number ≥ 1.`);

    const start = parseBookStart_(r[iStart], i + 1);

    const link = String(r[iLink] || '').trim();
    let aggSheetId = null;
    if (link) {
      const m = link.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if (!m) throw new Error(`${BOOKS_TAB} row ${i + 1}: cannot read a spreadsheet ID from "${link}".`);
      aggSheetId = m[1];
    }
    books.push({ book, start, aggSheetId });
  }

  if (books.length === 0) throw new Error(`${BOOKS_TAB} has no books.`);
  books.sort((a, b) => a.book - b.book);

  for (let i = 0; i < books.length; i++) {
    if (books[i].book !== i + 1) throw new Error(`${BOOKS_TAB}: Book # values must be 1, 2, 3 … with no gaps or repeats.`);
    if (i === 0) continue;
    if (!books[i].start) throw new Error(`${BOOKS_TAB}: Book ${books[i].book} needs a Start Date.`);
    const prev = books[i - 1].start; // Book 1 may have none
    if (prev && books[i].start <= prev) {
      throw new Error(`${BOOKS_TAB}: Book ${books[i].book} must start after Book ${books[i - 1].book}.`);
    }
  }
  return books;
}

/**
 * Decides which book an email belongs to from its date alone: the latest book
 * whose start has passed; the first book if none has.
 */
function resolveBook(emailDate, books) {
  let result = books[0].book;
  for (const b of books) {
    if (b.start && emailDate >= b.start) result = b.book;
  }
  return result;
}

/** Accepts a Date cell, or text "yyyy-MM-dd[ HH:mm[:ss]]" read as Eastern time. Blank → null. */
function parseBookStart_(value, rowNum) {
  if (value instanceof Date) return value;
  const s = String(value === null || value === undefined ? '' : value).trim();
  if (s === '') return null;
  // Pick the format from the text's shape; parseDate alone would quietly ignore trailing text.
  const m = /^\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}(:\d{2})?)?$/.exec(s);
  if (m) {
    const fmt = s.length === 10 ? 'yyyy-MM-dd' : (m[1] ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd HH:mm');
    try {
      return Utilities.parseDate(s, MASTER_TIMEZONE, fmt);
    } catch (e) { /* fall through to the error below */ }
  }
  throw new Error(`${BOOKS_TAB} row ${rowNum}: cannot read Start Date "${s}". Use a date cell or text like 2026-11-14 00:00.`);
}
