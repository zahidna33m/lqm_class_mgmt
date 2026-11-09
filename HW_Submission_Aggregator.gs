/**
 * HW Submission Aggregator — LQM Al-Falah 26
 *
 * Reads every group's HW_Submissions sheet (via the Master Google Sheet's
 * Groups tab) and copies rows into one central "All HW Submissions" spreadsheet
 * per book (listed in the Master sheet's Book_HW tab). Each source row is routed
 * by its Book column.
 * On each run:
 *   - New rows are appended.
 *   - Rows whose source data changed are updated in place (all 11 source columns
 *     plus a recalculated Internal Status).
 *   - Rows whose Book + Student ID + Lesson # already exists elsewhere get "Duplicate".
 *   - Unchanged rows are skipped.
 *   - Source rows missing a Submission ID (manually added rows) get a UUID
 *     written back to the source sheet before being processed.
 *   - A blank Book is treated as Book 1.
 *   - A group sheet that has not been migrated to the Book layout is skipped.
 *
 * Runs on an hourly time-based trigger. Run installAggregatorTrigger() once
 * from the Apps Script editor to register it.
 *
 * ── Setup (one-time) ──────────────────────────────────────────────────────────
 *   1. Create a standalone Apps Script project at script.google.com (not attached to any
 *      spreadsheet — it serves every book's aggregate, which it opens by ID from Book_HW).
 *   2. Paste this file, plus Books_Config.gs and Grader_Status_Report.gs (they share this
 *      project's constants), and Book_Migration.gs for the one-time migration.
 *   3. Run installAggregatorTrigger() once from the editor.
 *   4. Authorise Sheets access when prompted.
 *
 * ── Source tab columns (SRC_TAB_NAME, in each group sheet) ───────────────────
 *   Col A  Submission ID  (UUID — hidden)
 *   Col B  Timestamp
 *   Col C  Student ID
 *   Col D  Student Name
 *   Col E  Book
 *   Col F  Lesson #
 *   Col G  HW File Link
 *   Col H  Grader
 *   Col I  Grading Status
 *   Col J  Grade
 *   Col K  Comments
 *
 * ── Output tab columns (AGG_TAB_NAME) ────────────────────────────────────────
 *   Cols A–K are the source columns above, followed by:
 *   Col L  Internal Status   ("Duplicate" when Book + Student ID + Lesson # already exists; blank otherwise)
 *   Col M  Group ID
 *   Col N  Source Link       (URL of the Gxx HW_Submissions sheet this row came from)
 *
 * Column positions are named in AGG_COL (0-based); other scripts in this
 * project, such as Grader_Status_Report.gs, use the same constant.
 */

// ─── Configuration ─────────────────────────────────────────────────────────────

const MASTER_SHEET_ID = '14T7CyfSsci-Va0FOIc2NJiQiQ8T6VuiNYC_A7TSIeuE';
const AGG_TAB_NAME    = 'All_HW_Submissions';
const SRC_TAB_NAME    = 'HW_Submissions';

// The aggregate spreadsheet for each book comes from the Master sheet's Book_HW tab
// (see Books_Config.gs). A book with a blank link there is not aggregated.

const SRC_COLS = 11; // Submission ID through Comments

const AGG_HEADERS = [
  'Submission ID', 'Timestamp', 'Student ID', 'Student Name', 'Book', 'Lesson #', 'HW File Link',
  'Grader', 'Grading Status', 'Grade', 'Comments',
  'Internal Status', 'Group ID', 'Source Link',
];

// 0-based column positions, shared by the source and output tabs for cols 0–10.
const AGG_COL = {
  UUID: 0, TIMESTAMP: 1, STUDENT_ID: 2, STUDENT_NAME: 3, BOOK: 4, LESSON: 5,
  LINK: 6, GRADER: 7, STATUS: 8, GRADE: 9, COMMENTS: 10,
};

// ─── Entry Point ───────────────────────────────────────────────────────────────

function aggregateHomeworkSubmissions() {
  const masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);

  // ── Load existing output rows, one state object per book ──────────────────
  //
  // importKeyToEntry  Map<uuid, { rowNum: number, data: any[] }>
  //   rowNum is the 1-based sheet row (2 = first data row after the header).
  //   data   is the full 14-column array as read from the sheet.
  //
  // studentLessonSet  Set<"studentId|lessonNo">
  //   Tracks all Student ID + Lesson # pairs currently in that book's output
  //   sheet so we can detect duplicates when appending or recalculating after
  //   an update. Each book has its own set, so Book 1 and Book 2 never collide.

  const states = new Map(); // book number → state
  for (const target of loadBooks(masterSS)) {
    if (!target.aggSheetId) {
      Logger.log(`Book ${target.book}: no Aggregate Sheet Link in ${BOOKS_TAB} — not aggregated.`);
      continue;
    }
    const tab = getOrCreateAggTab(SpreadsheetApp.openById(target.aggSheetId));
    const state = {
      book: target.book,
      tab,
      importKeyToEntry: new Map(),
      studentLessonSet: new Set(),
      newRows: [], // rows to append at the end
      updates: [], // { rowNum, rowData } — rows to overwrite in place
    };

    const lastRow = tab.getLastRow();
    if (lastRow > 1) {
      const existing = tab.getRange(2, 1, lastRow - 1, AGG_HEADERS.length).getValues();
      for (let i = 0; i < existing.length; i++) {
        const row       = existing[i];
        const rowNum    = i + 2;
        const uuid      = String(row[AGG_COL.UUID] || '').trim();
        const studentId = String(row[AGG_COL.STUDENT_ID] || '').trim();
        const lessonNo  = String(row[AGG_COL.LESSON] || '').trim();
        if (uuid) state.importKeyToEntry.set(uuid, { rowNum, data: row });
        if (studentId && lessonNo) state.studentLessonSet.add(`${studentId}|${lessonNo}`);
      }
    }
    states.set(target.book, state);
  }

  // ── Walk every group in the Master Sheet's Groups tab ─────────────────────

  const groupRows = masterSS.getSheetByName('Groups').getDataRange().getValues();

  for (let i = 1; i < groupRows.length; i++) {
    const groupId   = String(groupRows[i][0] || '').trim(); // Col A
    const sheetLink = String(groupRows[i][6] || '').trim(); // Col G — HW Submission Sheet Link
    if (!groupId || !sheetLink) continue;

    const sheetId = extractSheetId(sheetLink);
    if (!sheetId) {
      Logger.log(`Group ${groupId}: cannot parse sheet ID from "${sheetLink}". Skipping.`);
      continue;
    }

    let srcSS;
    try {
      srcSS = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log(`Group ${groupId}: cannot open sheet ${sheetId} — ${e}`);
      continue;
    }

    const srcTab = srcSS.getSheetByName(SRC_TAB_NAME);
    if (!srcTab) {
      Logger.log(`Group ${groupId}: tab "${SRC_TAB_NAME}" not found. Skipping.`);
      continue;
    }

    // Guard against reading a sheet that still has the pre-Book layout.
    const bookHeader = String(srcTab.getRange(1, AGG_COL.BOOK + 1).getValue() || '').trim();
    if (bookHeader !== 'Book') {
      Logger.log(`Group ${groupId}: column E is "${bookHeader}", not "Book" — sheet not migrated. Skipping.`);
      continue;
    }

    const srcLastRow = srcTab.getLastRow();
    if (srcLastRow < 2) continue; // empty or header-only

    // Read all 11 source columns (Submission ID through Comments).
    const srcValues = srcTab.getRange(2, 1, srcLastRow - 1, SRC_COLS).getValues();

    // Fill missing UUIDs for any manually added rows, writing them back to the
    // source sheet so they are stable on future runs.
    const missingUuidRows = [];
    for (let j = 0; j < srcValues.length; j++) {
      if (!String(srcValues[j][AGG_COL.UUID] || '').trim()) {
        const newUuid   = Utilities.getUuid();
        srcValues[j][AGG_COL.UUID] = newUuid;
        missingUuidRows.push([j + 2, newUuid]); // [1-based sheet row, uuid]
      }
    }
    for (const [row, uuid] of missingUuidRows) {
      srcTab.getRange(row, 1).setValue(uuid);
    }

    for (const r of srcValues) {
      const uuid      = String(r[AGG_COL.UUID] || '').trim();
      const studentId = String(r[AGG_COL.STUDENT_ID] || '').trim();
      const lessonNo  = String(r[AGG_COL.LESSON] || '').trim();
      const srcData   = r.slice(0, SRC_COLS); // all 11 source columns

      if (!uuid) continue;

      // A blank Book counts as Book 1; a book with no aggregate is skipped.
      const book  = Number(r[AGG_COL.BOOK]) || 1;
      const state = states.get(book);
      if (!state) continue;
      const { importKeyToEntry, studentLessonSet } = state;

      if (importKeyToEntry.has(uuid)) {
        // ── Row already in the output sheet — check for changes ─────────────
        const entry = importKeyToEntry.get(uuid);
        // Compare columns 2–10 only (skip Submission ID at index 0 and Timestamp
        // at index 1 — Timestamp never changes and floating-point round-trip
        // differences cause false updates).
        if (rowsEqual(entry.data.slice(2, SRC_COLS), srcData.slice(2, SRC_COLS))) {
          continue; // nothing changed
        }

        // Source data changed. Remove the old student|lesson pair so the
        // recalculated status doesn't treat the row as a duplicate of itself.
        const oldStudentId = String(entry.data[AGG_COL.STUDENT_ID] || '').trim();
        const oldLessonNo  = String(entry.data[AGG_COL.LESSON] || '').trim();
        if (oldStudentId && oldLessonNo) {
          studentLessonSet.delete(`${oldStudentId}|${oldLessonNo}`);
        }

        const internalStatus = computeInternalStatus(studentId, lessonNo, studentLessonSet);
        if (studentId && lessonNo) studentLessonSet.add(`${studentId}|${lessonNo}`);

        const rowData = [...srcData, internalStatus, groupId, sheetLink];
        state.updates.push({ rowNum: entry.rowNum, rowData });

        // Keep the cached entry current so later rows in this run see the
        // updated student|lesson pair, not the stale one.
        importKeyToEntry.set(uuid, { rowNum: entry.rowNum, data: rowData });

      } else {
        // ── New row ──────────────────────────────────────────────────────────
        const internalStatus = computeInternalStatus(studentId, lessonNo, studentLessonSet);
        if (studentId && lessonNo) studentLessonSet.add(`${studentId}|${lessonNo}`);

        const rowData = [...srcData, internalStatus, groupId, sheetLink];
        importKeyToEntry.set(uuid, { rowNum: -1, data: rowData });
        state.newRows.push(rowData);
      }
    }
  }

  // ── Write results ─────────────────────────────────────────────────────────

  for (const state of states.values()) {
    for (const { rowNum, rowData } of state.updates) {
      state.tab.getRange(rowNum, 1, 1, AGG_HEADERS.length).setValues([rowData]);
    }

    if (state.newRows.length > 0) {
      state.tab
        .getRange(state.tab.getLastRow() + 1, 1, state.newRows.length, AGG_HEADERS.length)
        .setValues(state.newRows);
    }

    Logger.log(
      `Book ${state.book}: ${state.updates.length} updated, ${state.newRows.length} new row(s) added.`
    );
  }
}

// ─── Internal Status ───────────────────────────────────────────────────────────

/**
 * Derives the Internal Status for a row being added or recalculated.
 * Returns "Duplicate" when the Student ID + Lesson # pair already exists in the
 * same book's output sheet (the caller passes that book's set); blank otherwise.
 */
function computeInternalStatus(studentId, lessonNo, studentLessonSet) {
  if (studentId && lessonNo && studentLessonSet.has(`${studentId}|${lessonNo}`)) return 'Duplicate';
  return '';
}

// ─── Comparison Helpers ────────────────────────────────────────────────────────

/** Returns true when two same-length row arrays contain equivalent values. */
function rowsEqual(a, b) {
  for (let i = 0; i < a.length; i++) {
    if (normalizeCell(a[i]) !== normalizeCell(b[i])) return false;
  }
  return true;
}

/**
 * Normalises a cell value for comparison.
 * Dates are compared by epoch ms; everything else is trimmed to a string.
 */
function normalizeCell(v) {
  if (v instanceof Date) return v.getTime();
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Returns the aggregated tab, creating it with a frozen header row if absent.
 */
function getOrCreateAggTab(ss) {
  let tab = ss.getSheetByName(AGG_TAB_NAME);
  if (!tab) {
    tab = ss.insertSheet(AGG_TAB_NAME);
    tab.appendRow(AGG_HEADERS);
    tab.setFrozenRows(1);
    return tab;
  }
  if (tab.getLastRow() === 0) {
    tab.appendRow(AGG_HEADERS);
    tab.setFrozenRows(1);
  }
  return tab;
}

/** Extracts a sheet ID from a Google Sheets URL. */
function extractSheetId(url) {
  if (!url) return null;
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// ─── Trigger Setup ─────────────────────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor.
 * Removes any existing trigger for this function and registers a fresh hourly one.
 */
function installAggregatorTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'aggregateHomeworkSubmissions')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('aggregateHomeworkSubmissions')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly trigger installed for aggregateHomeworkSubmissions.');
}
