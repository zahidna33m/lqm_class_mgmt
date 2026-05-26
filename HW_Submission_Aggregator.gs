/**
 * HW Submission Aggregator — LQM Al-Falah 26
 *
 * Reads every group's HW_Submissions sheet (via the Master Google Sheet's
 * Groups tab) and copies rows into the central "All HW Submissions" spreadsheet.
 * On each run:
 *   - New rows are appended.
 *   - Rows whose source data changed are updated in place (all 10 source columns
 *     plus a recalculated Internal Status).
 *   - Rows whose Student ID + Lesson # already exists elsewhere get "Duplicate".
 *   - Unchanged rows are skipped.
 *   - Source rows missing a Submission ID (manually added rows) get a UUID
 *     written back to the source sheet before being processed.
 *
 * Runs on an hourly time-based trigger. Run installAggregatorTrigger() once
 * from the Apps Script editor to register it.
 *
 * ── Setup (one-time) ──────────────────────────────────────────────────────────
 *   1. Open the "All HW Submissions" Google Sheet → Extensions → Apps Script.
 *   2. Paste this file.
 *   3. Run installAggregatorTrigger() once from the editor.
 *   4. Authorise Sheets access when prompted.
 *
 * ── Output tab columns (AGG_TAB_NAME) ────────────────────────────────────────
 *   Col A  Submission ID  (UUID — hidden; stable import key, immune to URL/format changes)
 *   Col B  Timestamp
 *   Col C  Student ID
 *   Col D  Student Name
 *   Col E  Lesson #
 *   Col F  HW File Link
 *   Col G  Grader
 *   Col H  Grading Status
 *   Col I  Grade
 *   Col J  Comments
 *   Col K  Internal Status   ("Duplicate" when Student ID + Lesson # already exists; blank otherwise)
 *   Col L  Group ID
 *   Col M  Source Link       (URL of the Gxx HW_Submissions sheet this row came from)
 */

// ─── Configuration ─────────────────────────────────────────────────────────────

const MASTER_SHEET_ID = '14T7CyfSsci-Va0FOIc2NJiQiQ8T6VuiNYC_A7TSIeuE';
const AGG_SHEET_ID    = '1Pr2OCOhjCoxTuC7eDZmftahHERRyLgC04-kkaJC47p0';
const AGG_TAB_NAME    = 'All_HW_Submissions';
const SRC_TAB_NAME    = 'HW_Submissions';

const AGG_HEADERS = [
  'Submission ID', 'Timestamp', 'Student ID', 'Student Name', 'Lesson #', 'HW File Link',
  'Grader', 'Grading Status', 'Grade', 'Comments',
  'Internal Status', 'Group ID', 'Source Link',
];

// ─── Entry Point ───────────────────────────────────────────────────────────────

function aggregateHomeworkSubmissions() {
  const masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const aggSS    = SpreadsheetApp.openById(AGG_SHEET_ID);
  const aggTab   = getOrCreateAggTab(aggSS);

  // ── Load existing output rows ─────────────────────────────────────────────
  //
  // importKeyToEntry  Map<uuid, { rowNum: number, data: any[] }>
  //   rowNum is the 1-based sheet row (2 = first data row after the header).
  //   data   is the full 13-column array as read from the sheet.
  //
  // studentLessonSet  Set<"studentId|lessonNo">
  //   Tracks all Student ID + Lesson # pairs currently in the output sheet so
  //   we can detect duplicates when appending or recalculating after an update.

  const importKeyToEntry = new Map();
  const studentLessonSet = new Set();

  const lastRow = aggTab.getLastRow();
  if (lastRow > 1) {
    const existing = aggTab.getRange(2, 1, lastRow - 1, AGG_HEADERS.length).getValues();
    for (let i = 0; i < existing.length; i++) {
      const row       = existing[i];
      const rowNum    = i + 2;
      const uuid      = String(row[0] || '').trim();
      const studentId = String(row[2] || '').trim();
      const lessonNo  = String(row[4] || '').trim();
      if (uuid) importKeyToEntry.set(uuid, { rowNum, data: row });
      if (studentId && lessonNo) studentLessonSet.add(`${studentId}|${lessonNo}`);
    }
  }

  // ── Walk every group in the Master Sheet's Groups tab ─────────────────────

  const groupRows = masterSS.getSheetByName('Groups').getDataRange().getValues();
  const newRows   = []; // rows to append at the end
  const updates   = []; // { rowNum, rowData } — rows to overwrite in place

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

    const srcLastRow = srcTab.getLastRow();
    if (srcLastRow < 2) continue; // empty or header-only

    // Read all 10 source columns (Submission ID through Comments).
    const srcValues = srcTab.getRange(2, 1, srcLastRow - 1, 10).getValues();

    // Fill missing UUIDs for any manually added rows, writing them back to the
    // source sheet so they are stable on future runs.
    const missingUuidRows = [];
    for (let j = 0; j < srcValues.length; j++) {
      if (!String(srcValues[j][0] || '').trim()) {
        const newUuid   = Utilities.getUuid();
        srcValues[j][0] = newUuid;
        missingUuidRows.push([j + 2, newUuid]); // [1-based sheet row, uuid]
      }
    }
    for (const [row, uuid] of missingUuidRows) {
      srcTab.getRange(row, 1).setValue(uuid);
    }

    for (const r of srcValues) {
      const uuid      = String(r[0] || '').trim();
      const studentId = String(r[2] || '').trim();
      const lessonNo  = String(r[4] || '').trim();
      const srcData   = r.slice(0, 10); // all 10 source columns

      if (!uuid) continue;

      if (importKeyToEntry.has(uuid)) {
        // ── Row already in the output sheet — check for changes ─────────────
        const entry = importKeyToEntry.get(uuid);
        // Compare columns 2–9 only (skip Submission ID at index 0 and Timestamp
        // at index 1 — Timestamp never changes and floating-point round-trip
        // differences cause false updates).
        if (rowsEqual(entry.data.slice(2, 10), srcData.slice(2, 10))) {
          continue; // nothing changed
        }

        // Source data changed. Remove the old student|lesson pair so the
        // recalculated status doesn't treat the row as a duplicate of itself.
        const oldStudentId = String(entry.data[2] || '').trim();
        const oldLessonNo  = String(entry.data[4] || '').trim();
        if (oldStudentId && oldLessonNo) {
          studentLessonSet.delete(`${oldStudentId}|${oldLessonNo}`);
        }

        const internalStatus = computeInternalStatus(studentId, lessonNo, studentLessonSet);
        if (studentId && lessonNo) studentLessonSet.add(`${studentId}|${lessonNo}`);

        const rowData = [...srcData, internalStatus, groupId, sheetLink];
        updates.push({ rowNum: entry.rowNum, rowData });

        // Keep the cached entry current so later rows in this run see the
        // updated student|lesson pair, not the stale one.
        importKeyToEntry.set(uuid, { rowNum: entry.rowNum, data: rowData });

      } else {
        // ── New row ──────────────────────────────────────────────────────────
        const internalStatus = computeInternalStatus(studentId, lessonNo, studentLessonSet);
        if (studentId && lessonNo) studentLessonSet.add(`${studentId}|${lessonNo}`);

        const rowData = [...srcData, internalStatus, groupId, sheetLink];
        importKeyToEntry.set(uuid, { rowNum: -1, data: rowData });
        newRows.push(rowData);
      }
    }
  }

  // ── Write results ─────────────────────────────────────────────────────────

  for (const { rowNum, rowData } of updates) {
    aggTab.getRange(rowNum, 1, 1, AGG_HEADERS.length).setValues([rowData]);
  }

  if (newRows.length > 0) {
    aggTab
      .getRange(aggTab.getLastRow() + 1, 1, newRows.length, AGG_HEADERS.length)
      .setValues(newRows);
  }

  Logger.log(
    `Aggregation complete. ${updates.length} updated, ${newRows.length} new row(s) added.`
  );
}

// ─── Internal Status ───────────────────────────────────────────────────────────

/**
 * Derives the Internal Status for a row being added or recalculated.
 * Returns "Duplicate" when the Student ID + Lesson # pair already exists; blank otherwise.
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
