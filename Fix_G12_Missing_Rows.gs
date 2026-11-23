/**
 * One-off repair: restore G12 submission rows that were never written — LQM Al-Falah 26
 *
 * Background: the audit found PDFs in G12's HW upload folder (mostly Sep 26 - Oct 21, 2026) that have
 * no row in G12's HW_Submissions sheet. This script rebuilds those rows from the saved PDFs.
 *
 * Runs in the HW INTAKE project (it uses loadMasterData(), extractDriveFolderId(), extractSheetId(),
 * HW_SHEET_TAB and resolveBook() from the intake files). Delete this file when the repair is done.
 * All names start with "fixG12" so they cannot clash with other files.
 *
 * TWO STEPS
 *   1. fixG12Preview()  READ-ONLY. Logs the rows that WOULD be added and everything skipped, and why.
 *   2. fixG12Apply()    Adds exactly those rows. Re-computes the plan first, and refuses to write if
 *                       FIXG12_EXPECTED_ROWS is set and differs from the number of rows planned.
 *
 * WHICH PDFs ARE CONSIDERED
 *   - File name must start with "L" + 2-digit lesson + "_" + 4-digit student ID, e.g. L18_7188_xyz.pdf.
 *   - Lesson must be 4..23 inclusive. L01-L03 and files without the "L" prefix are skipped.
 *   - Skipped when the sheet already has a row for that Student ID + Book + Lesson #.
 *   - If several PDFs exist for the same student + lesson, only the EARLIEST is restored (the
 *     intake rule is "the first submission stands"); the others are listed as skipped.
 *   - The student must exist in Group_Members (name comes from there).
 *
 * THE ROW THAT IS ADDED (same 11 columns as the intake script writes)
 *   Submission ID = new UUID | Timestamp = the PDF's Drive creation time | Student ID | Student Name
 *   | Book = resolveBook(timestamp) | Lesson # | HW File Link = the PDF | Grader = alternates between
 *   the group's two graders, starting after whoever is on the last row | Grading Status = Assigned
 *   | Grade blank | Comments = a note that the row was restored.
 *   After applying, run aggregateHomeworkSubmissions() so the new rows reach the aggregate.
 */

const FIXG12_GROUP_ID      = 'G12';
const FIXG12_MIN_LESSON    = 4;
const FIXG12_MAX_LESSON    = 23;
// Set this to the "rows to add" number shown by fixG12Preview() before running fixG12Apply().
const FIXG12_EXPECTED_ROWS = null;
const FIXG12_COMMENT       = 'Row restored from the saved PDF (was not logged when the email was processed).';
const FIXG12_COLS          = 11; // Submission ID .. Comments

// ─── Step 1: read-only preview ─────────────────────────────────────────────────

function fixG12Preview() {
  const plan = fixG12BuildPlan_();
  fixG12LogPlan_(plan);
  Logger.log(`PREVIEW ONLY — nothing was written. To add these ${plan.rows.length} row(s): set ` +
             `FIXG12_EXPECTED_ROWS = ${plan.rows.length}, then run fixG12Apply().`);
}

// ─── Step 2: apply ─────────────────────────────────────────────────────────────

function fixG12Apply() {
  const plan = fixG12BuildPlan_();
  fixG12LogPlan_(plan);

  if (plan.rows.length === 0) { Logger.log('Nothing to add.'); return; }
  if (FIXG12_EXPECTED_ROWS === null || FIXG12_EXPECTED_ROWS === undefined) {
    Logger.log('NOT WRITTEN: set FIXG12_EXPECTED_ROWS to the number you reviewed in fixG12Preview() first.');
    return;
  }
  if (plan.rows.length !== FIXG12_EXPECTED_ROWS) {
    Logger.log(`NOT WRITTEN: ${plan.rows.length} row(s) planned but FIXG12_EXPECTED_ROWS = ${FIXG12_EXPECTED_ROWS}. ` +
               'The data changed since the preview — run fixG12Preview() again and re-check.');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('NOT WRITTEN: could not get the script lock (another run is active). Try again.'); return; }
  try {
    const sheet = plan.sheet;
    const firstRow = sheet.getLastRow() + 1;
    const values = plan.rows.map(r => r.values);
    sheet.getRange(firstRow, 1, values.length, FIXG12_COLS).setValues(values);
    SpreadsheetApp.flush();

    // Read back: every new Submission ID must be on the sheet.
    const written = sheet.getRange(firstRow, 1, values.length, 1).getValues().map(r => String(r[0]));
    const missing = values.filter(v => written.indexOf(String(v[0])) < 0);
    if (missing.length) {
      Logger.log(`ERROR: ${missing.length} of ${values.length} row(s) did not read back from the sheet. Check ${FIXG12_GROUP_ID}.`);
    } else {
      Logger.log(`DONE: added ${values.length} row(s) to ${FIXG12_GROUP_ID}, sheet rows ${firstRow}-${firstRow + values.length - 1}, all read back. ` +
                 'Now run aggregateHomeworkSubmissions() in the aggregator project.');
    }
  } finally {
    lock.releaseLock();
  }
}

// ─── Plan ──────────────────────────────────────────────────────────────────────

/** Works out the rows to add and everything skipped. Writes nothing. */
function fixG12BuildPlan_() {
  const master = loadMasterData();
  const group = master.groups[FIXG12_GROUP_ID];
  if (!group) throw new Error(`Group ${FIXG12_GROUP_ID} not found in the Groups tab.`);
  const folderId = extractDriveFolderId(group.hwUploadFolderLink);
  const sheetId  = extractSheetId(group.hwSheetLink);
  if (!folderId || !sheetId) throw new Error(`Cannot parse ${FIXG12_GROUP_ID}'s folder/sheet links.`);

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(HW_SHEET_TAB);
  if (!sheet) throw new Error(`Tab "${HW_SHEET_TAB}" not found in the ${FIXG12_GROUP_ID} sheet.`);
  const bookHeader = String(sheet.getRange(1, 5).getValue() || '').trim();
  if (bookHeader !== 'Book') throw new Error(`Column E of the ${FIXG12_GROUP_ID} sheet is "${bookHeader}", not "Book". Aborting.`);

  // Keys already on the sheet: Student ID | Book | Lesson # (blank Book counts as Book 1).
  const existing = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    for (const r of sheet.getRange(2, 3, lastRow - 1, 4).getValues()) { // C..F
      const lesson = parseInt(String(r[3]).trim(), 10);
      existing.add(`${String(r[0]).trim()}|${String(r[2]).trim() || '1'}|${isNaN(lesson) ? '' : lesson}`);
    }
  }

  // Collect the candidate PDFs.
  const skipped = []; // { name, reason }
  const candidates = [];
  const files = DriveApp.getFolderById(folderId).getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== MimeType.PDF) continue;
    const name = f.getName();
    const m = name.match(/^L(\d{2})_(\d{4})_/);
    if (!m) { skipped.push({ name, reason: 'name does not start with L<lesson>_<student ID>_' }); continue; }
    const lesson = parseInt(m[1], 10), studentId = m[2];
    if (lesson < FIXG12_MIN_LESSON || lesson > FIXG12_MAX_LESSON) { skipped.push({ name, reason: `lesson ${lesson} outside ${FIXG12_MIN_LESSON}-${FIXG12_MAX_LESSON}` }); continue; }
    candidates.push({ file: f, name, lesson, studentId, created: f.getDateCreated() });
  }
  candidates.sort((a, b) => a.created - b.created);

  // Grader alternation: continue after whoever is on the last row of the sheet.
  const graderNames = [group.grader1, group.grader2].filter(Boolean)
    .map(id => (master.graders[id] && master.graders[id].name) || '');
  let lastGrader = lastRow > 1 ? String(sheet.getRange(lastRow, SHEET_COL.GRADER).getValue() || '').trim() : '';
  const nextGrader = () => {
    if (graderNames.length === 0) return '';
    if (graderNames.length === 1) return graderNames[0];
    const pick = lastGrader === graderNames[0] ? graderNames[1] : graderNames[0];
    lastGrader = pick;
    return pick;
  };

  const rows = []; // { desc, values }
  const planned = new Set();
  for (const c of candidates) {
    const book = resolveBook(c.created, master.books);
    const key = `${c.studentId}|${book}|${c.lesson}`;
    const label = `${c.name} (student ${c.studentId}, lesson ${c.lesson}, saved ${c.created.toISOString().slice(0, 16)}Z)`;

    if (existing.has(key)) { skipped.push({ name: label, reason: 'the sheet already has a row for this student + lesson' }); continue; }
    if (planned.has(key))  { skipped.push({ name: label, reason: 'an earlier PDF for the same student + lesson is being restored instead' }); continue; }
    const member = lookupByStudentId(c.studentId, master.members);
    if (!member)           { skipped.push({ name: label, reason: 'student ID not found in Group_Members' }); continue; }
    planned.add(key);

    const grader = nextGrader();
    rows.push({
      desc: `${label} -> ${member.studentName}, Book ${book}, grader "${grader}"` +
            (member.groupId !== FIXG12_GROUP_ID ? `  [WARNING: Group_Members says this student is in ${member.groupId}]` : ''),
      values: [Utilities.getUuid(), c.created, c.studentId, member.studentName, book, c.lesson,
               c.file.getUrl(), grader, 'Assigned', '', FIXG12_COMMENT],
    });
  }
  return { sheet, rows, skipped, existingCount: existing.size };
}

function fixG12LogPlan_(plan) {
  Logger.log(`=== ${FIXG12_GROUP_ID}: ${plan.rows.length} row(s) would be ADDED (sheet currently has ${plan.existingCount} student+lesson rows) ===`);
  plan.rows.forEach((r, i) => Logger.log(`ADD ${i + 1}: ${r.desc}`));
  Logger.log(`--- Skipped ${plan.skipped.length} PDF(s) ---`);
  plan.skipped
    .filter(s => !/^already|already has a row/.test(s.reason))   // keep the log readable: hide the ~60 normal matches
    .forEach(s => Logger.log(`SKIP: ${s.name} — ${s.reason}`));
  const hidden = plan.skipped.filter(s => /already has a row/.test(s.reason)).length;
  Logger.log(`(${hidden} PDF(s) already have a row and are not listed.)`);
}
