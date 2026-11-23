/**
 * Audit: homework PDFs with no matching submission row — LQM Al-Falah 26   (v3)
 *
 * One-off, READ-ONLY diagnostic. Changes nothing in Drive, Sheets or Gmail.
 *
 * For every group in the Master sheet's Groups tab it lists the PDFs in the group's HW Upload
 * Folder and checks whether the group's HW_Submissions sheet has a row for the same
 * Student ID + Lesson #. Both come from the file name (e.g. L04_7061_...). The link in the
 * HW File Link column is deliberately NOT used.
 *
 *   MATCHED   a row exists for that Student ID + Lesson #.
 *   MISSING   no such row. Marked "(in aggregate)" if an aggregate sheet has a row for that
 *             student + lesson, i.e. the row existed once and was removed from the group sheet.
 *   UNPARSED  the file name has no student ID, so it can't be matched. Judge it by name.
 *
 * Limits: one row is enough to match, so a second PDF for the same student + lesson is not
 * reported; and if a lesson number was corrected by hand in the sheet, the PDF shows as MISSING.
 *
 * HOW TO USE
 *   Paste into the HW Intake or HW Aggregator project (needs MASTER_SHEET_ID and loadBooks(), which
 *   both have), run auditOrphanPdfs() and read the execution log. Replace any earlier copy, and
 *   delete the file afterwards. All names start with "audit" so they cannot clash with other files.
 */

const AUDIT_GROUP_TAB = 'HW_Submissions';     // tab name in each group sheet
const AUDIT_AGG_TAB   = 'All_HW_Submissions'; // tab name in each aggregate sheet

function auditOrphanPdfs() {
  const master = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const tz = master.getSpreadsheetTimeZone();

  // Student ID + lesson keys present in any book's aggregate sheet.
  const inAggregate = new Set();
  for (const book of loadBooks(master)) {
    if (!book.aggSheetId) continue;
    auditReadKeys_(SpreadsheetApp.openById(book.aggSheetId).getSheetByName(AUDIT_AGG_TAB))
      .forEach(k => inAggregate.add(k));
  }

  const groupRows = master.getSheetByName('Groups').getDataRange().getValues();
  const missing = [], unparsed = [];
  const totals = { pdfs: 0, matched: 0, missing: 0, unparsed: 0 };

  for (let i = 1; i < groupRows.length; i++) {
    const groupId    = String(groupRows[i][0] || '').trim();
    const folderLink = String(groupRows[i][5] || '').trim(); // F — HW Upload Folder Link
    const sheetLink  = String(groupRows[i][6] || '').trim(); // G — HW Submission Sheet Link
    if (!groupId || !folderLink || !sheetLink) continue;

    const folderId = auditMatch_(folderLink, /\/folders\/([A-Za-z0-9_-]+)/);
    const sheetId  = auditMatch_(sheetLink,  /\/d\/([A-Za-z0-9_-]+)/);
    if (!folderId || !sheetId) { Logger.log(`Group ${groupId}: cannot parse folder/sheet link. Skipped.`); continue; }

    let keys;
    try {
      keys = auditReadKeys_(SpreadsheetApp.openById(sheetId).getSheetByName(AUDIT_GROUP_TAB));
    } catch (e) { Logger.log(`Group ${groupId}: cannot read sheet — ${e}`); continue; }

    const g = { pdfs: 0, matched: 0, missing: 0, unparsed: 0 };
    try {
      const files = DriveApp.getFolderById(folderId).getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (f.getMimeType() !== MimeType.PDF) continue;
        g.pdfs++;
        const name = f.getName();
        const created = f.getDateCreated().getTime();

        const parsed = auditParseName_(name);
        if (!parsed) { g.unparsed++; unparsed.push([groupId, created, name, f.getUrl()]); continue; }

        const key = parsed.studentId + '|' + parsed.lesson;
        if (keys.has(key)) { g.matched++; continue; }

        g.missing++;
        missing.push([groupId, created, inAggregate.has(key) ? 'MISSING (in aggregate)' : 'MISSING',
          name, f.getUrl()]);
      }
    } catch (e) { Logger.log(`Group ${groupId}: cannot read folder — ${e}`); continue; }

    for (const k of Object.keys(totals)) totals[k] += g[k];
    Logger.log(`Group ${groupId}: ${g.pdfs} PDF(s) — matched ${g.matched}, MISSING ${g.missing}, unparsed ${g.unparsed}.`);
  }

  Logger.log(`=== TOTAL ${totals.pdfs} PDF(s): matched ${totals.matched}, MISSING ${totals.missing}, unparsed ${totals.unparsed} ===`);

  const byGroupThenDate = (a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]);
  const fmt = ms => Utilities.formatDate(new Date(ms), tz, 'yyyy-MM-dd HH:mm');

  Logger.log(`--- MISSING: ${missing.length} PDF(s) with no row for that student + lesson ---`);
  missing.sort(byGroupThenDate).forEach(([gid, ms, status, name, url]) =>
    Logger.log(`${gid} | ${fmt(ms)} | ${status} | ${name} | ${url}`));

  if (unparsed.length) {
    Logger.log(`--- UNPARSED: ${unparsed.length} PDF(s) whose name has no student ID ---`);
    unparsed.sort(byGroupThenDate).forEach(([gid, ms, name, url]) =>
      Logger.log(`${gid} | ${fmt(ms)} | ${name} | ${url}`));
  }
}

/** Set of "studentId|lesson" keys, one per data row of a group or aggregate sheet. */
function auditReadKeys_(tab) {
  const keys = new Set();
  if (!tab || tab.getLastRow() < 2) return keys;

  const lastCol = tab.getLastColumn();
  const header = tab.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const cSid = header.indexOf('Student ID'), cLes = header.indexOf('Lesson #');
  if (cSid < 0 || cLes < 0) return keys;

  const values = tab.getRange(2, 1, tab.getLastRow() - 1, lastCol).getValues();
  for (const row of values) {
    const sid = String(row[cSid] || '').trim();
    if (!sid) continue;
    const lesson = parseInt(String(row[cLes] || '').trim(), 10);
    keys.add(sid + '|' + (isNaN(lesson) ? '' : lesson));
  }
  return keys;
}

/** "L04_7061_x.pdf" or "4_7141_x.pdf" -> { lesson: 4, studentId: '7061' }; "7234_x.pdf" -> lesson ''. */
function auditParseName_(name) {
  let m = String(name).match(/^L?(\d{1,2})_(\d{4})_/);
  if (m) return { lesson: parseInt(m[1], 10), studentId: m[2] };
  m = String(name).match(/^(\d{4})_/);
  if (m) return { lesson: '', studentId: m[1] };
  return null;
}

function auditMatch_(text, pattern) {
  const m = String(text || '').match(pattern);
  return m ? m[1] : null;
}
