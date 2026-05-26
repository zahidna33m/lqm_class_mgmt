/**
 * HW Submission UUID Generator — LQM Al-Falah 26
 *
 * One-time migration script. Inserts a "Submission ID" column as Col A in
 * every group's HW_Submissions sheet, generates a UUID for every existing
 * data row, and hides the column from graders.
 *
 * Run generateSubmissionUUIDs() once from the Apps Script editor. The script
 * is idempotent — any sheet whose Col A header is already "Submission ID" is
 * skipped, so it is safe to re-run if execution fails partway through.
 *
 * ── Setup ────────────────────────────────────────────────────────────────────
 *   1. Add this file to the aggregator's Apps Script project (the one bound to
 *      the All HW Submissions sheet). It reuses MASTER_SHEET_ID, SRC_TAB_NAME,
 *      and extractSheetId() already defined in HW_Submission_Aggregator.gs.
 *   2. Run generateSubmissionUUIDs() from the editor.
 *   3. Verify a few source sheets: Col A should be hidden, labelled
 *      "Submission ID", with one UUID per data row.
 *   4. Delete this file from the project once migration is confirmed.
 *
 * ── What it does ─────────────────────────────────────────────────────────────
 *   For each group in the Master Sheet's Groups tab:
 *     1. Opens the group's HW_Submissions sheet.
 *     2. Skips if Col A header is already "Submission ID".
 *     3. Inserts a new blank column at position A, shifting all existing
 *        columns one position to the right.
 *     4. Writes "Submission ID" as the Col A header.
 *     5. Generates a UUID for every existing data row in a single batch write.
 *     6. Hides Col A.
 */

// ─── Entry Point ───────────────────────────────────────────────────────────────

function generateSubmissionUUIDs() {
  const masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const groupRows = masterSS.getSheetByName('Groups').getDataRange().getValues();

  let groupsProcessed = 0;
  let groupsSkipped   = 0;
  let totalUuids      = 0;

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
      Logger.log(`Group ${groupId}: cannot open sheet — ${e}. Skipping.`);
      continue;
    }

    const srcTab = srcSS.getSheetByName(SRC_TAB_NAME);
    if (!srcTab) {
      Logger.log(`Group ${groupId}: tab "${SRC_TAB_NAME}" not found. Skipping.`);
      continue;
    }

    // Idempotency guard — skip if already migrated
    if (String(srcTab.getRange(1, 1).getValue()).trim() === 'Submission ID') {
      Logger.log(`Group ${groupId}: already migrated. Skipping.`);
      groupsSkipped++;
      continue;
    }

    // Insert new Col A and write header
    srcTab.insertColumnBefore(1);
    srcTab.getRange(1, 1).setValue('Submission ID');

    // Generate and write UUIDs for all existing data rows in one batch write
    const lastRow = srcTab.getLastRow();
    if (lastRow > 1) {
      const uuids = Array.from({ length: lastRow - 1 }, () => [Utilities.getUuid()]);
      srcTab.getRange(2, 1, lastRow - 1, 1).setValues(uuids);
      totalUuids += lastRow - 1;
    }

    // Hide Col A from graders
    srcTab.hideColumns(1);

    Logger.log(`Group ${groupId}: ${lastRow > 1 ? lastRow - 1 : 0} UUID(s) written.`);
    groupsProcessed++;
  }

  Logger.log(
    `Migration complete — ${groupsProcessed} group(s) updated, ` +
    `${groupsSkipped} already migrated, ${totalUuids} total UUID(s) generated.`
  );
}

// extractSheetId() is defined in HW_Submission_Aggregator.gs and shared across
// all files in this project.
