/**
 * Grader Status Report — LQM Al-Falah 26
 *
 * Notifies graders who have pending homework rows ("Assigned" or "Grading in
 * progress") with a status summary and links to their assigned group sheet(s).
 *
 * Two delivery modes:
 *   createGraderStatusDrafts()  — saves to Gmail Drafts for manual review
 *                                 before sending. Available from the menu.
 *   sendGraderStatusEmails()    — sends directly. Run by the twice-weekly
 *                                 time-based trigger (Tuesday + Friday, 5 pm).
 *
 * Run installGraderStatusTriggers() once from the Apps Script editor to
 * register the automatic schedule. Run createGraderStatusDrafts() from the
 * "LQM Grading" menu to create drafts on demand.
 *
 * This script lives in the All HW Submissions spreadsheet's Apps Script
 * project and shares MASTER_SHEET_ID, AGG_SHEET_ID, AGG_TAB_NAME, and
 * AGG_HEADERS with HW_Submission_Aggregator.gs — those are not redeclared here.
 *
 * ── Data sources ─────────────────────────────────────────────────────────────
 *   Aggregated sheet (AGG_TAB_NAME)
 *     Col G  Grader         (name — matched case-insensitively to Graders tab)
 *     Col H  Grading Status ("Assigned" / "Grading in progress" / "Grading complete")
 *
 *   Master Sheet — Graders tab
 *     Col A  Grader ID
 *     Col B  Name
 *     Col C  Email
 *     Col F  M/F            ("M" → "Br.", "F" → "Sr.")
 *
 *   Master Sheet — Groups tab
 *     Col A  Group ID
 *     Col B  Group Name
 *     Col G  HW Submission Sheet Link
 *     Col H  Grader 1  (Grader ID)
 *     Col I  Grader 2  (Grader ID)
 */

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LQM Grading')
    .addItem('Create Grader Status Drafts', 'createGraderStatusDrafts')
    .addToUi();
}

// ─── Public Entry Points ────────────────────────────────────────────────────────

/** Saves one Gmail draft per grader with pending work. Run from the menu. */
function createGraderStatusDrafts() {
  processGraderStatus_(true);
}

/** Sends one email per grader with pending work. Run by the scheduled trigger. */
function sendGraderStatusEmails() {
  processGraderStatus_(false);
}

// ─── Core Logic ────────────────────────────────────────────────────────────────

function processGraderStatus_(asDraft) {
  const masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const aggTab   = SpreadsheetApp.openById(AGG_SHEET_ID).getSheetByName(AGG_TAB_NAME);

  if (!aggTab) {
    Logger.log(`Tab "${AGG_TAB_NAME}" not found. Aborting.`);
    return;
  }

  // ── Load grader records from Master Sheet ─────────────────────────────────
  // Columns: Grader ID(0) | Name(1) | Email(2) | Phone(3) | Comments(4) | M/F(5)

  const graderByName = {}; // lowercase name → { id, name, email, gender }
  const graderById   = {}; // grader ID      → same record

  const graderRows = masterSS.getSheetByName('Graders').getDataRange().getValues();
  for (let i = 1; i < graderRows.length; i++) {
    const row    = graderRows[i];
    const id     = String(row[0] || '').trim();
    const name   = String(row[1] || '').trim();
    const email  = String(row[2] || '').trim();
    const gender = String(row[5] || '').trim();
    if (!id || !name) continue;
    const record = { id, name, email, gender };
    graderByName[name.toLowerCase()] = record;
    graderById[id]                   = record;
  }

  // ── Load group assignments from Master Sheet ──────────────────────────────
  // Columns: Group ID(0) | Group Name(1) | ... | HW Sheet Link(6) | Grader1(7) | Grader2(8)
  // Build: grader ID → [{ groupName, hwSheetLink }]

  const groupsByGraderId = {};

  const groupRows = masterSS.getSheetByName('Groups').getDataRange().getValues();
  for (let i = 1; i < groupRows.length; i++) {
    const row         = groupRows[i];
    const groupName   = String(row[1] || '').trim();
    const hwSheetLink = String(row[6] || '').trim();
    const grader1Id   = String(row[7] || '').trim();
    const grader2Id   = String(row[8] || '').trim();
    if (!groupName || !hwSheetLink) continue;
    for (const gid of [grader1Id, grader2Id].filter(Boolean)) {
      if (!groupsByGraderId[gid]) groupsByGraderId[gid] = [];
      groupsByGraderId[gid].push({ groupName, hwSheetLink });
    }
  }

  // ── Count grading statuses per grader from the aggregated sheet ───────────
  // Col G (index 6) = Grader, Col H (index 7) = Grading Status

  const lastRow = aggTab.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in aggregated sheet. Nothing to report.');
    return;
  }

  const aggData      = aggTab.getRange(2, 1, lastRow - 1, AGG_HEADERS.length).getValues();
  const graderCounts = {}; // grader name (as in sheet) → { assigned, inProgress, complete }

  for (const row of aggData) {
    const graderName    = String(row[6] || '').trim();
    const gradingStatus = String(row[7] || '').trim();
    if (!graderName) continue;

    if (!graderCounts[graderName]) {
      graderCounts[graderName] = { assigned: 0, inProgress: 0, complete: 0 };
    }
    if      (gradingStatus === 'Assigned')            graderCounts[graderName].assigned++;
    else if (gradingStatus === 'Grading in progress') graderCounts[graderName].inProgress++;
    else if (gradingStatus === 'Grading complete')    graderCounts[graderName].complete++;
  }

  // ── Deliver to each grader with pending work ──────────────────────────────

  const subject = 'LQM Al-Falah 26 — Homework Grading Status Update';
  let delivered = 0;

  for (const [graderName, counts] of Object.entries(graderCounts)) {
    if (counts.assigned === 0 && counts.inProgress === 0) continue;

    const grader = graderByName[graderName.toLowerCase()];
    if (!grader) {
      Logger.log(`ERROR: Grader "${graderName}" not found in Graders tab. Skipping.`);
      continue;
    }
    if (!grader.email) {
      Logger.log(`ERROR: Grader "${grader.name}" has no email address. Skipping.`);
      continue;
    }

    const groups = groupsByGraderId[grader.id] || [];
    if (groups.length === 0) {
      Logger.log(`ERROR: Grader "${grader.name}" has no groups assigned in Groups tab. Skipping.`);
      continue;
    }

    const prefix   = grader.gender === 'F' ? 'Sr.' : 'Br.';
    const htmlBody = buildStatusEmailHtml(prefix, grader.name, counts, groups);

    if (asDraft) {
      GmailApp.createDraft(grader.email, subject, '', { htmlBody });
      Logger.log(`Draft created for ${grader.name} (${grader.email}).`);
    } else {
      GmailApp.sendEmail(grader.email, subject, '', { htmlBody });
      Logger.log(`Email sent to ${grader.name} (${grader.email}).`);
    }
    delivered++;
  }

  Logger.log(`Done. ${delivered} email(s) ${asDraft ? 'drafted' : 'sent'}.`);
}

// ─── Email Builder ─────────────────────────────────────────────────────────────

function buildStatusEmailHtml(prefix, name, counts, groups) {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:680px;">

<p>Assalaamu alaykum ${prefix} ${name},</p>

<p>Here is your twice-weekly update on your assigned Arabic homework grading tasks.</p>

<p><strong>Your Current Status:</strong></p>
<ul style="margin-top:0;line-height:1.9;">
  <li>Assigned (Not Started): <strong>${counts.assigned}</strong></li>
  <li>Grading in Progress: <strong>${counts.inProgress}</strong></li>
  <li>Grading Complete: <strong>${counts.complete}</strong></li>
</ul>

<p>You can access your homework grading tasks from the following Google Sheet(s):</p>
<ul style="margin-top:0;line-height:1.9;">
  ${groups.map(g => `<li><a href="${g.hwSheetLink}">${g.groupName}</a></li>`).join('\n  ')}
</ul>

<p>A gentle reminder: If you have any tasks currently in the <strong>Assigned</strong> or <strong>Grading in Progress</strong> status, please try to wrap them up as soon as your schedule allows.</p>

<p>Thank you so much for your incredible volunteer support. Every minute you spend grading directly helps the students connect more deeply with the Language of the Qur'an. They are incredibly grateful for your time and feedback, and so are we!<br><br>
<span style="font-size:18px;">بارك الله فيك</span></p>

</body>
</html>`;
}

// ─── Trigger Setup ─────────────────────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor.
 * Removes any existing triggers for sendGraderStatusEmails and registers
 * two fresh ones — Tuesday and Friday at 5 pm.
 */
function installGraderStatusTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendGraderStatusEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendGraderStatusEmails')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(17)
    .create();

  ScriptApp.newTrigger('sendGraderStatusEmails')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();

  Logger.log('Triggers installed: sendGraderStatusEmails runs Tuesday and Friday at 5 pm.');
}
