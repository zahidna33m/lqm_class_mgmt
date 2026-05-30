/**
 * Homework Submission Processor — LQM Al-Falah 26
 *
 * Monitors a Gmail mailing list for student homework submissions. For each
 * email that carries a PDF attachment, the script saves the file to the
 * student's group Drive folder and logs the submission in that group's
 * Google Sheet. Emails are labelled in Gmail so they are never processed twice.
 *
 * Runs on an hourly time-based trigger. Run installTrigger() once from the
 * Apps Script editor to register it.
 *
 * ── Setup (one-time) ──────────────────────────────────────────────────────────
 *   1. Open Google Apps Script (script.google.com) and paste this file.
 *   2. Set MASTER_SHEET_ID (below) to your master Google Sheet's ID.
 *   3. Run installTrigger() once from the editor to register the hourly job.
 *   4. Authorise Gmail, Drive, and Sheets access when prompted.
 *
 * ── Google Sheet dependency (MASTER_SHEET_ID) ─────────────────────────────────
 *   Tab: Graders
 *     Col A  Grader ID
 *     Col B  Name
 *     Col C  Email
 *     Col D  Phone
 *     Col E  Comments
 *
 *   Tab: Groups
 *     Col A  Group ID
 *     Col B  Group Name
 *     Col C  Group Lead
 *     Col D  WhatsApp Link
 *     Col E  Group Share Folder Link   (used by Group_Comm.gs; not read here)
 *     Col F  HW Upload Folder Link     (Drive folder where PDFs are saved)
 *     Col G  HW Submission Sheet Link  (Google Sheet where submission rows are logged)
 *     Col H  Grader 1                  (Grader ID — looked up in Graders tab)
 *     Col I  Grader 2                  (Grader ID — looked up in Graders tab)
 *     Col J  Email Msg                 (checkbox; used by Group_Comm.gs; not read here)
 *     NOTE: G00 must always be configured with valid links as the catch-all fallback.
 *
 *   Tab: Group_Members
 *     Col A  Gender
 *     Col B  Student Name
 *     Col C  Age
 *     Col D  Postal Code
 *     Col E  FSA
 *     Col F  City
 *     Col G  Effective Email           (address the student sends homework from)
 *     Col H  Mobile Phone
 *     Col I  Student ID                (numeric, range 7001–7999)
 *     Col J  Status
 *     Col K  PC-4
 *     Col L  Attendance
 *     Col M  Group ID
 *
 *   Per-group HW Submission Sheet (URL from Groups tab, Col G)
 *   Tab: HW_Submissions
 *     Col A  Submission ID             (UUID — hidden; import key for the aggregator)
 *     Col B  Timestamp
 *     Col C  Student ID
 *     Col D  Student Name              (email username if student is unrecognised)
 *     Col E  Lesson #
 *     Col F  HW File Link
 *     Col G  Grader
 *     Col H  Grading Status            ("Assigned" for normal groups; "Incorrect submission" for G00)
 *     Col I  Grade                     (filled in by grader)
 *     Col J  Comments                  (filled in by grader)
 *
 * ── Gmail labels ──────────────────────────────────────────────────────────────
 *   HW_Submitted             Processed successfully into the correct group.
 *   HW Submission Error      Fell back to G00 (unknown student or bad group config).
 *   HW No Attachment         Email had no PDF — needs manual follow-up.
 *   HW Duplicate Submission  Student already has a row for this lesson — PDF not saved.
 *   HW_No_Group_Assigned     Student identified but not yet placed in a study group.
 *   Labels are created automatically on first run if they do not exist.
 *
 * ── Functions ─────────────────────────────────────────────────────────────────
 *   processHomeworkEmails()     Entry point. Searches Gmail for unlabelled
 *                               submissions and calls processThread() on each.
 *   processThread()             Handles one email thread: resolves the student
 *                               and group, saves the PDF, logs the sheet row,
 *                               and applies the appropriate Gmail label.
 *   loadMasterData()            Reads Graders, Groups, and Group_Members tabs
 *                               once per run and returns lookup maps.
 *   lookupByEmail()             Finds a student record by sender email address.
 *   lookupByStudentId()         Finds a student record by numeric Student ID.
 *   resolveGroupWithFallback()  Returns the correct group for a student, falling
 *                               back to G00 if the group is missing or
 *                               misconfigured.
 *   resolveGrader()             Picks a grader name for the group (random if two
 *                               graders are assigned).
 *   getPdfAttachments()         Filters a message's attachments to PDFs only.
 *   extractEmail()              Strips display name from a "Name <addr>" header.
 *   extractLessonId()           Parses a labelled lesson number (1–23) from body text.
 *   extractLessonIdBroad()      Parses a lesson number from the subject: labelled
 *                               pattern first, then any standalone number 1–23.
 *   extractStudentId()          Parses a labelled student ID (7001–7999) from body text.
 *   extractStudentIdBroad()     Parses a student ID from the subject: any 4-digit
 *                               number in range 7001–7999 not adjacent to other digits,
 *                               regardless of label.
 *   extractDriveFolderId()      Pulls the folder ID from a Drive folder URL.
 *   extractSheetId()            Pulls the sheet ID from a Google Sheets URL.
 *   buildFilename()             Builds the saved PDF filename:
 *                               L{nn}_{StudentID}_{originalName}.
 *   buildNoPdfResponseHtml()    HTML body for the "no PDF attached" draft reply.
 *   buildDuplicateResponseHtml() HTML body for the "duplicate submission" draft reply.
 *   buildG00StudentResponseHtml() HTML body for the "student not in a group" draft reply.
 *   ensureLabels()              Returns handles to all three Gmail labels,
 *                               creating any that do not yet exist.
 *   getOrCreateLabel()          Gets or creates a single Gmail label by name.
 *   installTrigger()            One-time setup: registers the hourly trigger.
 */

// ─── Configuration ─────────────────────────────────────────────────────────────

const MASTER_SHEET_ID  = '14T7CyfSsci-Va0FOIc2NJiQiQ8T6VuiNYC_A7TSIeuE';
const MAILING_LIST     = 'alfalah26_hw@lqmississauga.com';
const CATCH_ALL_GROUP  = 'G00';
const SEARCH_WINDOW    = '2h'; // wider than 1-hour trigger to absorb timing drift

const LABEL_PROCESSED     = 'HW_Submitted';
const LABEL_ERROR         = 'HW Submission Error';
const LABEL_NO_ATTACHMENT = 'HW No Attachment';
const LABEL_DUPLICATE     = 'HW Duplicate Submission';
const LABEL_NO_GROUP      = 'HW_No_Group_Assigned';
const HW_SHEET_TAB        = 'HW_Submissions';

// ─── Entry Point ───────────────────────────────────────────────────────────────

function processHomeworkEmails() {
  const labels = ensureLabels();
  const master = loadMasterData();

  // Only fetch threads that have NOT yet received any of our labels.
  const query = [
    `to:${MAILING_LIST}`,
    `-label:${LABEL_PROCESSED}`,
    `-label:"${LABEL_ERROR}"`,
    `-label:"${LABEL_NO_ATTACHMENT}"`,
    `-label:"${LABEL_DUPLICATE}"`,
    `-label:${LABEL_NO_GROUP}`,
    `newer_than:${SEARCH_WINDOW}`,
  ].join(' ');

  const threads = GmailApp.search(query);
  Logger.log(`Found ${threads.length} unprocessed thread(s).`);

  for (const thread of threads) {
    try {
      processThread(thread, labels, master);
    } catch (e) {
      Logger.log(`Unhandled error on thread ${thread.getId()}: ${e}\n${e.stack}`);
    }
  }
}

// ─── Thread Processing ─────────────────────────────────────────────────────────

function processThread(thread, labels, master) {
  // Homework submissions are always single-message threads.
  const message = thread.getMessages()[0];
  const from    = extractEmail(message.getFrom());
  const subject = message.getSubject() || '';
  const body    = message.getPlainBody() || '';
  const text    = `${subject} ${body}`;

  // 1. Check for PDF attachments first — no PDFs means nothing to save.
  const pdfs = getPdfAttachments(message);
  if (pdfs.length === 0) {
    thread.addLabel(labels.noAttachment);
    message.createDraftReply('', { htmlBody: buildNoPdfResponseHtml(), name: 'Zahid Naeem (LQM)' });
    Logger.log(`No PDF from ${from}. Labelled "${LABEL_NO_ATTACHMENT}", draft reply created.`);
    return;
  }

  // 2. Resolve student.
  //    Priority: broad student ID scan of subject → email lookup → labelled ID in body.
  //    If a subject ID is in range but not in the sheet, fall back to email and note it.
  let student = null;
  let subjectSidComment = '';
  const sidFromSubject = extractStudentIdBroad(subject);
  if (sidFromSubject) {
    student = lookupByStudentId(sidFromSubject, master.members);
    if (!student) {
      subjectSidComment = `Student ID ${sidFromSubject} from subject not found in directory; looked up by sender email instead.`;
      student = lookupByEmail(from, master.members);
    }
  } else {
    student = lookupByEmail(from, master.members);
    if (!student) {
      const sid = extractStudentId(body);
      if (sid) student = lookupByStudentId(sid, master.members);
    }
  }

  // 3. Resolve the group, falling back to G00 if needed.
  const intendedGroupId = student ? student.groupId : null;
  const { group, usingG00 } = resolveGroupWithFallback(intendedGroupId, master.groups);

  if (!group) {
    Logger.log(`G00 group is missing or misconfigured. Cannot process thread ${thread.getId()}.`);
    thread.addLabel(labels.error);
    return;
  }

  // Scenario 6: Student is identified but their Group ID is G00 — they have not yet
  // been placed in a real study group. Reply with guidance and skip G00 entry creation.
  if (usingG00 && student && intendedGroupId === CATCH_ALL_GROUP) {
    thread.addLabel(labels.noGroup);
    message.createDraftReply('', { htmlBody: buildG00StudentResponseHtml(), name: 'Zahid Naeem (LQM)' });
    Logger.log(`Student ${student.studentId} is assigned to G00. No G00 entry created; draft reply saved.`);
    return;
  }

  // 4. Parse lesson number — broad scan of subject first, then labelled pattern in body.
  const lessonId = extractLessonIdBroad(subject) ?? extractLessonId(body);

  // 5. Open the Drive folder and the submission sheet.
  const folderId = extractDriveFolderId(group.hwUploadFolderLink);
  const sheetId  = extractSheetId(group.hwSheetLink);
  // resolveGroupWithFallback already verified these parse correctly,
  // but guard here in case G00 itself is misconfigured.
  if (!folderId || !sheetId) {
    Logger.log(`Cannot parse folder/sheet IDs for group ${group.groupId}. Applying error label.`);
    thread.addLabel(labels.error);
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const ss     = SpreadsheetApp.openById(sheetId);
  const sheet  = ss.getSheetByName(HW_SHEET_TAB);
  if (!sheet) {
    Logger.log(`Tab "${HW_SHEET_TAB}" not found in sheet ${sheetId}. Applying error label.`);
    thread.addLabel(labels.error);
    return;
  }

  // 6. Pick a grader for this group using count-based balancing: assign to
  //    whichever grader has fewer existing rows in the sheet. Ties are broken
  //    randomly. This converges to equal distribution regardless of sample size.
  const graderName = resolveGrader(group, master.graders, sheet);

  // 7. Reject if this student has already submitted this lesson.
  if (student && lessonId !== null && isDuplicateSubmission(sheet, student.studentId, lessonId)) {
    thread.addLabel(labels.duplicate);
    message.createDraftReply('', { htmlBody: buildDuplicateResponseHtml(), name: 'Zahid Naeem (LQM)' });
    Logger.log(
      `Duplicate submission from student ${student.studentId} for lesson ${lessonId}. ` +
      `Labelled "${LABEL_DUPLICATE}", draft reply created.`
    );
    return;
  }

  // 8. Save the first PDF only and append a submission row.
  if (pdfs.length > 1) {
    Logger.log(`${pdfs.length} PDFs found from ${from}. Only the first will be saved.`);
  }
  const pdf      = pdfs[0];
  const filename = buildFilename(lessonId, student ? student.studentId : null, pdf.getName());
  const file     = folder.createFile(pdf.copyBlob().setName(filename));
  const studentName = student ? student.studentName : from.split('@')[0];

  sheet.appendRow([
    Utilities.getUuid(),                             // Submission ID (hidden import key)
    new Date(),                                      // Timestamp
    student   ? student.studentId   : '',            // Student ID
    studentName,                                     // Student Name (email username if unknown)
    lessonId !== null ? lessonId    : '',             // Lesson #
    file.getUrl(),                                   // HW File Link
    graderName,                                      // Grader
    usingG00 ? 'Incorrect submission' : 'Assigned',  // Grading Status
    '',                                              // Grade  (grader fills in)
    subjectSidComment,                               // Comments (system note if subject ID was unrecognised)
  ]);
  Logger.log(
    `Saved "${filename}" → group ${group.groupId}` +
    (student ? ` (student ${student.studentId})` : ` (unknown student, logged as "${studentName}")`)
  );

  // 9. Label the thread so it is never reprocessed.
  thread.addLabel(usingG00 ? labels.error : labels.processed);
}

// ─── Master Data Loader ────────────────────────────────────────────────────────

/**
 * Reads all three relevant tabs from the master sheet once per run
 * and returns structured lookup maps.
 */
function loadMasterData() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);

  // ── Graders tab ──────────────────────────────────────────────────────────
  // Columns: Grader ID | Name | Email | Phone | Comments
  const graderRows = ss.getSheetByName('Graders').getDataRange().getValues();
  const graders = {};
  for (let i = 1; i < graderRows.length; i++) {
    const [id, name] = graderRows[i];
    if (id) graders[String(id).trim()] = { name: String(name || '').trim() };
  }

  // ── Groups tab ───────────────────────────────────────────────────────────
  // Columns: Group ID | Group Name | Group Lead | WhatsApp Link |
  //          Group Share Folder Link | HW Upload Folder Link |
  //          HW Submission Sheet Link | Grader 1 | Grader 2 | Email Msg
  const groupRows = ss.getSheetByName('Groups').getDataRange().getValues();
  const groups = {};
  for (let i = 1; i < groupRows.length; i++) {
    const [groupId, , , , , folderLink, sheetLink, grader1, grader2] = groupRows[i];
    if (!groupId) continue;
    groups[String(groupId).trim()] = {
      groupId:          String(groupId).trim(),
      hwUploadFolderLink: String(folderLink || '').trim(),
      hwSheetLink:      String(sheetLink  || '').trim(),
      grader1:          grader1 ? String(grader1).trim() : null,
      grader2:          grader2 ? String(grader2).trim() : null,
    };
  }

  // ── Group_Members tab ────────────────────────────────────────────────────
  // Columns: Gender(0) | Student Name(1) | Age(2) | POSTALCODE(3) | FSA(4) |
  //          City(5) | Effective Email(6) | Mobile Phone F(7) | Student ID(8) |
  //          Status(9) | PC-4(10) | Attendance(11) | Group ID(12)
  const memberRows = ss.getSheetByName('Group_Members').getDataRange().getValues();
  const byEmail = {};
  const bySid   = {};
  for (let i = 1; i < memberRows.length; i++) {
    const row         = memberRows[i];
    const studentName = String(row[1]  || '').trim();
    const email       = String(row[6]  || '').trim().toLowerCase();
    const studentId   = String(row[8]  || '').trim();
    const groupId     = String(row[12] || '').trim();
    if (!studentId) continue;
    const record = { studentName, studentId, groupId };
    if (email)     byEmail[email]     = record;
    if (studentId) bySid[studentId]   = record;
  }

  return { graders, groups, members: { byEmail, bySid } };
}

// ─── Student Lookups ───────────────────────────────────────────────────────────

function lookupByEmail(email, members) {
  return members.byEmail[email.toLowerCase()] || null;
}

function lookupByStudentId(sid, members) {
  return members.bySid[String(sid)] || null;
}

// ─── Group Resolution ──────────────────────────────────────────────────────────

/**
 * Returns the group to use and whether it is the G00 fallback.
 * Falls back to G00 (usingG00: true) when:
 *   - groupId is null/empty (student not identified),
 *   - groupId is G00 (student assigned to catch-all, i.e. not yet placed in a group),
 *   - the group row is missing, or
 *   - the group's folder/sheet links cannot be parsed.
 */
function resolveGroupWithFallback(groupId, groups) {
  if (groupId && groupId !== CATCH_ALL_GROUP && groups[groupId]) {
    const g = groups[groupId];
    if (extractDriveFolderId(g.hwUploadFolderLink) && extractSheetId(g.hwSheetLink)) {
      return { group: g, usingG00: false };
    }
    Logger.log(`Group ${groupId} has invalid folder/sheet links. Falling back to G00.`);
  }

  const g00 = groups[CATCH_ALL_GROUP];
  if (!g00) {
    Logger.log(`G00 row not found. Known group IDs: ${Object.keys(groups).join(', ')}`);
    return { group: null, usingG00: true };
  }
  const g00FolderId = extractDriveFolderId(g00.hwUploadFolderLink);
  const g00SheetId  = extractSheetId(g00.hwSheetLink);
  if (!g00FolderId || !g00SheetId) {
    Logger.log(`G00 link parsing failed.`);
    Logger.log(`  hwUploadFolderLink = "${g00.hwUploadFolderLink}" → folderId = ${g00FolderId}`);
    Logger.log(`  hwSheetLink        = "${g00.hwSheetLink}" → sheetId = ${g00SheetId}`);
    return { group: null, usingG00: true };
  }
  return { group: g00, usingG00: true };
}

/**
 * Returns the grader's name for the given group.
 * When both Grader 1 and Grader 2 are set, reads only the last row of the
 * submission sheet to see who was assigned most recently, then picks the
 * other one. This gives strict alternation going forward without touching
 * historical data. If the last row's grader doesn't match either current
 * grader (e.g. after a reassignment), falls back to a random pick so that
 * alternation self-corrects from that point on with no manual intervention.
 */
function resolveGrader(group, graders, sheet) {
  const candidates = [group.grader1, group.grader2].filter(Boolean);
  if (candidates.length === 0) return '';

  const getName = id => (graders[id] && graders[id].name) ? graders[id].name : '';

  if (candidates.length === 1) return getName(candidates[0]);

  const name1 = getName(candidates[0]);
  const name2 = getName(candidates[1]);

  // Read only the last row's grader cell (Col G = column 7, 1-based).
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastGrader = String(sheet.getRange(lastRow, 7).getValue() || '').trim();
    if (lastGrader === name1) return name2;
    if (lastGrader === name2) return name1;
  }

  // Sheet is empty or last grader doesn't match either current grader — pick randomly.
  return Math.random() < 0.5 ? name1 : name2;
}

// ─── Parsing Helpers ───────────────────────────────────────────────────────────

/** Returns only PDF attachments (checks MIME type and file extension). */
function getPdfAttachments(message) {
  return message.getAttachments().filter(a => {
    const type = (a.getContentType() || '').toLowerCase();
    const name = (a.getName()        || '').toLowerCase();
    return type === 'application/pdf' || name.endsWith('.pdf');
  });
}

/** Extracts a bare email address from a "Display Name <addr>" header. */
function extractEmail(fromField) {
  const match = fromField.match(/<([^>]+)>/);
  return (match ? match[1] : fromField).trim().toLowerCase();
}

/**
 * Looks for a labelled lesson number (1–23) in text.
 * Recognises: "Lesson 5", "Lesson #5", "Lesson No. 5", "Lesson No 5".
 * Used for body text. For subject scanning use extractLessonIdBroad().
 */
function extractLessonId(text) {
  const match = /\blesson\s*(?:#|no\.?)?\s*(\d{1,2})\b/i.exec(text);
  if (!match) return null;
  const n = parseInt(match[1], 10);
  return (n >= 1 && n <= 23) ? n : null;
}

/**
 * Looks for a lesson number in the email subject with no label required.
 * Tries the labelled pattern first for precision, then falls back to any
 * standalone 1–2 digit number in range 1–23 not adjacent to other digits.
 */
function extractLessonIdBroad(text) {
  const labeled = /\blesson\s*(?:#|no\.?)?\s*(\d{1,2})\b/i.exec(text);
  if (labeled) {
    const n = parseInt(labeled[1], 10);
    if (n >= 1 && n <= 23) return n;
  }
  const pattern = /(?<!\d)(\d{1,2})(?!\d)/g;
  let match;
  while ((match = pattern.exec(text)) !== null) {
    const n = parseInt(match[1], 10);
    if (n >= 1 && n <= 23) return n;
  }
  return null;
}

/**
 * Looks for a labelled student ID (7001–7999) in text.
 * Recognises: "Student ID: 7042", "Student_ID 7042", "StudentID:7042".
 * Used for body text. For subject scanning use extractStudentIdBroad().
 */
function extractStudentId(text) {
  const match = /\bstudent[_ -]?id\s*[:#]?\s*(\d{4})\b/i.exec(text);
  if (!match) return null;
  const n = parseInt(match[1], 10);
  return (n >= 7001 && n <= 7999) ? String(n) : null;
}

/**
 * Looks for a student ID in the email subject with no label required.
 * Matches any 4-digit number in range 7001–7999 that is not adjacent to
 * other digits (so "701234" is ignored but "student7122" and "7231" both match).
 */
function extractStudentIdBroad(text) {
  const match = /(?<!\d)(7\d{3})(?!\d)/.exec(text);
  if (!match) return null;
  const n = parseInt(match[1], 10);
  return (n >= 7001 && n <= 7999) ? String(n) : null;
}

/** Extracts a folder ID from a Google Drive folder URL. */
function extractDriveFolderId(url) {
  if (!url) return null;
  const match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/** Extracts a sheet ID from a Google Sheets URL. */
function extractSheetId(url) {
  if (!url) return null;
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Builds the filename for a saved PDF.
 * Format: {LessonNo}_{StudentID}_{originalName}
 * Parts with no value are omitted.
 */
function buildFilename(lessonId, studentId, originalName) {
  const parts = [];
  if (lessonId  !== null && lessonId  !== undefined) parts.push('L' + String(lessonId).padStart(2, '0'));
  if (studentId !== null && studentId !== undefined) parts.push(String(studentId));
  parts.push(originalName || 'attachment.pdf');
  return parts.join('_');
}

// ─── Duplicate Submission Check ────────────────────────────────────────────────

/**
 * Returns true if the HW_Submissions sheet already contains a row for the
 * given studentId and lessonId. Only called when both values are known.
 */
function isDuplicateSubmission(sheet, studentId, lessonId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  // Read Student ID (col C) and Lesson # (col E) — columns 3 and 5, width 3.
  const data = sheet.getRange(2, 3, lastRow - 1, 3).getValues();
  const sid   = String(studentId).trim();
  const lid   = String(lessonId).trim();
  return data.some(row => String(row[0]).trim() === sid && String(row[2]).trim() === lid);
}

// ─── Response Templates ────────────────────────────────────────────────────────

function buildNoPdfResponseHtml() {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:680px;">

<p>We were unable to process your submission because there was no PDF attached to your email.</p>

<p><strong>Why this happened:</strong></p>
<ul style="line-height:1.9;">
  <li><strong>It was missed:</strong> You may have simply forgotten to attach the file before hitting send.</li>
  <li><strong>A shared link was used:</strong> Our system cannot access links. If you sent a Google Drive or OneDrive link, the submission will fail.</li>
  <li><strong>The file was too large:</strong> If your PDF is too large, your email provider may have automatically converted your attachment into a shared link without you realizing it.</li>
</ul>

<p><strong>How to resolve this:</strong></p>
<ul style="line-height:1.9;">
  <li><strong>Check the file size:</strong> To guarantee your email system attaches the file directly rather than turning it into a link, keep your PDF file size under 15 MB.</li>
  <li><strong>Compress if necessary:</strong> If your file is larger than 15 MB, search online for a free PDF compression tool to reduce its size.</li>
  <li><strong>Attach and resubmit:</strong> Upload the PDF as a direct attachment to a new email and send it through.</li>
</ul>

<p><em>System generated response.</em></p>

</body>
</html>`;
}

function buildDuplicateResponseHtml() {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:680px;">

<p>We received your homework submission, but it was declined because it is a duplicate of one you already sent.</p>

<p><strong>What this means for you:</strong></p>
<ul style="line-height:1.9;">
  <li>Your original submission was successfully received and has already been assigned for grading.</li>
  <li>There is no further action needed on your part.</li>
  <li>To help our system run smoothly, please ensure you only submit each lesson's homework once.</li>
</ul>

<p>Thank you for your cooperation and dedication to your studies!</p>

<p><em>System generated response.</em></p>

</body>
</html>`;
}

function buildG00StudentResponseHtml() {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:680px;">

<p>Thank you for submitting your homework assignment. Unfortunately, we are unable to accept it for grading at this time because you are not currently registered in a study group.</p>

<p><strong>Why this matters:</strong> Our homework grading service is exclusively managed and structured around study group membership.</p>

<p><strong>How to get your homework graded:</strong> To take full advantage of our grading system, you just need to join a group! Please reach out to the LQM team to discuss how you can join a study group and get your assignments on track for grading.</p>

<p>We look forward to seeing your submissions once you are plugged into a group!</p>

<p><em>System generated response.</em></p>

</body>
</html>`;
}

// ─── Label Helpers ─────────────────────────────────────────────────────────────

function ensureLabels() {
  return {
    processed:    getOrCreateLabel(LABEL_PROCESSED),
    error:        getOrCreateLabel(LABEL_ERROR),
    noAttachment: getOrCreateLabel(LABEL_NO_ATTACHMENT),
    duplicate:    getOrCreateLabel(LABEL_DUPLICATE),
    noGroup:      getOrCreateLabel(LABEL_NO_GROUP),
  };
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// ─── Trigger Setup ─────────────────────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor.
 * It removes any existing trigger for this script and creates a fresh hourly one.
 */
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processHomeworkEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('processHomeworkEmails')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly trigger installed for processHomeworkEmails.');
}
