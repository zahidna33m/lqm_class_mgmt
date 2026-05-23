/**
 * Group & Grader Communication — LQM Al-Falah 26
 *
 * Two sets of Gmail draft generators for the Master Google Sheet:
 *
 *   1. createStudyGroupDrafts()
 *      Sends group assignment emails to study group members. Triggered by
 *      the "Email Msg" checkbox (Col J) in the Groups tab.
 *
 *   2. createGraderDrafts()
 *      Sends grading assignment emails to homework graders. Triggered by
 *      the "Email Msg" checkbox (Col G) in the Graders tab.
 *
 * Both functions are accessible from the "LQM Admin" menu added by onOpen().
 *
 * ── Google Sheet dependency ───────────────────────────────────────────────────
 *   This script must be bound to (or run from) the Master Google Sheet.
 *
 *   Tab: Groups
 *     Col A  Group ID
 *     Col B  Group Name
 *     Col C  Group Lead
 *     Col D  WhatsApp Link
 *     Col E  Group Share Folder Link
 *     Col H  Grader 1                  (Grader ID)
 *     Col I  Grader 2                  (Grader ID, optional)
 *     Col J  Email Msg                 (checkbox — tick to trigger study group draft)
 *
 *   Tab: Group_Members
 *     Col B  Student Name              (index 1)
 *     Col G  Effective Email           (index 6)
 *     Col M  Group ID                  (index 12)
 *
 *   Tab: Graders
 *     Col A  Grader ID
 *     Col B  Name
 *     Col C  Email
 *     Col D  Phone
 *     Col E  Comments
 *     Col F  M/F                       ("M" → "Br." prefix, "F" → "Sr." prefix)
 *     Col G  Email Msg                 (checkbox — tick to trigger grader draft)
 */

// ─── Study Group Drafts ────────────────────────────────────────────────────────

function createStudyGroupDrafts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupsSheet = ss.getSheetByName("Groups");
  const studentSheet = ss.getSheetByName("Group_Members");

  const groupsData = groupsSheet.getDataRange().getValues();
  const studentData = studentSheet.getDataRange().getValues();

  for (let i = 1; i < groupsData.length; i++) {
    const groupId              = groupsData[i][0];
    const groupName            = groupsData[i][1];
    const groupLead            = groupsData[i][2];
    const whatsAppLink         = groupsData[i][3];
    const groupShareFolderLink = groupsData[i][4];
    const sendEmailTrigger     = groupsData[i][9];

    if (sendEmailTrigger !== true) continue;

    let bccList = [];
    let groupMembers = [];

    for (let j = 1; j < studentData.length; j++) {
      const studentGroupId = studentData[j][12]; // Col M

      if (studentGroupId === groupId) {
        const studentName  = studentData[j][1];  // Col B
        const studentEmail = studentData[j][6];  // Col G

        if (studentEmail) bccList.push(studentEmail);
        if (studentName)  groupMembers.push(studentName);
      }
    }

    if (bccList.length === 0) continue;

    const subject = `LQM Al-Falah 26 - Study Group ${groupName}`;

    let membersHtmlList = "<ol style='margin-left: 20px; font-weight: normal;'>";
    groupMembers.forEach(member => { membersHtmlList += `<li>${member}</li>`; });
    membersHtmlList += "</ol>";

    const htmlBody = `
      <p>Assalaamu alaykum,</p>
      <p>You are a member of the following LQM Al-Falah 26 study group.</p>
      <p>
        Group Name: <strong>${groupName}</strong><br>
        Group Lead: ${groupLead}<br>
        Group Members:
        ${membersHtmlList}
        WhatsApp Group Link: <a href="${whatsAppLink}">${whatsAppLink}</a>
      </p>
      <p>Join the group using the above link.</p>
      <p>
        Study Group Shared Folder: <a href="${groupShareFolderLink}">${groupShareFolderLink}</a><br><br>
        Please bookmark this link in your browser as you will need to access it to check your homework grade.
      </p>
    `;

    GmailApp.createDraft("", subject, "", { bcc: bccList.join(","), htmlBody });

    // Clear the Email Msg checkbox (Col J = column 10, 1-based).
    groupsSheet.getRange(i + 1, 10).setValue(false);

    Logger.log(`Draft created for: ${groupName}`);
  }
}

// ─── Grader Assignment Drafts ──────────────────────────────────────────────────

function createGraderDrafts() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const gradersSheet = ss.getSheetByName('Graders');
  const groupsSheet  = ss.getSheetByName('Groups');

  const gradersData = gradersSheet.getDataRange().getValues();
  const groupsData  = groupsSheet.getDataRange().getValues();

  // Grader ID → name lookup (used when building the "Assigned Graders" column).
  const graderNames = {};
  for (let i = 1; i < gradersData.length; i++) {
    const id = String(gradersData[i][0] || '').trim();
    if (id) graderNames[id] = String(gradersData[i][1] || '').trim();
  }

  let draftsCreated = 0;

  for (let i = 1; i < gradersData.length; i++) {
    const row      = gradersData[i];
    const graderId = String(row[0] || '').trim(); // Col A
    const name     = String(row[1] || '').trim(); // Col B
    const email    = String(row[2] || '').trim(); // Col C
    const gender   = String(row[5] || '').trim(); // Col F
    const emailMsg = row[6];                      // Col G — checkbox boolean

    if (!emailMsg || !graderId || !email) continue;

    // Find every group where this grader appears as Grader 1 or Grader 2.
    const assignedGroups = [];
    for (let j = 1; j < groupsData.length; j++) {
      const grader1Id = String(groupsData[j][7] || '').trim(); // Col H
      const grader2Id = String(groupsData[j][8] || '').trim(); // Col I
      if (grader1Id !== graderId && grader2Id !== graderId) continue;
      assignedGroups.push({
        groupId:          String(groupsData[j][0] || '').trim(), // Col A
        groupName:        String(groupsData[j][1] || '').trim(), // Col B
        sharedFolderLink: String(groupsData[j][4] || '').trim(), // Col E
        grader1Id,
        grader2Id,
      });
    }

    if (assignedGroups.length === 0) {
      Logger.log(`Grader ${graderId} (${name}): no groups assigned. Skipping.`);
      continue;
    }

    const prefix    = gender === 'F' ? 'Sr.' : 'Br.';
    const tableHtml = buildGroupTable(assignedGroups, graderNames);
    const htmlBody  = buildGraderEmailHtml(prefix, name, tableHtml);

    GmailApp.createDraft(
      email,
      'LQM Al-Falah26 HW Grading Study Group Assignment',
      '',
      { htmlBody }
    );
    Logger.log(`Draft created for ${name} (${email}).`);

    // Clear the Email Msg checkbox (Col G = column 7, 1-based).
    gradersSheet.getRange(i + 1, 7).setValue(false);
    draftsCreated++;
  }

  Logger.log(`Done. ${draftsCreated} grader draft(s) created.`);
}

/**
 * Renders the grader's group assignment as a bordered HTML table.
 * Columns: Group ID | Group Name | Assigned Graders | Shared Folder.
 */
function buildGroupTable(groups, graderNames) {
  const thStyle = 'padding:8px 14px;border:1px solid #ccc;text-align:left;background:#f2f2f2;font-weight:bold;';
  const tdStyle = 'padding:8px 14px;border:1px solid #ccc;';

  const rows = groups.map(g => {
    const assignedNames = [g.grader1Id, g.grader2Id]
      .filter(Boolean)
      .map(id => graderNames[id] || id)
      .join(', ');

    const folderCell = g.sharedFolderLink
      ? `<a href="${g.sharedFolderLink}">Open Folder</a>`
      : '—';

    return `<tr>
        <td style="${tdStyle}">${g.groupId}</td>
        <td style="${tdStyle}">${g.groupName}</td>
        <td style="${tdStyle}">${assignedNames}</td>
        <td style="${tdStyle}">${folderCell}</td>
      </tr>`;
  }).join('\n');

  return `<table style="border-collapse:collapse;font-size:14px;margin:12px 0;">
    <thead>
      <tr>
        <th style="${thStyle}">Group ID</th>
        <th style="${thStyle}">Group Name</th>
        <th style="${thStyle}">Assigned Graders</th>
        <th style="${thStyle}">Shared Folder</th>
      </tr>
    </thead>
    <tbody>
      ${rows}
    </tbody>
  </table>`;
}

/**
 * Assembles the full HTML email body for the grader assignment message.
 */
function buildGraderEmailHtml(prefix, name, groupTable) {
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:720px;">

<p>Assalaamu alaykum ${prefix} ${name},</p>

<p>Jazak Allah Khair for your help with grading the homework for the Al-Falah 26 class. It means a lot to us and to the students who will benefit immensely from your effort, إن شاء الله.</p>

<p>As mentioned in our presentation, homework grading will be assigned by study groups.</p>

<p>Here is your group assignment:</p>

${groupTable}

<p>Group Shared Folder contains an HW Submissions Google Sheet that lists every student's homework. It also shows the assigned grader for each homework assignment. This is where graders pick up the grading task and capture the grade and their comments. Here's a high-level description of how graders may use this Google Sheet.</p>

<ol style="line-height:1.9;">
  <li>Graders should first find the row assigned to them that is "not graded," meaning it has no value in the Grade column and the Grading Status is <strong>Assigned</strong>.</li>
  <li>As the grader starts working on a row, they should change the Grading Status to <strong>"Grading in progress."</strong></li>
  <li>Then, click on the HW File Link cell in that row to open the PDF file containing the submitted homework.</li>
  <li>Students must include their Student ID in the file. Confirm if the Student ID and the Lesson # submitted is the same as captured in the respective cells of the Google Sheet. If they do not match, correct the Google Sheet.</li>
  <li>Review the complete PDF for all exercise questions and assign a grade value (0, 1, 2, 3 or 4).</li>
  <li>Capture the value in the Grade cell of that row. There should be no other value other than 0, 1, 2, 3 or 4.</li>
  <li>Capture any observations you'd like to share with the student in the Comments column. This is optional.</li>
  <li>Once the homework grading is complete, change the Grading Status to <strong>"Grading complete"</strong>.</li>
  <li>Please do not change the Grade or the Comments for any row marked as <strong>"Grading complete."</strong></li>
  <li>After reviewing multiple homework assignments for the same lesson, if you have any advice to share with the teacher, please send it in a separate email to <a href="mailto:zahid.naeem@lqmississauga.com">zahid.naeem@lqmississauga.com</a>.</li>
</ol>

<p>
  Please <a href="https://chat.whatsapp.com/EZxLATkRIB57SG1WQxAgW7?mode=gi_t">join the WhatsApp group</a> to participate in the homework grading discussion and to stay informed about what's happening with the homework submissions.
</p>

<p>
  Please review the <a href="https://drive.google.com/file/d/1p0cyj5ADgPR_MXsneLPIbMgg1ycU5yNC/view?usp=drive_link">Homework Grading Guide</a> before you start grading.
</p>

<p>Regards,</p>

</body>
</html>`;
}

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LQM Admin')
    .addItem('1. Generate Study Group Email Drafts', 'createStudyGroupDrafts')
    .addItem('2. Generate Grader Assignment Email Drafts', 'createGraderDrafts')
    .addToUi();
}
