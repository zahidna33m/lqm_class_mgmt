/**
 * Study Group Communication — LQM Al-Falah 26
 *
 * Generates Gmail draft messages to be sent to each study group's members.
 * Each draft is BCC'd to all students in the group and includes the group
 * name, lead, member list, WhatsApp link, and shared Drive folder link.
 *
 * To trigger draft creation for a group, tick the "Email Msg" checkbox
 * (Col J) in the Groups tab, then run createStudyGroupDrafts() from the
 * Apps Script editor or via the "LQM Admin" menu inside the spreadsheet.
 * The checkbox is automatically cleared after the draft is created.
 *
 * ── Google Sheet dependency ───────────────────────────────────────────────────
 *   This script must be bound to (or run from) the Master Google Sheet.
 *
 *   Tab: Groups
 *     Col A  Group ID
 *     Col B  Group Name                (included in email subject and body)
 *     Col C  Group Lead                (included in email body)
 *     Col D  WhatsApp Link             (included in email body)
 *     Col E  Group Share Folder Link   (included in email body)
 *     Col J  Email Msg                 (checkbox — tick to trigger draft creation)
 *
 *   Tab: Group_Members
 *     Col B  Student Name              (index 1  — listed in the email body)
 *     Col G  Effective Email           (index 6  — added to BCC)
 *     Col M  Group ID                  (index 12 — used to match students to group)
 *
 * ── Functions ─────────────────────────────────────────────────────────────────
 *   createStudyGroupDrafts()   Main function. Iterates over all rows in the
 *                              Groups tab. For each group with "Email Msg"
 *                              ticked, collects the member list and email
 *                              addresses, builds an HTML email body, creates
 *                              a Gmail draft, and clears the checkbox.
 *   onOpen()                   Adds the "LQM Admin" custom menu to the
 *                              spreadsheet UI so createStudyGroupDrafts() can
 *                              be run without opening the script editor.
 */

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

    if (sendEmailTrigger === true) {
      let bccList = [];
      let groupMembers = [];
      
      for (let j = 1; j < studentData.length; j++) {
        const studentGroupId = studentData[j][12]; // Col M
        
        if (studentGroupId === groupId) {
          const studentName = studentData[j][1];  // Col B
          const studentEmail = studentData[j][6]; // Col G
          
          if (studentEmail) bccList.push(studentEmail);
          if (studentName) groupMembers.push(studentName);
        }
      }
      
      const subject = `LQM Al-Falah 26 - Study Group ${groupName}`;
      
      // Generate the numbered list with indentation
      let membersHtmlList = "<ol style='margin-left: 20px; font-weight: normal;'>";
      groupMembers.forEach(member => {
        membersHtmlList += `<li>${member}</li>`;
      });
      membersHtmlList += "</ol>";

      // HTML Body with specific bolding
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

      if (bccList.length > 0) {
        GmailApp.createDraft("", subject, "", {
          bcc: bccList.join(","),
          htmlBody: htmlBody
        });
        
        // Clear the checkbox in Column J (Email Msg)
        groupsSheet.getRange(i + 1, 10).setValue(false);
        
        Logger.log(`Draft created for: ${groupName}`);
      }
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LQM Admin')
    .addItem('1. Generate Study Group Email Drafts', 'createStudyGroupDrafts')
    .addToUi();
}
