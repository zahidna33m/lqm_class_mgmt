/**
 * Quiz Communication — LQM Al-Falah 26
 *
 * Generates a Gmail draft and a WhatsApp message body announcing a quiz
 * assignment to class attendees. The quiz to announce is selected by ticking
 * a checkbox in the Quiz Assignments sheet; the script fills in all details
 * from that row and the matching Quiz Master row, then clears the checkbox
 * and timestamps the action.
 *
 * Run createQuizDraft() or generateWhatsAppMessage() from the "LQM Quiz Tools"
 * menu added by onOpen(), or directly from the Apps Script editor.
 *
 * ── Google Sheet dependency ───────────────────────────────────────────────────
 *   This script must be bound to (or run from) the quiz spreadsheet.
 *
 *   Tab: Quiz Master
 *     Col A  Quiz ID
 *     Col B  Lessons/Topics   (items separated by "--"; rendered as nested bullets)
 *     Col C  Quiz Link        (default link, overridden by Quiz Assignments col D)
 *     Col D  No. of Questions
 *     Col E  Allocated Time   (minutes)
 *
 *   Tab: Quiz Assignments
 *     Col A  Quiz ID          (must match a row in Quiz Master)
 *     Col B  (unused)
 *     Col C  Due Date/Time
 *     Col D  Quiz Link        (overrides Quiz Master col C when set)
 *     Col E  Email Msg        (checkbox — tick to trigger email draft creation)
 *     Col F  WhatsApp Msg     (checkbox — tick to trigger WhatsApp message)
 *     Col G  Email Sent       (timestamp written after draft is created)
 *     Col H  WhatsApp Sent    (timestamp written after message is generated)
 *
 * ── Functions ─────────────────────────────────────────────────────────────────
 *   createQuizDraft()         Finds the checked Email Msg row, builds an HTML
 *                             draft with the quiz details, saves it to Gmail
 *                             Drafts (BCC: alfalah26_attendees), and timestamps
 *                             the row.
 *   generateWhatsAppMessage() Finds the checked WhatsApp Msg row, builds the
 *                             plain-text message, and displays it in a modal
 *                             dialog for copying into WhatsApp.
 *   getTemplateContent()      Returns the message template (HTML or WhatsApp)
 *                             with {{placeholders}} for runtime substitution.
 *   buildTopicsBlock()        Parses the "--"-separated Lessons/Topics string
 *                             and renders it as nested bullets (HTML or WhatsApp).
 *   getQuizData()             Reads the checked row from Quiz Assignments and
 *                             assembles the replacements map.
 *   getSignatureFromDraft()   Fetches the sender's email signature from a Gmail
 *                             draft whose subject is "#SIGNATURE#".
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('LQM Quiz Tools')
      .addItem('Create Email Draft', 'createQuizDraft')
      .addItem('Generate WhatsApp Message', 'generateWhatsAppMessage')
      .addToUi();
}

/**
 * CONFIGURATION: The template text is defined here.
 */
function getTemplateContent(isWhatsApp = false) {
  const b = isWhatsApp ? "*" : "<b>"; 
  const be = isWhatsApp ? "*" : "</b>"; 
  
  const indent = isWhatsApp ? "   - " : '<li style="margin-left: 20px;">';
  const listStart = isWhatsApp ? "" : '<ul style="list-style-type: disc; padding-left: 20px; margin-top: 0px;">';
  const listEnd = isWhatsApp ? "" : "</ul>";
  const itemEnd = isWhatsApp ? "\n" : "</li>";

  // Quiz Link style: Large font, Bold, and Indented
  const linkStart = isWhatsApp ? "     *" : '<div style="margin-left: 40px; font-size: 18px; font-weight: bold;">';
  const linkEnd = isWhatsApp ? "*" : '</div>';
  
  let content = "Assalaamu alaykum brothers and sisters,\n\n" +
                "Quiz {{Quiz ID}} has been assigned.\n\n";

  // Section 1: Quiz Details
  content += listStart +
    indent + "Due Date & Time: " + b + "{{Day}}, {{Due Date/Time}}" + be + itemEnd +
    indent + "Time Allocated: " + b + "{{Allocated Time}} minutes" + be + itemEnd +
    indent + "Number of Questions: " + b + "{{No. of Questions}}" + be + itemEnd +
    "{{Topics Block}}" +
    listEnd;

  content += "Use the following link to access the quiz when you are ready to take it:\n" +
             linkStart + "{{Quiz Link}}" + linkEnd + "\n" + // Reduced from \n\n
             b + "Please note:\n\n" + be;

  // Section 2: Important Notes
  content += listStart +
    indent + "When you click the link, it will ask you to enter your email address. You must enter the same email you used for the Arabic class registration. No other email will work. TestMoz will then send a message containing a generated quiz link to the email address you entered. Open the email and click that link to start the quiz." + itemEnd +
    indent + "The due date and time listed above will not be extended and the quiz will lock afterwards. Please finish the quiz at your earliest convenience; don't wait until the last minute." + itemEnd +
    indent + "Once you start the quiz, you must finish it within the allocated time specified above." + itemEnd +
    indent + "The quiz is assigned to students with regular attendance. Students who have dropped out or have not maintained regular attendance in recent classes will not be able to access the quiz." + itemEnd +
    listEnd + "\n";

  content += "May Allah grant you success!";

  return isWhatsApp ? content : content.replace(/\n/g, "<br>");
}

/**
 * Parses the "--"-separated Lessons/Topics string from the Quiz Master and
 * returns a formatted nested-bullet block ready for insertion into the template.
 * For HTML: a <li> containing a nested <ul>.
 * For WhatsApp: a "   - " parent line followed by indented "• " sub-items.
 */
function buildTopicsBlock(rawTopics, isWhatsApp) {
  const items = String(rawTopics || '').split('--').map(s => s.trim()).filter(Boolean);

  if (isWhatsApp) {
    const subBullets = items.map(item => `      • ${item}`).join('\n');
    return `   - Lessons/Topics Included:\n${subBullets}`;
  }

  const liStyle  = 'margin-left: 20px;';
  const subUlStyle = 'list-style-type: disc; padding-left: 20px; margin-top: 4px;';
  const subItems = items.map(item => `<li style="${liStyle}">${item}</li>`).join('');
  return `<li style="${liStyle}">Lessons/Topics Included:<ul style="${subUlStyle}">${subItems}</ul></li>`;
}

/**
 * Fetches the signature by looking for a draft with the subject #SIGNATURE#
 */
function getSignatureFromDraft() {
  const SIGNATURE_SUBJECT = "#SIGNATURE#";
  const drafts = GmailApp.getDrafts();
  for (let i = 0; i < drafts.length; i++) {
    if (drafts[i].getMessage().getSubject() === SIGNATURE_SUBJECT) {
      return drafts[i].getMessage().getBody();
    }
  }
  return "<br><br>Zahid Naeem (LQM)"; 
}

/**
 * Shared Helper: Fetches data from checkboxes
 */
function getQuizData(columnIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assignmentsSheet = ss.getSheetByName("Quiz Assignments");
  const masterSheet = ss.getSheetByName("Quiz Master");
  const data = assignmentsSheet.getDataRange().getValues();
  
  let rowIndex = -1;
  let rowData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][columnIndex] === true) { 
      rowData = data[i];
      rowIndex = i + 1; 
      break;
    }
  }

  if (!rowData) return null;

  const quizId = rowData[0];
  const masterData = masterSheet.getDataRange().getValues();
  let mRow = masterData.find(r => r[0] == quizId);

  if (!mRow) return null;

  const due = new Date(rowData[2]);
  const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  
  return {
    rowIndex: rowIndex,
    sheet: assignmentsSheet,
    replacements: {
      "{{Quiz ID}}": quizId,
      "{{Day}}": days[due.getDay()],
      "{{Date}}": Utilities.formatDate(due, Session.getScriptTimeZone(), "MMM dd, yyyy"),
      "{{Due Date/Time}}": Utilities.formatDate(due, Session.getScriptTimeZone(), "MMM dd, yyyy, hh:mm a"),
      "{{Quiz Link}}": rowData[3] || mRow[2],
      "{{No. of Questions}}": mRow[3],
      "{{Allocated Time}}": mRow[4],
      "{{Lessons/Topics}}": mRow[1]
    }
  };
}

function createQuizDraft() {
  const data = getQuizData(4); 
  if (!data) { SpreadsheetApp.getUi().alert("No 'Email Msg' checkbox selected."); return; }

  let subject = "LQM Al-Falah 26 Quiz {{Quiz ID}} Assignment - Due by {{Day}}, {{Date}}";
  let quizContentHtml = getTemplateContent(false);
  const signatureHtml = getSignatureFromDraft();

  data.replacements["{{Topics Block}}"] = buildTopicsBlock(data.replacements["{{Lessons/Topics}}"], false);

  for (let p in data.replacements) {
    subject = subject.split(p).join(data.replacements[p]);
    quizContentHtml = quizContentHtml.split(p).join(data.replacements[p]);
  }
  
  const finalBody = `<div style="font-family: 'Trebuchet MS', Helvetica, sans-serif; color: #333; line-height: 1.4;">` + 
                    quizContentHtml + `<br><br>` + signatureHtml + `</div>`;

  try {
    GmailApp.createDraft("", subject, "", {
      htmlBody: finalBody,
      bcc: "alfalah26_attendees@lqmississauga.com",
      name: "Zahid Naeem (LQM)"
    });
    
    data.sheet.getRange(data.rowIndex, 7).setValue(new Date()); 
    data.sheet.getRange(data.rowIndex, 5).setValue(false);      
    SpreadsheetApp.getUi().alert("Email Draft Created.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.toString());
  }
}

function generateWhatsAppMessage() {
  const data = getQuizData(5); 
  if (!data) { SpreadsheetApp.getUi().alert("No 'WhatsApp Msg' checkbox selected."); return; }

  let waMessage = getTemplateContent(true);

  data.replacements["{{Topics Block}}"] = buildTopicsBlock(data.replacements["{{Lessons/Topics}}"], true);

  for (let p in data.replacements) {
    waMessage = waMessage.split(p).join(data.replacements[p]);
  }

  data.sheet.getRange(data.rowIndex, 8).setValue(new Date()); 
  data.sheet.getRange(data.rowIndex, 6).setValue(false);      

  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;">Copy and paste to WhatsApp:</p>' +
    '<textarea style="width:100%; height:200px; font-family:sans-serif;" readonly>' + waMessage + '</textarea>'
  ).setWidth(450).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'WhatsApp Message Content');
}