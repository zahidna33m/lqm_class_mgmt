/**
 * Progress Report — LQM Al-Falah 26
 *
 * Bound to the Progress spreadsheet. For each row in the Progress_Cert tab
 * where "Prog R. Include" is checked, gathers attendance, quiz, and homework
 * data from external sheets and creates a per-student progress report email.
 *
 * Two delivery modes:
 *   createProgressReportDrafts()  — saves to Gmail Drafts for manual review
 *                                   before sending. Wired to the menu.
 *   sendProgressReportEmails()    — sends directly. Used later, once drafts
 *                                   have been verified.
 *
 * After each report is created, the script writes today's date into
 * "Prog R. Sent" and unchecks "Prog R. Include" on that row.
 *
 * ── Setup (one-time) ──────────────────────────────────────────────────────────
 *   1. Open the Progress spreadsheet → Extensions → Apps Script.
 *   2. Paste this file.
 *   3. Set the three external sheet IDs below if they ever change.
 *   4. Reload the spreadsheet — the "LQM Progress" menu appears.
 *
 * ── Data sources ─────────────────────────────────────────────────────────────
 *   Progress (active spreadsheet) — Progress_Cert tab
 *     A  Gender              G  Arabic Name           M  Book 1 Att %
 *     B  Student Name        H  blank                 N  Lessons Submitted
 *     C  Effective Email     I  blank                 O  Lesson Submitted %
 *     D  Mobile Phone        J  Quiz Attempted %      P  Avg. Grade
 *     E  Student ID          K  Quiz Score %          Q  blank
 *     F  Status              L  Recent Att %          R  Prog R. Include
 *                                                     S  Prog R. Sent
 *
 *   Quiz spreadsheet — Quiz_Aggr tab
 *     I  Student ID
 *     M  Taken               (M3 = total quizzes assigned)
 *     N  Taken %
 *     O  Score
 *     P  Score %             (P5 = class avg Score %)
 *     Y onwards — quiz columns; row 1 = "Book 1", row 2 = quiz IDs,
 *                              row 3 = max score, row 4 = avg grade,
 *                              row 5 = avg %, row 8+ = student scores
 *
 *   Attendance spreadsheet — Attendance tab
 *     D  Student ID          H  Book 1 (per-student attended count)
 *                            (H2 = total classes held)
 *
 *   HW spreadsheet — Student_HW_Grades tab
 *     A  Student ID
 *     G onwards — lesson columns; row 2 = lesson #, rows below = student grades
 *                 (blank → "Not submitted", -1 → "Yet to be graded")
 */

// ─── Configuration ─────────────────────────────────────────────────────────────

const QUIZ_SHEET_ID       = '1zad6BXxjIZxtZCh8LGROowHo1d2R5G-IZI_Y85FQqvg';
const ATTENDANCE_SHEET_ID = '1sudUgyWsS4xBFr3uqVXklq9CiWnKNgBgzW7zLx6h_rk';
const HW_SHEET_ID         = '1Pr2OCOhjCoxTuC7eDZmftahHERRyLgC04-kkaJC47p0';

const PROGRESS_TAB    = 'Progress_Cert';
const QUIZ_TAB        = 'Quiz_Aggr';
const ATTENDANCE_TAB  = 'Attendance';
const HW_TAB          = 'Student_HW_Grades';

const SUBJECT = 'LQM Al-Falah 26 Student Progress Report - Level 1';

// Progress_Cert column numbers (1-indexed)
const PC_INCLUDE_COL      = 18; // R
const PC_SENT_COL         = 19; // S
const PC_HEADER_ANCHOR    = 'Student ID'; // header text used to locate the header row
const PC_HEADER_ANCHOR_COL = 5;            // column E — where the anchor lives
const PC_HEADER_SCAN_ROWS = 20;            // scan no further than this when hunting

// Quiz_Aggr layout (1-indexed column positions, 0-indexed used in array math)
const QUIZ_FIRST_COL    = 25; // Y — first quiz column
const QUIZ_DATA_FIRST_ROW = 8; // first student row

// HW Student_HW_Grades layout
const HW_FIRST_LESSON_COL = 7; // G

// Total homework-bearing lessons in the level — drives the report-stage selector.
// Lessons 1–3 have no homework, so HW submissions are expected for lessons 4–23 (20 total).
// Stage thresholds (HW lessons released / LEVEL_TOTAL_HW_LESSONS):
//   early  : pct  < 0.50      (≤ 9 HW released → roughly through lesson 12)
//   mid    : pct  < 0.75      (10–14 HW released → roughly lessons 13–17)
//   late   : pct  < 1.00      (15–19 HW released → roughly lessons 18–22)
//   final  : pct == 1.00      (all 20 HW released, i.e. lesson 23 — assumes grading is complete)
const LEVEL_TOTAL_HW_LESSONS = 20;

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LQM Progress')
    .addItem('Create Progress Report Drafts', 'createProgressReportDrafts')
    .addItem('Send Progress Report Emails',  'sendProgressReportEmails')
    .addToUi();
}

// ─── Public Entry Points ────────────────────────────────────────────────────────

/** Saves one Gmail draft per flagged student. Wired to the menu. */
function createProgressReportDrafts() {
  processProgressReports_(true);
}

/** Sends one email per flagged student. Use only after drafts are verified. */
function sendProgressReportEmails() {
  processProgressReports_(false);
}

// ─── Core Logic ────────────────────────────────────────────────────────────────

function processProgressReports_(asDraft) {
  const progressSS  = SpreadsheetApp.getActiveSpreadsheet();
  const progressTab = progressSS.getSheetByName(PROGRESS_TAB);
  if (!progressTab) {
    Logger.log(`Tab "${PROGRESS_TAB}" not found in the active spreadsheet. Aborting.`);
    return;
  }

  const headerRow = findHeaderRow_(progressTab);
  if (!headerRow) {
    Logger.log(`Could not find header row "${PC_HEADER_ANCHOR}" in column ${PC_HEADER_ANCHOR_COL} ` +
               `within first ${PC_HEADER_SCAN_ROWS} rows of ${PROGRESS_TAB}. Aborting.`);
    return;
  }
  const dataFirstRow = headerRow + 1;

  const lastRow = progressTab.getLastRow();
  if (lastRow < dataFirstRow) {
    Logger.log(`No data rows in ${PROGRESS_TAB}. Nothing to do.`);
    return;
  }

  const progressData = progressTab
    .getRange(dataFirstRow, 1, lastRow - dataFirstRow + 1, PC_SENT_COL)
    .getValues();

  // Certificate-eligibility thresholds — row 2 of Progress_Cert
  const thresholdRow = progressTab.getRange(2, 1, 1, PC_SENT_COL).getValues()[0];
  const thresholds = {
    attendancePct:    thresholdRow[12], // M2
    quizAttemptedPct: thresholdRow[9],  // J2
    quizScorePct:     thresholdRow[10], // K2
    hwSubmittedPct:   thresholdRow[14], // O2
    hwAvgGrade:       thresholdRow[15], // P2
  };

  // Load supporting sheets once
  const quizData       = loadQuizData_();
  const attendanceData = loadAttendanceData_();
  const hwData         = loadHwData_();

  // Pick the encouraging intro paragraphs based on how far the level has progressed.
  // The 'final' stage fires only when all lessons are released; it assumes the operator
  // has confirmed all homework has been graded before running.
  const reportStage    = pickReportStage_(hwData.lessons.length);
  const stageIntroHtml = buildStageIntroHtml_(reportStage);
  Logger.log(`Report stage: ${reportStage} (${hwData.lessons.length} of ${LEVEL_TOTAL_HW_LESSONS} HW lessons released).`);

  const today    = new Date();
  const todayFmt = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');

  let delivered = 0;

  for (let i = 0; i < progressData.length; i++) {
    const row    = progressData[i];
    const rowNum = i + dataFirstRow;

    if (!row[17]) continue; // Prog R. Include (col R, index 17) not checked

    const studentId   = String(row[4] || '').trim();
    const studentName = String(row[1] || '').trim();
    const email       = String(row[2] || '').trim();
    const gender      = String(row[0] || '').trim();
    const arabicName  = String(row[6] || '').trim();

    if (!email) {
      Logger.log(`ERROR: Student ID ${studentId} (${studentName}) has no Effective Email. Skipping.`);
      continue;
    }

    const title         = /^f/i.test(gender) ? 'Sr.' : 'Br.';
    const studentQuiz   = quizData.byStudentId[studentId];
    const studentAtt    = attendanceData.byStudentId[studentId];
    const studentHw     = hwData.byStudentId[studentId];

    const totalClasses    = attendanceData.totalClasses;
    const attendedClasses = studentAtt ? studentAtt.attended : '';
    const attendanceLine  = buildRatioLine_(attendedClasses, totalClasses);

    const quizzesAssigned  = quizData.totalAssigned;
    const quizzesAttempted = studentQuiz ? studentQuiz.taken : 0;
    const quizAttemptedLine = buildRatioLine_(quizzesAttempted, quizzesAssigned);

    const totalLessonsAssigned = hwData.lessons.length;
    let submittedCount = 0;
    if (studentHw) {
      for (const l of hwData.lessons) {
        const g = studentHw.perLesson[l.no];
        if (!(g === '' || g === null || g === undefined)) submittedCount++;
      }
    }
    const hwSubmittedLine = buildRatioLine_(submittedCount, totalLessonsAssigned);

    const quizScoreTable = buildQuizScoreTable_(studentQuiz, quizData);
    const hwGradeTable   = buildHomeworkGradeTable_(studentHw, hwData);

    const htmlBody = buildProgressReportHtml_({
      title, studentName, arabicName, studentId, todayFmt,
      stageIntroHtml,
      attendanceLine, quizAttemptedLine, quizScoreTable,
      hwSubmittedLine, hwGradeTable, thresholds,
    });

    if (asDraft) {
      GmailApp.createDraft(email, SUBJECT, '', { htmlBody });
      Logger.log(`Draft created for student ${studentId} (${email}).`);
    } else {
      GmailApp.sendEmail(email, SUBJECT, '', { htmlBody });
      Logger.log(`Email sent to student ${studentId} (${email}).`);
    }

    progressTab.getRange(rowNum, PC_SENT_COL).setValue(today);
    progressTab.getRange(rowNum, PC_INCLUDE_COL).setValue(false);

    delivered++;
  }

  Logger.log(`Done. ${delivered} report(s) ${asDraft ? 'drafted' : 'sent'}.`);
}

// ─── Header Discovery ──────────────────────────────────────────────────────────

/**
 * Scans the first PC_HEADER_SCAN_ROWS rows of column PC_HEADER_ANCHOR_COL
 * looking for the cell that equals PC_HEADER_ANCHOR. Returns the 1-based row
 * number of that cell, or null if not found. Lets the script tolerate any
 * number of reserved/blank rows above the header.
 */
function findHeaderRow_(tab) {
  const scanRows = Math.min(PC_HEADER_SCAN_ROWS, tab.getLastRow());
  if (scanRows < 1) return null;
  const values = tab.getRange(1, PC_HEADER_ANCHOR_COL, scanRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === PC_HEADER_ANCHOR) return i + 1;
  }
  return null;
}

// ─── Report Stage ──────────────────────────────────────────────────────────────

/**
 * Returns one of 'early' | 'mid' | 'late' | 'final' based on how many lessons
 * have been released so far (i.e. whose submission date has already passed).
 *   pct  < 0.50  → early
 *   pct  < 0.75  → mid
 *   pct  < 1.00  → late
 *   pct == 1.00  → final
 * 'final' assumes the operator has verified all homework grading is complete.
 */
function pickReportStage_(lessonsReleased) {
  if (lessonsReleased >= LEVEL_TOTAL_HW_LESSONS) return 'final';
  const pct = lessonsReleased / LEVEL_TOTAL_HW_LESSONS;
  if (pct < 0.50) return 'early';
  if (pct < 0.75) return 'mid';
  return 'late';
}

/**
 * Returns the stage-specific intro HTML (one or two short paragraphs) that
 * sits between the greeting and the info block. The 'final' stage omits the
 * second paragraph because the first paragraph already closes the message.
 */
function buildStageIntroHtml_(stage) {
  const earlyP = `<p>Before you look at the numbers, a quick note. This is an <strong>early progress report</strong>. We've only covered a small portion of the lessons and quizzes for this level so far, so if any of your figures are below the Certificate Eligibility values, please don't worry about it. There's plenty of time over the coming weeks to bring them up, إن شاء الله. And if you're already at or above the thresholds, keep it up!</p>`;
  const midP   = `<p>Before you look at the numbers, a quick note. We're now around the <strong>midpoint of the level</strong>, with a good number of lessons and quizzes still ahead. If some of your figures are below the Certificate Eligibility values, there's still time to bring them up with steady effort over the coming weeks, إن شاء الله. And if you're already at or above the thresholds, keep it up!</p>`;
  const lateP  = `<p>Before you look at the numbers, a quick note. We're now in the <strong>final stretch of the level</strong>, with only a few lessons and quizzes left. If some of your figures are below the Certificate Eligibility values, there's still time to push toward them. Even a focused effort over these last weeks can shift things, إن شاء الله. And if you're already at or above the thresholds, keep it up!</p>`;
  const finalP = `<p>A quick note before you check the numbers. This is your <strong>final progress report for the level</strong>. The figures below reflect your effort across the whole term, and Certificate Eligibility is being assessed against them. Whether you've met every threshold or not, the time you've spent learning the Qur'an is itself the most important part of this journey. Once this level wraps up, we'll start Level 2, which will be a fresh opportunity to work toward the next certificate, إن شاء الله.</p>`;
  const secondP = `<p>And please remember, the certificate is a nice recognition but it isn't the main point of this journey. What really matters is the connection you're building with the Qur'an. As long as you're putting in the effort and learning, you're on the right path.</p>`;

  switch (stage) {
    case 'early': return earlyP + secondP;
    case 'mid':   return midP   + secondP;
    case 'late':  return lateP  + secondP;
    case 'final': return finalP;
  }
  return '';
}

// ─── External Sheet Loaders ────────────────────────────────────────────────────

function loadQuizData_() {
  const tab     = SpreadsheetApp.openById(QUIZ_SHEET_ID).getSheetByName(QUIZ_TAB);
  const lastRow = tab.getLastRow();
  const lastCol = tab.getLastColumn();
  const allData = tab.getRange(1, 1, lastRow, lastCol).getValues();

  // M3 (row 3, col M / index 12) = total quizzes assigned
  const totalAssigned = allData[2][12];

  // P5 (row 5, col P / index 15) = class avg Score %
  const overallClassAvg = allData[4][15];

  // Quiz definitions from columns Y onwards
  const quizzes = [];
  for (let col = QUIZ_FIRST_COL - 1; col < lastCol; col++) {
    const quizId = String(allData[1][col] || '').trim(); // row 2
    if (!quizId) continue;
    quizzes.push({
      id:         quizId,
      colIndex:   col,
      maxScore:   allData[2][col], // row 3
      avgPercent: allData[4][col], // row 5
    });
  }

  // Per-student data
  const byStudentId = {};
  for (let r = QUIZ_DATA_FIRST_ROW - 1; r < lastRow; r++) {
    const sid = String(allData[r][8] || '').trim(); // col I
    if (!sid) continue;
    const perQuiz = {};
    for (const q of quizzes) perQuiz[q.id] = allData[r][q.colIndex];
    byStudentId[sid] = {
      taken:    allData[r][12], // col M
      score:    allData[r][14], // col O
      scorePct: allData[r][15], // col P
      perQuiz,
    };
  }

  return { totalAssigned, overallClassAvg, quizzes, byStudentId };
}

function loadAttendanceData_() {
  const tab     = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID).getSheetByName(ATTENDANCE_TAB);
  const lastRow = tab.getLastRow();
  const lastCol = tab.getLastColumn();
  const allData = tab.getRange(1, 1, lastRow, lastCol).getValues();

  // H2 (row 2, col H / index 7) = total classes held
  const totalClasses = allData[1][7];

  const byStudentId = {};
  for (let r = 1; r < lastRow; r++) {
    const sid = String(allData[r][3] || '').trim(); // col D
    if (!sid) continue;
    byStudentId[sid] = { attended: allData[r][7] }; // col H
  }

  return { totalClasses, byStudentId };
}

function loadHwData_() {
  const tab     = SpreadsheetApp.openById(HW_SHEET_ID).getSheetByName(HW_TAB);
  const lastRow = tab.getLastRow();
  const lastCol = tab.getLastColumn();
  const allData = tab.getRange(1, 1, lastRow, lastCol).getValues();

  // Today at midnight — used to keep only lessons whose submission date is strictly before today.
  const todayMidnight = new Date();
  todayMidnight.setHours(0, 0, 0, 0);

  // Lesson numbers from row 2, submission dates from row 3, col G onwards.
  // Include only columns where row 3 is a Date strictly before todayMidnight.
  const lessons = [];
  for (let col = HW_FIRST_LESSON_COL - 1; col < lastCol; col++) {
    const lessonNo       = allData[1][col]; // row 2
    const submissionDate = allData[2][col]; // row 3
    if (lessonNo === '' || lessonNo === null) continue;
    if (!(submissionDate instanceof Date) || submissionDate >= todayMidnight) continue;
    lessons.push({ no: lessonNo, colIndex: col });
  }

  const byStudentId = {};
  for (let r = 1; r < lastRow; r++) {
    const sid = String(allData[r][0] || '').trim(); // col A
    if (!sid) continue;
    const perLesson = {};
    for (const l of lessons) perLesson[l.no] = allData[r][l.colIndex];
    byStudentId[sid] = { perLesson };
  }

  return { lessons, byStudentId };
}

// ─── Table Builders ────────────────────────────────────────────────────────────

function buildQuizScoreTable_(studentQuiz, quizData) {
  const rows = quizData.quizzes.map(q => {
    const raw = studentQuiz ? studentQuiz.perQuiz[q.id] : '';
    const attempted = !(raw === '' || raw === null || raw === undefined);

    const scoreCell = attempted ? `${raw} / ${q.maxScore}` : 'Not attempted';
    const pctCell   = (attempted && typeof raw === 'number' && typeof q.maxScore === 'number' && q.maxScore > 0)
      ? `${Math.round((raw / q.maxScore) * 100)}%`
      : '—';
    const avgCell   = formatPct_(q.avgPercent);

    return tableRow_([q.id, scoreCell, pctCell, avgCell]);
  });

  // Overall row — Score shows "student total / sum of max scores for attempted quizzes"
  let overallScore = '—';
  if (studentQuiz) {
    const attemptedMaxSum = quizData.quizzes.reduce((sum, q) => {
      const raw = studentQuiz.perQuiz[q.id];
      const attempted = !(raw === '' || raw === null || raw === undefined);
      return sum + (attempted && typeof q.maxScore === 'number' ? q.maxScore : 0);
    }, 0);
    overallScore = attemptedMaxSum > 0
      ? `${formatNum_(studentQuiz.score)} / ${attemptedMaxSum}`
      : formatNum_(studentQuiz.score);
  }
  const overallPct = studentQuiz ? formatPct_(studentQuiz.scorePct) : '—';
  const overallAvg = formatPct_(quizData.overallClassAvg);

  rows.push(`<tr style="background:#f7f7f7;font-weight:bold;">
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">Overall</td>
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">${overallScore}</td>
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">${overallPct}</td>
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">${overallAvg}</td>
      </tr>`);

  return wrapTable_(['Quiz ID', 'Score', 'Score %', 'Class Average'], rows);
}

function buildHomeworkGradeTable_(studentHw, hwData) {
  const rows = hwData.lessons.map(l => {
    const grade = studentHw ? studentHw.perLesson[l.no] : '';
    let display;
    if (grade === '' || grade === null || grade === undefined) display = 'Not submitted';
    else if (grade === -1)                                     display = 'Yet to be graded';
    else                                                       display = grade;
    return tableRow_([l.no, display]);
  });

  // Overall row — average of graded lessons (numeric grade, excluding -1 "Yet to be graded")
  const gradedGrades = [];
  if (studentHw) {
    for (const l of hwData.lessons) {
      const g = studentHw.perLesson[l.no];
      if (typeof g === 'number' && g !== -1) gradedGrades.push(g);
    }
  }
  const overallGrade = gradedGrades.length > 0
    ? (gradedGrades.reduce((a, b) => a + b, 0) / gradedGrades.length).toFixed(2)
    : '—';

  rows.push(`<tr style="background:#f7f7f7;font-weight:bold;">
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">Overall</td>
        <td style="padding:6px 14px;text-align:center;border-top:2px solid #999;">${overallGrade}</td>
      </tr>`);

  return wrapTable_(['Lesson No', 'Grade'], rows);
}

// ─── HTML Helpers ──────────────────────────────────────────────────────────────

function wrapTable_(headers, bodyRows) {
  const th = headers.map(h =>
    `<th style="padding:6px 14px;text-align:center;border-bottom:2px solid #999;background:#f0f0f0;">${h}</th>`
  ).join('');
  return `<table style="border-collapse:collapse;margin-top:6px;">
    <thead><tr>${th}</tr></thead>
    <tbody>
${bodyRows.join('\n')}
    </tbody>
  </table>`;
}

function tableRow_(cells) {
  const tds = cells.map(c =>
    `<td style="padding:5px 14px;text-align:center;border-bottom:1px solid #e0e0e0;">${c}</td>`
  ).join('');
  return `      <tr>${tds}</tr>`;
}

/** Formats a value as a percent. Accepts decimal fractions (0.83 → 83%) or
 *  already-percent numbers (83 → 83%). Returns '—' for empty values. */
function formatPct_(v) {
  if (v === '' || v === null || v === undefined) return '—';
  if (typeof v !== 'number') return String(v);
  if (v <= 1.0 && v >= -1.0) return `${Math.round(v * 100)}%`;
  return `${Math.round(v)}%`;
}

function formatNum_(v) {
  if (v === '' || v === null || v === undefined) return '—';
  return v;
}

/**
 * Formats a "numerator / denominator (percent%)" line. Falls back gracefully
 * when either value is non-numeric or when the denominator is 0.
 *   buildRatioLine_(2, 3)   → "2 / 3 (66.67%)"
 *   buildRatioLine_(5, 0)   → "5 / 0"
 *   buildRatioLine_('', 10) → "—"
 */
function buildRatioLine_(numerator, denominator) {
  const numIsNum = typeof numerator   === 'number';
  const denIsNum = typeof denominator === 'number';
  if (!numIsNum && !denIsNum) return '—';
  if (!numIsNum || !denIsNum) return `${numerator} / ${denominator}`;
  if (denominator <= 0)       return `${numerator} / ${denominator}`;
  const pct = (numerator / denominator * 100).toFixed(2);
  return `${numerator} / ${denominator} (${pct}%)`;
}

// ─── Email Builder ─────────────────────────────────────────────────────────────

function buildProgressReportHtml_({
  title, studentName, arabicName, studentId, todayFmt,
  stageIntroHtml,
  attendanceLine, quizAttemptedLine, quizScoreTable,
  hwSubmittedLine, hwGradeTable, thresholds,
}) {
  const sectionHeaderStyle = 'font-size:20px;color:#1565c0;background:#eeeeee;padding:8px 14px;margin-top:28px;margin-bottom:8px;';
  const sectionBodyStyle   = 'margin-left:24px;';
  const eligibilityStyle   = 'color:#555;font-style:italic;font-size:13px;';

  const studentNameLine = arabicName
    ? `<strong>Student Name:</strong> <span style="font-size:28px;vertical-align:middle;">${arabicName}</span><br>`
    : '';

  const eligibility = (value, asPct) => {
    const formatted = asPct ? formatPct_(value) : formatNum_(value);
    return `<span style="${eligibilityStyle}">Certificate Eligibility:* ${formatted}</span>`;
  };

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7;color:#222;max-width:700px;">

<p>Assalaamu alaykum ${title} ${studentName},</p>

${stageIntroHtml}

<p style="background:#f7f7f7;padding:10px 14px;border-left:4px solid #888;">
  <strong>Student ID:</strong> ${studentId}<br>
  ${studentNameLine}<strong>Report Date:</strong> ${todayFmt}
</p>

<h3 style="${sectionHeaderStyle}">Attendance</h3>
<div style="${sectionBodyStyle}">
  <p>
    <strong>Attendance:</strong> ${attendanceLine}<br>
    ${eligibility(thresholds.attendancePct, true)}
  </p>
</div>

<h3 style="${sectionHeaderStyle}">Quizzes</h3>
<div style="${sectionBodyStyle}">
  <p>
    <strong>Quiz Attempted:</strong> ${quizAttemptedLine}<br>
    ${eligibility(thresholds.quizAttemptedPct, true)}
  </p>
  <p><strong>Quiz Scores:</strong></p>
  ${quizScoreTable}
  <p style="margin-top:6px;">${eligibility(thresholds.quizScorePct, true)}</p>
</div>

<h3 style="${sectionHeaderStyle}">Homework<span style="font-size:13px;vertical-align:super;">**</span></h3>
<div style="${sectionBodyStyle}">
  <p>
    <strong>Lesson Homework Submitted:</strong> ${hwSubmittedLine}<br>
    ${eligibility(thresholds.hwSubmittedPct, true)}
  </p>
  ${hwGradeTable}
  <p style="margin-top:6px;">${eligibility(thresholds.hwAvgGrade, false)}</p>
</div>

<p style="margin-top:24px;color:#555;font-size:12px;">
  <em>* Certificate Eligibility values might be adjusted as we get closer to the end of this level.</em><br>
  <em>** Lessons whose submission date is in future are not included in this list.</em>
</p>

<p style="margin-top:8px;color:#888;font-size:12px;"><em>System generated email.</em></p>

</body>
</html>`;
}
