# Progress Report — Specification

`Progress_Report.gs` — a Google Apps Script bound to the **Progress** spreadsheet. For each flagged student, it gathers attendance, quiz, and homework data from external sheets and produces an HTML progress-report email (initially as a Gmail draft; sending mode is wired in but not menued).

---

## 1. Purpose

For every student row in `Progress_Cert` where **Prog R. Include** is checked, build a personalised HTML email that summarises:

- Identification (Student ID + English/Arabic name + date generated)
- Attendance (total classes held, absences)
- Quizzes (counts, per-quiz scores, overall totals, class averages)
- Homework (per-lesson grade)

After delivery, the script writes today's date to **Prog R. Sent** and unchecks **Prog R. Include** so the row is not re-processed on the next run.

---

## 2. Delivery Modes

| Function | Purpose | Wired to |
| --- | --- | --- |
| `createProgressReportDrafts()` | Saves Gmail drafts (no send) | Custom menu **LQM Progress → Create Progress Report Drafts** |
| `sendProgressReportEmails()`   | Sends emails directly | Not menued; reserved for when drafts are verified |
| `processProgressReports_(asDraft)` | Shared core logic | Both of the above |

---

## 3. Configuration Constants

All defined at the top of the script.

| Constant | Value | Meaning |
| --- | --- | --- |
| `QUIZ_SHEET_ID` | `1zad6BXxjIZxtZCh8LGROowHo1d2R5G-IZI_Y85FQqvg` | Quiz spreadsheet ID |
| `ATTENDANCE_SHEET_ID` | `1sudUgyWsS4xBFr3uqVXklq9CiWnKNgBgzW7zLx6h_rk` | Attendance spreadsheet ID |
| `HW_SHEET_ID` | `1Pr2OCOhjCoxTuC7eDZmftahHERRyLgC04-kkaJC47p0` | Homework spreadsheet ID |
| `PROGRESS_TAB` | `Progress_Cert` | Tab in the bound (active) spreadsheet |
| `QUIZ_TAB` | `Quiz_Aggr` | Tab in the Quiz spreadsheet |
| `ATTENDANCE_TAB` | `Attendance` | Tab in the Attendance spreadsheet |
| `HW_TAB` | `Student_HW_Grades` | Tab in the HW spreadsheet |
| `SUBJECT` | `LQM Al-Falah 26 Student Progress Report` | Email subject line |
| `PC_INCLUDE_COL` | 18 (R) | Prog R. Include column |
| `PC_SENT_COL` | 19 (S) | Prog R. Sent column |
| `PC_HEADER_ANCHOR` | `Student ID` | Text that identifies the header row |
| `PC_HEADER_ANCHOR_COL` | 5 (E) | Column where the anchor lives |
| `PC_HEADER_SCAN_ROWS` | 20 | Maximum rows scanned when locating the header |
| `QUIZ_FIRST_COL` | 25 (Y) | First quiz column in `Quiz_Aggr` |
| `QUIZ_DATA_FIRST_ROW` | 8 | First student row in `Quiz_Aggr` |
| `HW_FIRST_LESSON_COL` | 7 (G) | First lesson column in `Student_HW_Grades` |

---

## 4. Header-Row Discovery

The script tolerates any number of reserved/blank rows above the header in `Progress_Cert`.

1. Scan column `PC_HEADER_ANCHOR_COL` (E) for at most `PC_HEADER_SCAN_ROWS` (20) rows.
2. The first row whose value equals `PC_HEADER_ANCHOR` (`Student ID`) is the header row.
3. Data starts on the next row.
4. If no anchor is found, the run aborts with a log entry.

This means the layout can shift up or down freely; only the anchor text in column E must be preserved.

---

## 5. Data Sources

### 5.1 Progress (active spreadsheet) — `Progress_Cert`

| Col | Header | Used as |
| --- | --- | --- |
| A | Gender | `Br.` / `Sr.` derivation (case-insensitive starts-with `f` → `Sr.`, else `Br.`) |
| B | Student Name | English name |
| C | Effective Email | Recipient |
| D | Mobile Phone (F) | — |
| E | Student ID | Lookup key into all external sheets |
| F | Status | — |
| G | Arabic Name | "Student Name" line in the info block (28px font) |
| H | (blank) | — |
| I | (blank) | — |
| J | Quiz Attempted % | — |
| K | Quiz Score % | — |
| L | Recent Att % | — |
| M | Book 1 Att % | — |
| N | Lessons Submitted | — |
| O | Lesson Submitted % | — |
| P | Avg. Grade | — |
| Q | (blank) | — |
| R | **Prog R. Include** | Checkbox — script processes only checked rows; unchecks after delivery |
| S | **Prog R. Sent** | Script writes today's date here after delivery |

### 5.2 Quiz spreadsheet — `Quiz_Aggr`

**Per-student columns** (read at each student's row):

| Col | Index | Header | Used as |
| --- | --- | --- | --- |
| I | 8 | Student ID | Lookup key |
| M | 12 | Taken | `quizzesAttempted` |
| N | 13 | Taken % | — |
| O | 14 | Score | Overall row "Score" |
| P | 15 | Score % | Overall row "Score %" |

**Aggregate cells:**

| Cell | Meaning |
| --- | --- |
| `M3` | `quizzesAssigned` (total quizzes assigned to the class) |
| `P5` | Overall class average Score % (used in the Overall row's "Class Average" column) |

**Quiz columns (Y onwards):**

| Row | Contents |
| --- | --- |
| 1 | Section label (e.g. "Book 1") — not used by the script |
| 2 | Quiz ID (e.g. `Q1001`) |
| 3 | Max score for that quiz |
| 4 | Class average grade — not used |
| 5 | Class average percent (used as "Class Average" per-quiz) |
| 8+ | Student scores |

The script reads every column from `QUIZ_FIRST_COL` (Y) to `getLastColumn()`. Any column whose row-2 cell is blank is skipped.

### 5.3 Attendance spreadsheet — `Attendance`

| Col | Index | Header | Used as |
| --- | --- | --- | --- |
| D | 3 | Student ID | Lookup key |
| H | 7 | Book 1 | Per-student attended count |

`H2` holds the **total classes held**.

**Absences** = `totalClasses − attendedClasses`. If either value is non-numeric, displays `—`.

### 5.4 HW spreadsheet — `Student_HW_Grades`

| Col | Index | Header | Used as |
| --- | --- | --- | --- |
| A | 0 | Student ID | Lookup key |

**Lesson columns (G onwards):**

| Row | Contents |
| --- | --- |
| 2 | Lesson number |
| 3+ | Student grades |

Every column from `HW_FIRST_LESSON_COL` (G) to `getLastColumn()` is included; columns whose row-2 cell is blank are skipped.

---

## 6. Email Template

Subject: **`LQM Al-Falah 26 Student Progress Report`**

Body layout (HTML, max width 700px, Arial 14px):

1. **Greeting** — `Assalaamu alaykum {{title}} {{student_name_en}},`
2. **Intro line** — boilerplate sentence.
3. **Info block** (light grey background, left blue border):
   - `Student ID:` value
   - `Student Name:` Arabic name in 28px font (omitted if Arabic name is blank)
   - `Progress Report Date:` `MMMM d, yyyy`
4. **Section: Attendance** (blue header on grey banner, content indented 24px)
   - Total Classes Held
   - Number of Absences
5. **Section: Quizzes** (blue header on grey banner, content indented 24px)
   - Total Quizzes Assigned
   - Quizzes Attempted (with percentage suffix in brackets, e.g. `2 (66.67%)`)
   - Quiz Scores table
6. **Section: Homework** (blue header on grey banner, content indented 24px)
   - Homework grade table
7. **Footer** — small grey italic `System generated email.`

### 6.1 Section header styling

`font-size:20px; color:#1565c0; background:#eeeeee; padding:8px 14px;`

### 6.2 Quiz Scores table

Four columns: **Quiz ID**, **Score**, **Score %**, **Class Average**

| Row type | Quiz ID | Score | Score % | Class Average |
| --- | --- | --- | --- | --- |
| Per-quiz (attempted) | Quiz ID from row 2 | `studentScore / maxScore` | Computed `round(score/max × 100)%` | row 5 of that quiz column (formatted) |
| Per-quiz (not attempted) | Quiz ID | `Not attempted` | `—` | row 5 of that quiz column |
| **Overall** (last row, bold, grey) | `Overall` | `studentTotal / sum(maxScores for attempted quizzes)` | `Pn` (Score %) | `P5` (class avg %) |

### 6.3 Homework table

Two columns: **Lesson No**, **Grade**

- Blank grade → `Not submitted`
- Grade of `-1` → `Yet to be graded`
- Otherwise the numeric grade

### 6.4 Percentage formatting

`formatPct_()` accepts either:
- A decimal fraction (e.g. `0.83`) → renders as `83%`
- An already-percent number (e.g. `83`) → renders as `83%`
- Blank/null → `—`

This means the script is tolerant of either storage convention in the source sheets.

---

## 7. Placeholder Mapping

| Placeholder | Source |
| --- | --- |
| `{{title}}` | `Br.` if Gender does not start with `f` (case-insensitive); `Sr.` if it does |
| `{{student_name_en}}` | `Progress_Cert!B` |
| `{{student_name_ar}}` | `Progress_Cert!G` (rendered as the "Student Name:" line in 28px; line omitted if blank) |
| `{{student_id}}` | `Progress_Cert!E` |
| `{{date_generated}}` | Today, formatted `MMMM d, yyyy` (e.g. `June 26, 2026`) in script timezone |
| `{{total_classes_held}}` | `Attendance!H2` |
| `{{absences}}` | `H2 − Hn` (n = student's row) |
| `{{quizzes_assigned}}` | `Quiz_Aggr!M3` |
| `{{quizzes_attempted}}` | `Quiz_Aggr!Mn` plus ` (XX.XX%)` suffix (= attempted / assigned × 100, two decimal places) |
| `{{quiz_score_table}}` | See §6.2 |
| `{{homework_grade_table}}` | See §6.3 |

---

## 8. Per-Row Processing Logic

For each row from `dataFirstRow` to `lastRow`:

1. Skip if `Prog R. Include` (col R) is not truthy.
2. Read Gender, Student Name, Effective Email, Student ID, Arabic Name from row.
3. If Effective Email is blank → log `ERROR: Student ID NNNN (Name) has no Effective Email. Skipping.` and skip the row (Prog R. Include stays checked).
4. Compute title from Gender (`Sr.` if gender starts with `f`/`F`, else `Br.`).
5. Look up the student's record in `quizData.byStudentId`, `attendanceData.byStudentId`, and `hwData.byStudentId` (by Student ID). Each lookup may return undefined; tables handle missing data gracefully.
6. Compute absences (or `—` if non-numeric).
7. Compute the quizzes-attempted percentage suffix.
8. Build the Quiz Scores table and Homework table.
9. Build the HTML body and either `createDraft` or `sendEmail`.
10. Write today's `Date` to **Prog R. Sent** (col S).
11. Set **Prog R. Include** (col R) to `false`.

At end of run: log `Done. N report(s) drafted.` (or `sent`).

---

## 9. Custom Menu

`onOpen()` adds:

- **LQM Progress** → **Create Progress Report Drafts**

The send function is intentionally not menued — it should only be wired up after drafts have been confirmed accurate.

---

## 10. Edge-Case Behaviour

| Situation | Behaviour |
| --- | --- |
| Header row not found in first 20 rows of column E | Run aborts, logs the failure. Prog R. Include flags untouched. |
| Student row in Progress_Cert has no Effective Email | Logged and skipped; row's flags untouched so it can be fixed and re-run. |
| Student ID not found in Quiz/Attendance/HW sheet | Report still generated; tables show `Not attempted` / `Not submitted` / `—` accordingly. |
| Quiz with blank student cell | Renders as `Not attempted` with `—` percent. |
| Quiz with no class average in row 5 | Renders as `—`. |
| Homework grade of `-1` | Renders as `Yet to be graded`. |
| Homework grade blank | Renders as `Not submitted`. |
| Class avg / score-% cells stored either as decimal `0.83` or whole-number `83` | Both render as `83%`. |
| `Quiz_Aggr!M3` or `H2` non-numeric | Displayed verbatim; absences computation falls back to `—`. |

---

## 11. Future Switch to Live Sending

To send instead of drafting, add a menu entry (or trigger) wired to `sendProgressReportEmails()`. No other code change is needed — both paths share `processProgressReports_(asDraft)`.
