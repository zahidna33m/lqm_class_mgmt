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
| `sendProgressReportEmails()`   | Sends emails directly | Custom menu **LQM Progress → Send Progress Report Emails** |
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
| `SUBJECT` | `LQM Al-Falah 26 Student Progress Report - Level 1` | Email subject line |
| `PC_INCLUDE_COL` | 18 (R) | Prog R. Include column |
| `PC_SENT_COL` | 19 (S) | Prog R. Sent column |
| `PC_HEADER_ANCHOR` | `Student ID` | Text that identifies the header row |
| `PC_HEADER_ANCHOR_COL` | 5 (E) | Column where the anchor lives |
| `PC_HEADER_SCAN_ROWS` | 20 | Maximum rows scanned when locating the header |
| `QUIZ_FIRST_COL` | 25 (Y) | First quiz column in `Quiz_Aggr` |
| `QUIZ_DATA_FIRST_ROW` | 8 | First student row in `Quiz_Aggr` |
| `HW_FIRST_LESSON_COL` | 7 (G) | First lesson column in `Student_HW_Grades` |
| `LEVEL_TOTAL_HW_LESSONS` | 20 | Total HW-bearing lessons in the level (lessons 4–23; lessons 1–3 have no homework). Drives the encouraging-intro stage selector. |

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
| J | Quiz Attempted % | Per-student column unused; **J2** is the Quiz Attempted % cert. threshold |
| K | Quiz Score % | Per-student column unused; **K2** is the overall Quiz Score % cert. threshold |
| L | Recent Att % | — |
| M | Book 1 Att % | Per-student column unused; **M2** is the Attendance % cert. threshold |
| N | Lessons Submitted | — |
| O | Lesson Submitted % | Per-student column unused; **O2** is the Lesson Homework Submitted % cert. threshold |
| P | Avg. Grade | Per-student column unused; **P2** is the Overall HW Grade cert. threshold |
| Q | (blank) | — |
| R | **Prog R. Include** | Checkbox — script processes only checked rows; unchecks after delivery |
| S | **Prog R. Sent** | Script writes today's date here after delivery |

#### Row 2 — Certificate Eligibility Thresholds

| Cell | Threshold |
| --- | --- |
| `J2` | Quiz Attempted % minimum |
| `K2` | Overall Quiz Score % minimum |
| `M2` | Attendance % minimum |
| `O2` | Lesson Homework Submitted % minimum |
| `P2` | Overall HW Grade minimum |

These thresholds are pulled once per run and shown in the email under each corresponding student metric. They may be tweaked as the level progresses without any code change.

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
| 3 | Submission due date |
| 4+ | Student grades |

Every column from `HW_FIRST_LESSON_COL` (G) to `getLastColumn()` is evaluated; a column is **included** only if **both** are true:

- Row 2 (lesson number) is non-blank.
- Row 3 (submission date) is a real `Date` value **strictly before today** (today taken at 00:00 in the script timezone).

This means lessons whose submission date is today or in the future are silently excluded from the report.

---

## 6. Email Template

Subject: **`LQM Al-Falah 26 Student Progress Report - Level 1`**

Body layout (HTML, max width 700px, Arial 14px):

1. **Greeting** — `Assalaamu alaykum {{title}} {{student_name_en}},`
2. **Stage intro** — one or two short paragraphs of encouraging context (see §6.6); chosen automatically by `pickReportStage_` based on how far the level has progressed.
3. **Info block** (light grey background, left blue border):
   - `Student ID:` value
   - `Student Name:` Arabic name in 28px font (omitted if Arabic name is blank)
   - `Report Date:` `MMMM d, yyyy`
4. **Section: Attendance** (blue header on grey banner, content indented 24px)
   - `Attendance:` `attended / total (pct%)`
   - `Certificate Eligibility:*` `M2` (formatted as percent)
5. **Section: Quizzes** (blue header on grey banner, content indented 24px)
   - `Quiz Attempted:` `attempted / assigned (pct%)`
   - `Certificate Eligibility:*` `J2` (formatted as percent)
   - `Quiz Scores:` table
   - `Certificate Eligibility:*` `K2` (formatted as percent — applies to the table's Overall Score %)
6. **Section: Homework**`**` (blue header on grey banner with a small `**` superscript marker, content indented 24px)
   - `Lesson Homework Submitted:` `submitted / assigned (pct%)`
   - `Certificate Eligibility:*` `O2` (formatted as percent)
   - Homework grade table
   - `Certificate Eligibility:*` `P2` (raw numeric grade — applies to the table's Overall grade)
7. **Footnotes** — small italic grey:
   - `* Certificate Eligibility values might be adjusted as we get closer to the end of this level.`
   - `** Lessons whose submission date is in future are not included in this list.`
8. **Footer** — small grey italic `System generated email.`

All five **Certificate Eligibility** lines render in a smaller (13px), italic, muted grey (`#555`) style to distinguish supplementary info from the student's own metrics.

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

| Row type | Lesson No | Grade |
| --- | --- | --- |
| Per-lesson (submitted, graded) | Lesson number from row 2 | Numeric grade |
| Per-lesson (submitted, ungraded) | Lesson number | `Yet to be graded` (raw value `-1`) |
| Per-lesson (not submitted) | Lesson number | `Not submitted` (raw value blank) |
| **Overall** (last row, bold, grey) | `Overall` | Average of graded grades, rounded to 2 decimal places (e.g. `8.67`); `—` if no graded lessons |

"Graded" for the Overall row means: the cell is a number AND not `-1`. Blank cells and `-1` are both excluded from the average.

### 6.4 Percentage formatting

`formatPct_()` accepts either:
- A decimal fraction (e.g. `0.83`) → renders as `83%`
- An already-percent number (e.g. `83`) → renders as `83%`
- Blank/null → `—`

This means the script is tolerant of either storage convention in the source sheets.

### 6.5 Stage intro (encouraging text)

A short, tone-setting block sits immediately after the greeting. It is chosen automatically based on how many homework-bearing lessons have been released so far (i.e. `hwData.lessons.length`, which already excludes lessons whose submission date is today or later) versus `LEVEL_TOTAL_HW_LESSONS` (20 — lessons 4–23).

| Stage | Condition (`released / 20`) | HW released | Lesson window | Planned send |
| --- | --- | --- | --- | --- |
| `early` | `< 0.50` | 0–9 | lessons 4–12 | After lesson 8 (5 HW released) |
| `mid`   | `< 0.75` | 10–14 | lessons 13–17 | After lesson 13 (10 HW released); flexible — any send through lesson 17 also lands here |
| `late`  | `< 1.00` | 15–19 | lessons 18–22 | After lesson 19 (16 HW released) |
| `final` | `= 1.00` | 20 | lesson 23 | After lesson 23 — **assumes operator has verified all grading is complete** before running |

Each stage contributes one short paragraph tailored to where the level is at. The paragraph uses a conditional framing: it reassures students whose numbers are below the Certificate Eligibility thresholds, and encourages students whose numbers are already at or above them. The `final` paragraph additionally mentions that Level 2 will begin after this level wraps up — a fresh opportunity for any student who didn't qualify this time.

For `early`, `mid`, and `late`, a second timeless paragraph also appears, reminding the student that the certificate is a recognition rather than the main point of the journey, and that what matters is the connection they're building with the Qur'an. The `final` stage omits this second paragraph because its first paragraph already lands on that note.

The exact wording lives in `buildStageIntroHtml_`. Update there to change copy.

The chosen stage is logged at run time (`Report stage: X (N of Y lessons released).`).

### 6.6 Ratio-line formatting

`buildRatioLine_(numerator, denominator)` is used for the Attendance, Quiz Attempted, and Lesson Homework Submitted lines. Output format:

| Inputs | Output |
| --- | --- |
| both numeric, denominator > 0 | `n / d (pct%)` — `pct` to 2 decimals |
| both numeric, denominator = 0 | `n / d` |
| one numeric, one non-numeric | `n / d` (string-concatenated) |
| both non-numeric | `—` |

---

## 7. Placeholder Mapping

| Placeholder | Source |
| --- | --- |
| `{{title}}` | `Br.` if Gender does not start with `f` (case-insensitive); `Sr.` if it does |
| `{{student_name_en}}` | `Progress_Cert!B` |
| `{{student_name_ar}}` | `Progress_Cert!G` (rendered as the "Student Name:" line in 28px; line omitted if blank) |
| `{{student_id}}` | `Progress_Cert!E` |
| `{{date_generated}}` | Today, formatted `MMMM d, yyyy` (e.g. `June 26, 2026`) in script timezone |
| `{{stage_intro_html}}` | One or two short paragraphs from `buildStageIntroHtml_(stage)`. Stage is `pickReportStage_(hwData.lessons.length)` compared against `LEVEL_TOTAL_HW_LESSONS`. See §6.5. |
| `{{attendance_line}}` | `Hn / H2 (pct%)` from `Attendance` (n = student's row); rendered via `buildRatioLine_` |
| `{{quiz_attempted_line}}` | `Mn / M3 (pct%)` from `Quiz_Aggr` (n = student's row); rendered via `buildRatioLine_` |
| `{{quiz_score_table}}` | See §6.2 |
| `{{hw_submitted_line}}` | `submittedCount / totalLessonsAssigned (pct%)`; submitted = lessons with any value other than blank/null (includes `-1` "Yet to be graded") |
| `{{homework_grade_table}}` | See §6.3 |
| `{{cert_eligibility_attendance}}` | `Progress_Cert!M2`, formatted as percent |
| `{{cert_eligibility_quiz_attempted}}` | `Progress_Cert!J2`, formatted as percent |
| `{{cert_eligibility_quiz_score}}` | `Progress_Cert!K2`, formatted as percent |
| `{{cert_eligibility_hw_submitted}}` | `Progress_Cert!O2`, formatted as percent |
| `{{cert_eligibility_hw_grade}}` | `Progress_Cert!P2`, raw numeric grade |

---

## 8. Per-Row Processing Logic

Before the per-row loop, the script reads **row 2** of `Progress_Cert` once and extracts the five certificate-eligibility thresholds from cells `J2`, `K2`, `M2`, `O2`, and `P2`. It also computes the `reportStage` (one of `early` / `mid` / `late` / `final`) from `hwData.lessons.length` versus `LEVEL_TOTAL_HW_LESSONS` and builds the corresponding `stageIntroHtml` once. All three (thresholds, stage, intro) are reused for every student in the run.

For each row from `dataFirstRow` to `lastRow`:

1. Skip if `Prog R. Include` (col R) is not truthy.
2. Read Gender, Student Name, Effective Email, Student ID, Arabic Name from row.
3. If Effective Email is blank → log `ERROR: Student ID NNNN (Name) has no Effective Email. Skipping.` and skip the row (Prog R. Include stays checked).
4. Compute title from Gender (`Sr.` if gender starts with `f`/`F`, else `Br.`).
5. Look up the student's record in `quizData.byStudentId`, `attendanceData.byStudentId`, and `hwData.byStudentId` (by Student ID). Each lookup may return undefined; tables handle missing data gracefully.
6. Build the Attendance ratio line via `buildRatioLine_(attended, totalClasses)`.
7. Build the Quiz Attempted ratio line via `buildRatioLine_(attempted, assigned)`.
8. Count lessons with any non-blank grade (including `-1`) and build the HW Submitted ratio line via `buildRatioLine_(submitted, totalLessons)`.
9. Build the Quiz Scores table and Homework grade table.
10. Build the HTML body and either `createDraft` or `sendEmail`.
11. Write today's `Date` to **Prog R. Sent** (col S).
12. Set **Prog R. Include** (col R) to `false`.

At end of run: log `Done. N report(s) drafted.` (or `sent`).

---

## 9. Custom Menu

`onOpen()` adds the **LQM Progress** menu with two items:

- **Create Progress Report Drafts** — runs `createProgressReportDrafts()` (saves to Gmail Drafts).
- **Send Progress Report Emails** — runs `sendProgressReportEmails()` (sends immediately, no draft step).

Both paths share `processProgressReports_(asDraft)`. The drafts path is the safe default for previewing; the send path is for runs you've already verified.

---

## 10. Edge-Case Behaviour

| Situation | Behaviour |
| --- | --- |
| Header row not found in first 20 rows of column E | Run aborts, logs the failure. Prog R. Include flags untouched. |
| Student row in Progress_Cert has no Effective Email | Logged and skipped; row's flags untouched so it can be fixed and re-run. |
| Student ID not found in Quiz/Attendance/HW sheet | Report still generated; tables show `Not attempted` / `Not submitted` / `—` accordingly. |
| Quiz with blank student cell | Renders as `Not attempted` with `—` percent. |
| Quiz with no class average in row 5 | Renders as `—`. |
| Homework grade of `-1` | Renders as `Yet to be graded`; excluded from Overall average. |
| Homework grade blank | Renders as `Not submitted`; excluded from Overall average. |
| HW lesson column with no submission date in row 3, or a date today/future | Column excluded from the run entirely (not shown in summary line, not shown in table). |
| Student has zero graded HW lessons | Homework Overall row shows `—`. |
| Any threshold cell (J2/K2/M2/O2/P2) is blank | The corresponding Certificate Eligibility line renders as `—` (via `formatPct_` / `formatNum_`). |
| Class avg / score-% cells stored either as decimal `0.83` or whole-number `83` | Both render as `83%`. |
| Either side of a ratio line (Attendance / Quiz Attempted / HW Submitted) is non-numeric | The numbers are still printed but the `(pct%)` suffix is omitted. |
| Both sides non-numeric | Line renders as `—`. |
| Denominator is 0 (no classes/quizzes/lessons yet) | The numbers are printed but the `(pct%)` suffix is omitted. |

---

## 11. Future Switch to Live Sending

To send instead of drafting, add a menu entry (or trigger) wired to `sendProgressReportEmails()`. No other code change is needed — both paths share `processProgressReports_(asDraft)`.
