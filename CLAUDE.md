# Alfalah26 Homework Submission Processor

## What this project is

A Google Apps Script that watches a Gmail mailing list for student homework submissions (PDF attachments), saves each PDF to the correct group's Google Drive folder, and logs the submission in that group's Google Sheet.

The script lives in `HW_Submission_Intake.gs` and is deployed as a standalone Google Apps Script project at script.google.com. It runs on an hourly time-based trigger.

## Mailing list being watched

`alfalah26_hw@lqmississauga.com`

## Class structure

- ~150 students divided into ~20–25 study groups of 6–12 students each.
- Each group has a shared Google Drive folder and a homework submission Google Sheet inside it.
- 1–2 graders are assigned per group.
- All configuration lives in a single **Master Google Sheet** (set via `MASTER_SHEET_ID` in `HW_Submission_Intake.gs` and `HW_Submission_Aggregator.gs`).

## Master Google Sheet — tabs and columns

### Graders tab
| Grader ID | Name | Email | Phone | Comments |

### Groups tab
| Group ID | Group Name | Group Lead | WhatsApp Link | Group Share Folder Link | HW Upload Folder Link | HW Submission Sheet Link | Grader 1 | Grader 2 | Email Msg |

- "Group Share Folder Link": shared Drive folder link sent to students in group emails.
- "HW Upload Folder Link": Drive folder where homework PDFs are saved by the submission script.
- "HW Submission Sheet Link": Google Sheet where submission rows are logged.
- All three folder/sheet columns contain full Google Drive / Sheets URLs.
- "Grader 1" and "Grader 2" contain **Grader ID** values (looked up in the Graders tab to get names).
- "Email Msg": checkbox — when ticked, `createStudyGroupDrafts()` generates a draft for that group and clears it.
- **G00** is the catch-all group. It must have valid folder and sheet links at all times.

### Group_Members tab
| Gender | Student Name | Age | POSTALCODE | FSA | City | Effective Email | Mobile Phone (F) | Student ID | Status | PC-4 | Attendance | Group ID |

- Column indices (0-based): Student Name=1, Effective Email=6, Student ID=8, Group ID=12.
- "Effective Email" is the address students send homework from.
- Student IDs are numeric, range 7001–7999.

## Per-group Homework Submission Sheet

Each group folder contains one Google Sheet. It has a single tab named **HW_Submission** with these columns:

| Submission ID (hidden) | Timestamp | Student ID | Student Name | Book | Lesson # | HW File Link | Grader | Grading Status | Grade | Comments |

- Timestamp: date/time the script ran.
- Book: set by `resolveBook()` (Books_Config.gs) from the email's own date and the Master sheet's `Book_HW` tab — never from what the student writes or the lesson number. The duplicate check is Book + Student ID + Lesson #.
- Lesson ID: numeric 1–23, parsed from email subject/body.
- HW File Link: URL of the saved PDF in the group's Drive folder.
- Grader: the grader's **name** (not ID), randomly selected if both Grader 1 and Grader 2 are set.
- Grade and Comments: left blank for graders to fill in.

## Script logic (HW_Submission_Intake.gs)

1. Search Gmail for emails to the mailing list that have none of the three labels (see below).
2. **No PDF attached** → apply `HW No Attachment` label, skip. No sheet entry created.
3. **Resolve student**: look up sender's email in Group_Members → Effective Email column. If not found, scan subject + body for "Student ID: NNNN" / "Student_ID NNNN" (must be 7001–7500).
4. **Resolve group**: use student's Group ID → look up in Groups tab. If group missing or folder/sheet links are unparseable, fall back to G00.
5. **Lesson ID**: scan subject + body for patterns like "Lesson 5", "Lesson #5", "Lesson No. 5" (value must be 1–23). Leave blank if not found. Then resolve the Book from the email date (see above); the saved filename is prefixed `B{n}_`.
6. **Grader**: pick Grader 1 or 2 randomly (if both set); look up name from Graders tab.
7. **Save PDF**: upload to the group's Drive folder with filename `B{n}_L{nn}_{StudentID}_{originalName}` (skip missing parts).
8. **Multiple PDFs**: only the first PDF is saved and logged; any others are ignored (the log notes it).
9. **Log submission**: append row to the group's HW_Submission sheet. If the sheet's column E header is not "Book" (not yet migrated), the thread is left unlabelled and retried next run.
10. **Label the thread**:
    - Normal success → `HW_Submitted`
    - Mapped to G00 (catch-all fallback) → `HW Submission Error`
    - No PDF → `HW No Attachment`
    - Any labelled thread is never reprocessed.

## Gmail labels used

| Label | Meaning |
|---|---|
| `HW_Submitted` | Processed successfully into correct group |
| `HW Submission Error` | Fell back to G00 (unknown student or bad group config) |
| `HW No Attachment` | Email had no PDF — needs manual follow-up |

Labels are created automatically by the script on first run if they don't exist.

## Deployment

Two **standalone** Apps Script projects (created at script.google.com, not from inside a spreadsheet), plus per-spreadsheet bound scripts for the communication tools. Each file lives in exactly one project, except `Books_Config.gs`, which is in both. Do not leave backup copies or an old `Code.gs` in a project: a second file declaring the same top-level name (e.g. `MASTER_SHEET_ID`) makes every run fail with `SyntaxError: Identifier ... has already been declared`, and triggers fail silently.

| Project | Files | Triggers |
|---|---|---|
| **HW Intake** | `HW_Submission_Intake.gs`, `Books_Config.gs` | `installTrigger()` (hourly) |
| **HW Aggregator** | `HW_Submission_Aggregator.gs`, `Books_Config.gs`, `Grader_Status_Report.gs`, `Book_Migration.gs` | `installAggregatorTrigger()` (hourly); `installGraderStatusTriggers()` (Tue + Fri 5 pm) only when grader reminders should go out |

Setup, per project:
1. Paste the files above and set `MASTER_SHEET_ID` (and, in the intake project, check `CLOCK_OVERRIDE = null` and `SEARCH_WINDOW = '2h'`).
2. Authorise Gmail, Drive, and Sheets access when prompted.
3. Test by running `processHomeworkEmails()` / `aggregateHomeworkSubmissions()` manually and checking the execution log.
4. Install the trigger(s), then set each trigger's failure notification to "Notify me immediately".

Test projects: use a test Master sheet (set `MASTER_SHEET_ID` to it), test aggregates in its `Book_HW` tab, and never link the real aggregates there. The Master sheet's timezone must be `America/New_York` ("(GMT-05:00) Eastern Time"); the Sheets UI's "Toronto" entry is stored as `America/Nassau` and fails the check.

Scripts bound to their own spreadsheet (they use `getActiveSpreadsheet()`): `Group_Comm.gs`, `Quiz_Comm.gs`, and `Progress_Report.gs` (Progress & Certificates sheet).

## Progress report

`Progress_Report.gs` emails each ticked student a progress report. The stage (early / mid / late / final) is picked from how many homework lessons have passed their submission date. The **final** report uses a per-student scenario read from the Progress_Cert tab: column W = "Y" → A1 (met all criteria), else column Y = "Y" → A2 (certificate by manual override), else B (did not qualify); subject is "Final Progress Report". `createProgressReportDrafts()` only saves Gmail drafts and never touches the sheet; `sendProgressReportEmails()` sends, then sets the Sent date (col S) and clears Include (col R). See `Progress_Report_Spec.md`.

## Key constraints

- The script **never creates** Drive folders or Google Sheets. If a folder/sheet is missing, it falls back to G00.
- G00 must always be configured with valid links in the Groups tab.
- The 2-hour search window (`SEARCH_WINDOW = '2h'`) is intentionally wider than the 1-hour trigger interval to absorb timing drift. Labels prevent double-processing.

## Books and aggregates

- **Master sheet `Book_HW` tab** — columns `Book #`, `Start Date`, `Aggregate Sheet Link`, one row per book. An email belongs to the latest book whose Start Date has passed (before the first start date: the first book). Book 1's Start Date may be blank. Eastern time; the Master sheet's timezone must be America/New_York (File → Settings → "(GMT-05:00) Eastern Time"; the UI's "Toronto" entry is stored as America/Nassau and fails the check). Book 2 starts 2026-11-14 00:00 (Book 1 accepted through 23:59:59 on Nov 13). The scripts stop with an error if the tab is missing or invalid, rather than guess.
- **Starting a new book = add one row to `Book_HW`** (and create its empty aggregate spreadsheet). No script change. A blank Aggregate Sheet Link stops aggregating that book.
- One aggregate spreadsheet per book; the aggregator routes each source row by its Book column (blank = Book 1). Group sheets are shared by all books.
- `Grader_Status_Report.gs` reports on one book, `REPORT_BOOK`; change that number when a new book takes over grading.
- `Books_Config.gs` must be in both Apps Script projects (intake, and the aggregator project). `Book_Migration.gs` is one-time: it inserts the Book column (before Lesson #) in all group sheets and aggregates and backfills it from each row's Timestamp. Run `migrateBookColumnDryRun()` first.
- `CLOCK_OVERRIDE` in the intake script is for testing only and must be `null` when live.
