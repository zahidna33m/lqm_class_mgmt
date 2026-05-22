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
- All configuration lives in a single **Master Google Sheet** (set via `MASTER_SHEET_ID` in `Code.gs`).

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

| Timestamp | Student ID | Student Name | Lesson # | HW File Link | Grader | Grading Status | Grade | Comments |

- Timestamp: date/time the script ran.
- Lesson ID: numeric 1–23, parsed from email subject/body.
- HW File Link: URL of the saved PDF in the group's Drive folder.
- Grader: the grader's **name** (not ID), randomly selected if both Grader 1 and Grader 2 are set.
- Grade and Comments: left blank for graders to fill in.

## Script logic (Code.gs)

1. Search Gmail for emails to the mailing list that have none of the three labels (see below).
2. **No PDF attached** → apply `HW No Attachment` label, skip. No sheet entry created.
3. **Resolve student**: look up sender's email in Group_Members → Effective Email column. If not found, scan subject + body for "Student ID: NNNN" / "Student_ID NNNN" (must be 7001–7500).
4. **Resolve group**: use student's Group ID → look up in Groups tab. If group missing or folder/sheet links are unparseable, fall back to G00.
5. **Lesson ID**: scan subject + body for patterns like "Lesson 5", "Lesson #5", "Lesson No. 5" (value must be 1–23). Leave blank if not found.
6. **Grader**: pick Grader 1 or 2 randomly (if both set); look up name from Graders tab.
7. **Save PDF**: upload to the group's Drive folder with filename `{LessonNo}_{StudentID}_{originalName}` (skip missing parts).
8. **Multiple PDFs**: save each one and create one sheet row per file.
9. **Log submission**: append row to the group's HW_Submission sheet.
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

1. Go to script.google.com, create a new project, paste `Code.gs`.
2. Set `MASTER_SHEET_ID` at the top of the file.
3. Run `installTrigger()` once from the editor to create the hourly trigger.
4. Authorise Gmail, Drive, and Sheets access when prompted.
5. Test by running `processHomeworkEmails()` manually and checking View → Logs.

## Key constraints

- The script **never creates** Drive folders or Google Sheets. If a folder/sheet is missing, it falls back to G00.
- G00 must always be configured with valid links in the Groups tab.
- The 2-hour search window (`SEARCH_WINDOW = '2h'`) is intentionally wider than the 1-hour trigger interval to absorb timing drift. Labels prevent double-processing.
