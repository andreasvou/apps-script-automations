# Google Apps Script automations

Small, reusable automations for training teams and small offices that work in Google Workspace.
They are simplified, rewritten examples of the kind of automations I build: no real data and no employer code.

| Folder | What it does |
|---|---|
| [`sheet-email-reminders`](sheet-email-reminders/Code.gs) | Sends a reminder email to each learner N days before their session, from a Google Sheet, and marks each row so nobody gets the same email twice. Has a dry-run mode. |
| [`enrolment-data-checker`](enrolment-data-checker/Code.gs) | Checks an enrolment sheet before a course starts: empty required fields, invalid emails and duplicate learners. Highlights the cells and writes a "Checks" report. |
| [`sample-data`](sample-data) | Fake CSV data (example.com addresses) to try the scripts. |

## How to try them
1. Create a new Google Sheet and import the matching CSV from `sample-data` (File › Import).
2. Rename the sheet tab to `Sessions` or `Enrolments`.
3. Open **Extensions › Apps Script**, paste the `Code.gs` file and save.
4. Run the main function once and accept the permissions.
   - `sheet-email-reminders`: keep `DRY_RUN: true` first and check the log.
   - `enrolment-data-checker`: reload the sheet and use the **Data checks** menu.

## Why
Repetitive checks and reminders take time every week and are easy to get wrong by hand.
A few lines of Apps Script do them the same way every time, so the team can focus on learners.

---
Andreas Voulgaris · [LinkedIn](https://www.linkedin.com/in/andreas-voulgaris-728ab6234)
