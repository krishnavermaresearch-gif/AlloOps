# AlloOps

AlloOps is a Google Sheets + Apps Script workflow for managing recurring operational tasks, proving completion, tracking SLA alerts, and escalating unresolved work. It is designed for teams that need a lightweight task tracker without a separate backend.

The implementation lives in the [assignment](./assignment) folder and is structured as a Google Apps Script project.

## What it does

- Creates recurring task cycles from a task schedule
- Displays assigned work in a sidebar for employees
- Validates completion evidence before marking tasks complete
- Sends manager overdue alerts and escalation emails
- Tracks acknowledgements, completion history, and task events
- Supports working-day and holiday-aware scheduling
- Provides a manager dashboard for overview and review

## Features

- Task setup and recurring cycle generation
- Custom proof requirements: `TICKET_ID`, `DAILY_CONFIRMATION`, or `TEXT`
- Manager alert acknowledgement workflow
- Leadership escalation when tasks remain incomplete or unacknowledged
- Email notification deduplication and audit trail
- Scheduled automation via Apps Script triggers
- Sample task data for quick demo setup

## Project structure

- `assignment/Code.gs` — main Apps Script logic, scheduling, automation, validation, notifications
- `assignment/Sidebar.html` — employee and manager sidebar UI
- `assignment/appsscript.json` — Apps Script project metadata and OAuth scopes
- `assignment/.gitignore` — default ignore rules for generated files

## Sheets created automatically

Running the setup creates the following sheets:

- `Config`
- `Tasks`
- `Cycles`
- `Manager Dashboard`
- `Events`
- `Notifications`

## Setup

1. Open a new Google Sheet.
2. Go to `Extensions` -> `Apps Script`.
3. Copy the files from the `assignment` folder into the Apps Script project.
4. Save the project.
5. From the Apps Script editor or the spreadsheet menu, run `setupTracker`.
6. Grant the requested permissions when prompted.
7. Review the generated sheets and update the task data.

## Configuration

The `Config` sheet stores settings such as:

- `TIME_ZONE`
- `WORKING_DAYS`
- `WORKDAY_START`
- `WORKDAY_END`
- `HOLIDAYS`
- `GENERATION_HORIZON_DAYS`

These values control cycle generation, working-day logic, and SLA calculations.

## Task data

Each task in the `Tasks` sheet includes:

- Task ID
- Enabled flag
- Task Name
- Assignee Email
- Frequency (days)
- Due Time
- Proof Type
- Proof Instructions
- Manager Email
- Leadership Email(s)
- Manager Ack SLA Hours
- Escalation SLA Hours
- Start Date

Supported frequency values are typically `1`, `7`, `15`, and `30` days.

## Usage

After setup:

- Employees use `AlloOps` -> `Open My Tasks` to view pending work and submit completion evidence.
- Managers can review team tasks and acknowledge overdue alerts from the same sidebar.
- `Generate task cycles` creates future task instances.
- `Run alerts now` executes the automation cycle for alerting and escalation.

## Notes

This is a spreadsheet-driven workflow meant to be easy to customize and deploy in a Google Workspace environment. The script uses Apps Script triggers and email notifications rather than a cloud-hosted database or API service.

## License

This project is intended for internal operational use and has no explicit license file at the moment.

## How to use the Sheet and Apps Script

1. In `Config`, set the time zone, working days, working hours, holidays, and cycle-generation horizon.
2. In `Tasks`, enter each task's assignee, frequency, proof requirement, manager and leadership email addresses, and SLAs. Keep a task disabled until its details and recipients are verified.
3. Use `AlloOps` -> `Open My Tasks` in the Sheet. Employees select a pending task, enter its Ticket ID or daily confirmation, and submit once. Managers use the same sidebar to review team tasks and acknowledge alerts.
4. `Cycles` records each task instance and its completion and acknowledgement details. `Manager Dashboard` summarizes open and overdue work; `Events` and `Notifications` record automation activity.
5. To install or repair the tracker on a new Sheet, use `AlloOps` -> `Set up task tracker` and authorize the requested permissions. This also creates the hourly `processAutomation` trigger. Do not run setup on a Sheet containing data you need to preserve.
6. Use `Generate task cycles` to create due cycles and `Run alerts now` to process overdue alerts and escalations. Verify recipients before running alerts: they can send real email. A `SENT` notification record confirms the script recorded a send, not that the recipient received or read the email.
7. Record test evidence and status in `QA_Test_Results` when available. Mark a test as passed only after checking the actual Sheet result; verify email delivery separately in the recipient's inbox.
