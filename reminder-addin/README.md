# Bee Team Production Reminders — isolated prototype

This directory is a **new** add-in, separate from the existing timestamp application.
The main branch and all pre-existing files are unchanged. **No emails have been sent.**

## Functionality
- Read-only task-pane preview of `Process Workflow` dates and statuses.
- Scans daily rows 11–30 and weekly rows 36–40, with activity in A, work date in B and status in F.
- Ignores blank task names, missing dates, future work and `Completed` (case-insensitive).
- Sample recipients are **finny.jason@beeteam.in** and **jawahar.muthu@beeteam.in**.
- Scheduler sample executes weekdays at 04:40 UTC = 10:10 AM IST, subject to GitHub Actions scheduling delays.
- The scheduled job is **inactive** until `ENABLE_PRODUCTION_REMINDERS=true` is configured in repository variables.

## Important prerequisites
1. The Excel taskpane manifest refers to `https://kpbee.github.io/production-timestamps/reminder-addin/taskpane.html`. **That URL does not exist on the published main branch yet**. To test, host this directory on an HTTPS URL and update a *local copy* of the manifest's SourceLocation. Do not replace existing manifests or production pages.
2. GitHub-hosted runners cannot access the company's local Windows shared drive. Register a secured, Windows **self-hosted runner**, labelled `production-reminders`, that can read the file as `WORKBOOK_PATH`; configure this runner with least privileges.
3. Assign repository variables: `PRODUCTION_WORKBOOK_PATH`, `REMINDER_LOG_PATH` (persistent, exclusive), `MAIL_SENDER_UPN`, `ENABLE_PRODUCTION_REMINDERS` and `SEND_ENABLED`.
4. Assign GitHub Actions secrets `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET` for a dedicated Microsoft Entra application with Microsoft Graph application `Mail.Send`. **Restrict the app to an approved sender mailbox** using Microsoft 365 Exchange application RBAC / applicable scoping controls. Prefer a managed certificate or federated credentials over long-lived client secrets in production.
5. Keep `SEND_ENABLED=false` initially. Review preview output and send only to authorized test recipients before setting it true.
6. Duplicate prevention uses a persistent date log; production deployment should use an atomic database or lock to avoid duplicate sends on retries/concurrent runs. Add alerting, audit logging and secret rotation before rollout.

### Current implementation limitations
- Does not infer that every empty template row is overdue: work date is required.
- Weekday-only schedule is a prototype assumption; edit the cron separately if weekends are needed.
- The two specified recipients receive one consolidated *sample* email, rather than individual associate-specific emails.
- A repo-stored Excel file would be stale; live workbook access is essential.
- Existing timestamps and all other workbook contents are never modified by these scripts.
