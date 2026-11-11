# Teams production reminders — Microsoft 365 native flow

This implementation is separate from the GitHub Actions reminder prototype and the existing timestamp Excel add-in. It uses a scheduled Teams/Power Automate flow and an Excel Office Script.

## Scheduler
Weekdays at 12:20 PM India Standard Time. The flow runs the read-only `CheckProductionTasks` script on `Production_Tracker_Workflow_Updated.xlsx`, tests whether `count > 0` and posts `message` to the existing Teams group chat containing Finny and Jawahar.

## Excel script
See `CheckProductionTasks.ts`. Its required input is `asOfDate` in ISO format (yyyy-MM-dd) supplied via:

`formatDateTime(convertTimeZone(utcNow(), 'UTC', 'India Standard Time'), 'yyyy-MM-dd')`

Use Excel Online (Business) **Run script**. For the Power Automate **Condition** (where the action's actual name is `Run_script`), use `body('Run_script')?['result']?['count']` > `0`. In the Yes branch use Microsoft Teams **Post message in a chat or channel** with `body('Run_script')?['result']?['message']`. No branch must send nothing.

## Important
Microsoft 365 connections and schedule must be created in the owner's Teams Workflows/Power Automate UI. These are not currently remotely creatable via the connected tools in this chat. The standalone GitHub Action remains inactive to avoid duplicate sends. The workbook has blank work/due dates as last checked, so the script sends nothing until legitimate dated tasks are entered.
