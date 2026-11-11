# Aurora Live Invite Importer

A Google Apps Script that adds Aurora Live 1-to-1 meeting invites to Google Calendar.

## Why it's needed

Aurora Live sends its 1-to-1 invites from Outlook with the attendee list hidden. That leaves out the guest entry Gmail needs to recognise an invite. Outlook users get the meetings automatically. In Google Workspace they arrive as ordinary emails and never reach the calendar.

## What it does

- Checks Gmail every 10 minutes for emails from `calendar@auroralive.com`.
- Reads the invite inside each email and adds the meeting to your primary calendar, with the Teams join link as the location.
- Keeps Aurora Live's meeting ID, so later reschedules and cancellations update or remove the same event instead of creating duplicates.
- Ignores invites older than the one already in your calendar.
- Labels each processed email **Aurora Live/Imported**.

## Install

| Step | Action |
|---|---|
| 1 | Open [script.google.com](https://script.google.com) with your work account and click **New project**. |
| 2 | Replace the contents of `Code.gs` with `aurora_live_invite_importer.gs` and save. |
| 3 | In the sidebar, click **+** next to **Services**, add **Google Calendar API** and keep the name `Calendar`. |
| 4 | *(Optional)* Under **Project Settings**, set the time zone to `Europe/Stockholm`. |
| 5 | Select `importAuroraInvites` and click **Run**. This first run is a test run, because the script ships with `DRY_RUN: true`. |
| 6 | Allow access to Gmail and Calendar. If you see *"Google hasn't verified this app"*, click **Advanced → Go to … (unsafe)**. |
| 7 | Check the log: lines starting `[dry run] would add/update …` show what the script would do. |
| 8 | Set `DRY_RUN: false`, save, and run `importAuroraInvites` again. |
| 9 | Run `installTrigger` to start the 10-minute schedule. |

> **About the Gmail permission:** Google describes it broadly ("read, compose, send and delete…") because Apps Script has no read-only option for this. The script only reads messages and adds a label.

## Use

Once the schedule is running, there's nothing to do. Past runs and their logs are under **Executions** in the sidebar.

| Function | Purpose |
|---|---|
| `importAuroraInvites` | Runs an import now. |
| `installTrigger` | Starts the schedule, replacing any existing one. |
| `uninstallTrigger` | Stops the schedule. |
| `resetProcessed` | Clears the list of handled emails so the next run checks everything in the search window again. |

## Settings

Change these in `CONFIG` at the top of the script.

| Setting | Default | Notes |
|---|---|---|
| `SEARCH_QUERY` | `from:calendar@auroralive.com newer_than:45d` | Which emails to check. |
| `CALENDAR_ID` | `primary` | Which calendar to add events to. |
| `LABEL_NAME` | `Aurora Live/Imported` | Set to `''` to turn off labelling. |
| `DRY_RUN` | `true` | `true` means log only, change nothing. |
| `TRIGGER_MINUTES` | `10` | Allowed values: 1, 5, 10, 15 or 30. |

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| *"This app is blocked"* | Your Workspace admin has restricted Apps Script. Ask them to allow it under **Security → API controls**. |
| `Calendar is not defined` | The Google Calendar API service hasn't been added (install step 3). |
| Error updating an event you added by hand | Google may not let the script change events where you're a guest. Leave the event as it is or delete it, then run `resetProcessed` so the script adds it again. |
| `Unknown time zone "…"` warning | Add the Windows time zone name to `WINDOWS_TZ` in the script, with its IANA equivalent. |
| A change made in the Partner Portal doesn't show up | The script only sees changes that arrive by email. The Partner Portal is the reliable record of your schedule. |

## Uninstall

Run `uninstallTrigger`, then delete the project at script.google.com. Events the script already added stay in your calendar.
