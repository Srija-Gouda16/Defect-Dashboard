# Setup guide: hourly-refreshing dashboard (current setup)

Reads directly from your real, live OneDrive folder - never modifies the
source files, only reads them. No Azure, no Power Automate, no login
flow. Refreshes once an hour via Windows Task Scheduler.

## 1. Confirm config.py

Already set to:
- `DATA_SOURCE = "local"`
- `ONEDRIVE_FOLDER` - your real source folder
- `FILES` - all 8 lines
- `OUTPUT_PATH` - where the dashboard HTML gets saved

Just double check the folder path and filenames match what's actually in
File Explorer.

## 2. Install dependencies

```
pip install -r requirements.txt
```

## 3. Test a single manual run

Double-click **`Update_Once.bat`**. Check `dashboard_update.log` for
"Dashboard updated successfully", then open the dashboard HTML at your
`OUTPUT_PATH` and confirm it looks right.

## 4. Schedule it to run every hour (Task Scheduler)

1. Open **Task Scheduler** → **Create Task**
2. General tab: name it "Defect Dashboard Hourly Update"
3. Triggers tab → New:
   - On a schedule → Daily → Recur every 1 days
   - Check "Repeat task every" → 1 hour → for a duration of Indefinitely
4. Actions tab → New:
   - Start a program → Browse to `Update_Once.bat` in this folder
5. Conditions tab: uncheck "only if on AC power" if this runs on a laptop
6. OK to save

To test immediately instead of waiting an hour: right-click the task →
**Run**.

## What "hourly" actually means here

Every run rebuilds everything in one pass - Hourly, Daily, and Weekly
views all refresh together, since they all come from reading the same 8
files. There's no separate "watch for changes" step anymore - just a
plain timer that reads whatever's currently in the files once every hour.

## Troubleshooting

- Check `dashboard_update.log` in this folder for the full run history
  and any errors
- If a specific file won't parse (wrong date format, missing sheet, etc.),
  the error message will name the exact issue - screenshot a sample row
  from that file if you need help fixing it
- The PC needs to be on (not asleep) for the scheduled task to actually
  fire each hour
