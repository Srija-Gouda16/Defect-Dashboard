# Defect Log Dashboard Pipeline

A locally-running Python pipeline that reads defect log data from OneDrive
Excel files and generates a live, interactive HTML dashboard - no server,
no cloud hosting, just a script that runs on a schedule and writes a file
you can open in any browser.

## What's here

- `build_data.py` - reads the DEFECT LOG sheets, cleans/normalizes the
  data, computes all stats (station breakdowns, trends, Pareto charts,
  scrap cost, tester yield, etc.)
- `dashboard_template.html` - the dashboard itself (HTML/CSS/JS with
  Chart.js), gets fresh data injected into it on every run
- `run_dashboard_update.py` - main entry point, ties everything together
  and writes the final dashboard HTML
- `watch_and_update.py` - continuous watcher that rebuilds automatically
  on file changes (alternative to running on a timer)
- `parse_production_report.py` - reads HPR/DPR/WPR production reports for
  tester yield (FPY/FY) calculations
- `config.py` - **all your local settings** (file paths, line names, cost
  file location) - this is the file you'll edit most often
- `Update_Once.bat` / `Start_Live_Dashboard.bat` - double-click launchers,
  no command line needed
- `SETUP.md` - full setup walkthrough

## Quick start

See `SETUP.md` for the full walkthrough. Short version:
```
pip install -r requirements.txt
python run_dashboard_update.py
```

## Making changes

1. Pull the latest first: `git pull`
2. Edit whatever file needs changing
3. Test it locally (`python run_dashboard_update.py`, check the output)
4. Commit and push:
   ```
   git add .
   git commit -m "describe what changed"
   git push
   ```

You can also edit simple files (like `config.py`) directly on GitHub's
website - click the file, click the pencil/edit icon, make your change,
commit. Just remember to `git pull` on this PC afterward so the running
copy actually picks up the change - editing on GitHub alone doesn't
update what's running locally.
