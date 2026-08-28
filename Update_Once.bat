@echo off
REM Double-click this file to update the dashboard one time right now,
REM without starting the continuous live watcher.

cd /d "%~dp0"

echo Updating dashboard...
python run_dashboard_update.py

echo.
echo Done. Check dashboard_update.log if something looks wrong.
pause
