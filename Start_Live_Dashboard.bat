@echo off
REM Double-click this file to start the live dashboard watcher.
REM It will keep running and restart itself automatically if it ever
REM crashes or closes unexpectedly - just leave this window open.
REM To stop it, close this window or press Ctrl+C.

cd /d "%~dp0"

:restart_loop
echo.
echo ============================================
echo  Starting Defect Dashboard Live Watcher
echo  (This window must stay open)
echo ============================================
echo.

python watch_and_update.py

echo.
echo ============================================
echo  Watcher stopped or crashed. Restarting in 10 seconds...
echo  (Close this window to stop permanently)
echo ============================================
timeout /t 10
goto restart_loop
