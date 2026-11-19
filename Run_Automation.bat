@echo off
cd /d "%~dp0"

echo ======================================================================
echo                OMEGA Watch Image Collector
echo ======================================================================
echo.

echo Launching Microsoft Edge in Remote Debugging mode...
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebug"
) else if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebug"
) else (
    start "" msedge --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebug"
)

echo.
echo Running Automation.py...
echo.
python Automation.py

echo.
echo ======================================================================
echo Process finished. Press any key to close this window.
echo ======================================================================
pause
