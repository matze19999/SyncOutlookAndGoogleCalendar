@ECHO OFF
cd /d %~dp0

where python3
if %errorlevel% == 1 goto nopythoninstalled
if %errorlevel% == 0 goto run

:nopythoninstalled
echo "Es ist kein Python3 installiert!"
pause
exit 1

:run
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1'"
python3 gcal.py
exit 0
