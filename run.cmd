@ECHO OFF

set mode=1


cd /d %~dp0

del events.csv events2.csv

where python3
if %errorlevel% == 1 goto nopythoninstalled
if %errorlevel% == 0 goto run

:nopythoninstalled
echo Es ist kein Python3 installiert, bitte hier downloaden:
start https://www.microsoft.com/de-de/p/python-37/9nj46sx7x90p
pause
exit 1

:run
if %mode% == 1 goto outlooktogoogle
if %mode% == 2 goto googletooutlook

:outlooktogoogle
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1' -mode 1"
python3 "gcal.py" "1"
exit 0

:googletooutlook
python3 "gcal.py" "2"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1' -mode 2"

exit 0

:both
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1' -mode 1"
python3 "gcal.py" "1"
python3 "gcal.py" "2"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1' -mode 2"
exit 0
