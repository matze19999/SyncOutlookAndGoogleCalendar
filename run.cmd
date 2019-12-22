cd /d %~dp0
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& './Get-OLCalendar.ps1'"
python3 gcal.py
exit 0