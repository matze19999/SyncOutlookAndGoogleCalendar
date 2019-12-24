# SyncOutlookCalendarToGoogle
Dieses Script exportiert die Kalendereinträge von Microsoft Outlook in eine CSV Datei und importiert die Datei in den Google Kalender.

# Getestet mit
- Windows 10 2004
- Office 365
- Office 2019
- Office 2016

# Anforderungen
- Python 3 oder höher
- Powershell 3 oder höher

# Anleitung:

Downloade zuerst die credentials.json von folgender Seite: https://developers.google.com/calendar/quickstart/python

Klicke auf den Button "Enable the Google Calendar API" und dann auf "Download Client Configuration"

Verschiebe die Datei in den gleichen Ordner wie die 3 Scripte.

Starte die run.cmd

# Good to know

Wenn beim Ausführen der run.cmd sich eine Meldung von Outlook mit folgendem Text:

Ein Programm versucht, auf ihre E-Mail-Adressinformationen in Outlook zuzugreifen....

öffnet, muss nach dieser Anleitung das Programm in Trust Center dauerhaft zugelassen werden:

https://support.microsoft.com/de-at/help/3189806/a-program-is-trying-to-send-an-e-mail-message-on-your-behalf-warning-i
