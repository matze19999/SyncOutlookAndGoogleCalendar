# SyncOutlookAndGoogleCalendar
Dieses Script exportiert die Kalendereinträge von Microsoft Outlook in eine CSV Datei und importiert die Datei in den Google Kalender.

# Getestet mit
- Windows 10 1909 und 2004
- Office 2016
- Office 2019
- Office 365

# Anforderungen
- Python 3 oder höher
- Powershell 3 oder höher

# Anleitung:

Downloade die neueste Release und extrahiere den Code in einen Ordner auf deiner Festplatte

Installiere dir Python3 von hier: https://www.microsoft.com/de-de/p/python-37/9nj46sx7x90p

Downloade die credentials.json von folgender Seite: https://developers.google.com/calendar/quickstart/python

Klicke auf den Button "Enable the Google Calendar API" und dann auf "Download Client Configuration"

Verschiebe die Datei in den gleichen Ordner wie die 3 Scripte.

Starte die run.cmd

# Good to know

1.
Wenn beim Ausführen der run.cmd sich eine Meldung von Outlook mit folgendem Text:

Ein Programm versucht, auf ihre E-Mail-Adressinformationen in Outlook zuzugreifen....

öffnet, muss nach dieser Anleitung das Programm in Trust Center dauerhaft zugelassen werden:

https://support.microsoft.com/de-at/help/3189806/a-program-is-trying-to-send-an-e-mail-message-on-your-behalf-warning-i


2.
Standardmäßig wird der Outlook Kalender zum Google Kalender synchronisiert. Wenn du den Google Kalender zum Outlook Kalender synchronisieren willst, ändere in der run.cmd die Variable "mode" in Zeile 3 zu 1.
Wenn du beide nacheinander synchronisieren willst, schreibe den Wert 3 in die Variable Mode.
