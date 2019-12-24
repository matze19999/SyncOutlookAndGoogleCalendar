#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Downloade zuerst die credentials.json von folgender Seite: https://developers.google.com/calendar/quickstart/python
# Klicke auf den Button "Enable the Google Calendar API" und dann auf "Download Client Configuration"
# Verschiebe die Datei in den gleichen Ordner wie die 3 Scripte.

# Wenn die Einträge nicht im Google Standard Kalender erstellt werden sollen, muss überall der Wert von calendarId=primary mit der ID des gewünschten Kalenders ersetzt werden. Siehe Ausgabe.


from __future__ import print_function
try:
    import os
    import datetime
    import pickle
    import requests
    import os.path
    import time
    import webbrowser
    import csv, pytz
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
except:
    print("Fehlende Pakete werden installiert...\n\n\n")
    os.system("pip3 install google-api-python-client google-auth-httplib2 google-auth-oauthlib datetime requests python-csv pytz")
    print("\n\nBitte Script neustarten!")
    exit()


try:
    import httplib
except:
    import http.client as httplib

# Teste Internet Verbindung
def have_internet():
    conn = httplib.HTTPConnection("www.google.com", timeout=2)
    try:
        conn.request("HEAD", "/")
        conn.close()
        return True
    except:
        conn.close()
        print("Keine Internet Verbindung!")
        exit()
        return False

have_internet()
 

SCOPES = ['https://www.googleapis.com/auth/calendar.events','https://www.googleapis.com/auth/calendar']   

def main():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                print("\nDu hast vergessen die Datei credentials.json herunterzuladen!\n\nBitte auf folgender Website den Button 'Enable the Google Calender API' drücken.")
                time.sleep(3)
                url = "https://developers.google.com/calendar/quickstart/python"
                webbrowser.open(url)
                exit(1)
        # Speichere Zugangsdaten in Datei
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            cal_summary = calendar_list_entry['summary']
            cal_id = calendar_list_entry['id']
            print("Kalender " +  cal_summary + " hat folgende ID: " + cal_id, end='\n')
            page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    now = datetime.datetime.now(pytz.timezone('Europe/Berlin')).isoformat()
    print("Schreibe die nächsten 10 Termine im Google Kalender...")
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

    # Importiere CSV
    try:
        with open('events.csv', encoding='utf-16') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for row in readCSV:
                Subject = row[0]
                StartDate = row[1] + "_" + row[2]
                EndDate = row[3] + "_" + row[4]
                AllDayEvent = row[5]
                Description = row[6]
                Location = row[7]

                StartDate_n = datetime.datetime.strptime(StartDate, '%m/%d/%Y_%H:%M:%S')
                StartDate_n = StartDate_n.isoformat()

                # Rufe die Google Kalender API auf
                print('Suche nach Termin...')
                events_result = service.events().list(calendarId='primary', timeMin=StartDate_n+"+01:00",
                                                    maxResults=10, singleEvents=True,
                                                    orderBy='startTime', q=Subject).execute()
                events = events_result.get('items', [])
                for event in events:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    print(start, event['summary'])

                if not events:

                    EndDate_n = datetime.datetime.strptime(EndDate, '%m/%d/%Y_%H:%M:%S')
                    EndDate_n = EndDate_n.isoformat()

                    event = {
                    'summary': Subject,
                    'location': Location,
                    'description': Description,
                    'start': {
                        'dateTime': StartDate_n,
                        'timeZone': 'Europe/Berlin',
                    },
                    'end': {
                        'dateTime': EndDate_n,
                        'timeZone': 'Europe/Berlin',
                #    },
                #    'recurrence': [
                #        'RRULE:FREQ=DAILY;COUNT=2'
                #    ],
                #    'attendees': [
                #        {'email': 'lpage@example.com'},
                #        {'email': 'sbrin@example.com'},
                #    ],
                #    'reminders': {
                #        'useDefault': False,
                #        'overrides': [
                #        {'method': 'email', 'minutes': 24 * 60},
                #        {'method': 'popup', 'minutes': 10},
                #        ],
                    },
                    }

                    event = service.events().insert(calendarId='primary', body=event).execute()
                    print("Termin " + Subject + " erstellt")
                for event in events:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    print("Termin gefunden: " + start, event['summary'])

    except FileNotFoundError:
        print("\nKeine CSV Datei gefunden! Evtl. wurden in Outlook keine Termine gefunden oder das PowerShell Script wurde nicht gestartet!\n")
        exit(1)

    os.remove("events.csv")
    
if __name__ == '__main__':
    main()
