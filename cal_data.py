from __future__ import print_function
import datetime
import dateutil.parser
import pickle
import os.path
import csv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Retrieve calendars (kind, summary = title, id = calendarId, backgroundColor)
    calendarList = service.calendarList().list().execute()["items"]

    # Call the Calendar API
    start = datetime.datetime.utcfromtimestamp
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    weekAgo = (datetime.datetime.now() - datetime.timedelta(days=7)).isoformat() + 'Z'
    print('Getting last week data...')
    with open('calendar_data.csv', mode='w+') as calendar_data:
        calendar_writer = csv.writer(calendar_data, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        calendar_writer.writerow(["Calendar", "CalendarColor", "Start", "End", "Duration", "Event"])
        for calendar in calendarList:
            events_result = service.events().list(
                calendarId=calendar["id"],
                timeMin=weekAgo,
                timeMax=now,
                maxResults=None, 
                singleEvents=True,
                orderBy='startTime'
                ).execute()
            events = events_result.get('items', [])

            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                start = dateutil.parser.parse(start)
                try:
                    end = event['end'].get('dateTime', event['end'].get('date'))
                    end = dateutil.parser.parse(end)
                    duration = end - start
                    calendar_writer.writerow([calendar["summary"], calendar["backgroundColor"], start, end, duration, event["summary"].encode('ascii', 'ignore')])
                except:
                    calendar_writer.writerow([calendar["summary"], calendar["backgroundColor"], start, 0, 0, event["summary"].encode('ascii', 'ignore')])

if __name__ == '__main__':
    main()