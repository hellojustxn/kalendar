from __future__ import print_function
import priv
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from enum import Enum

from contextlib import suppress

DEBUG = False

class Fields(Enum):
    SUMMARY = 0
    LOCATION = 1
    DESCRIPTION = 2
    START_DATE = 3
    END_DATE = 4

def main():
    MAX_RESULT = 20
    sheetData = getSheetData(id = priv.TODO_SPREADSHEET_ID, range = priv.SCHEDULE_RANGE)
    # showCalendar(id = priv.AI_CAL_ID, maxResults = MAX_RESULT)
    existingEvents = getCalendarData(id = priv.AI_CAL_ID, maxResults = 20)
    now = datetime.datetime.now().isoformat() + 'Z' # 'Z' indicates UTC time
   
    
    for event in sheetData:
        eventExists = False
        pastEvent = False
        print(event)
        for existingEvent in existingEvents:
            if (existingEvent['summary'] == event[Fields.SUMMARY.value]):
                print('Existing event found: ', existingEvent['summary'])    
                eventExists = True
                
            if event[Fields.START_DATE.value] < now:
                # print("Calendar Event: ", event['summary'], " -----  Existing event found: ", newEvent['summary'])
                # print('Lets not create an event for the past: ', existingEvent['summary'])    
                pastEvent = True
            
        if not pastEvent and not eventExists:
            # print("Creating event: event[Fields.SUMMARY.value]")
            createEvent(priv.AI_CAL_ID,
                    event[Fields.SUMMARY.value],
                    event[Fields.LOCATION.value],
                    event[Fields.DESCRIPTION.value],
                    event[Fields.START_DATE.value],
                    event[Fields.END_DATE.value])


    # # TODO: Handle comparisons elsewhere, likely cause of "Calendar usage limits"
    # existinEvents = getCalendarData(id = id, maxResults = 20)
    # now = datetime.datetime.now().isoformat() + 'Z' # 'Z' indicates UTC time
    # # Check if the event exists
    # eventExists = False
    # pastEvent = False
    # for event in events:
    #     # print("Calendar Event: ", event['summary'], " -----  New event: ", newEvent['summary'])
    #     if event['summary'] == newEvent['summary']:
    #         # print("Calendar Event: ", event['summary'], " -----  Existing event found: ", newEvent['summary'])
    #         eventExists = True
    #         break
    #     if newEvent['start']['dateTime'] < now:
    #         # print("Calendar Event: ", event['summary'], " -----  Existing event found: ", newEvent['summary'])
    #         print('Lets not create an event in the past')    
    #         pastEvent = True
    #         break

    # if eventExists or pastEvent:
    #     print('Event already exists!')
    # else:
    #     print('Creating event: ',newEvent['summary'])
    #     service = build('calendar', 'v3', credentials=authorize('calendar'))

# Will only add/remove today's and future events, will not modify the past
def sync():
    pass

def authorize(service_type):
    if DEBUG: print("Authorizing for:", service_type, "...")
    # Select URI and scope file according to the service type

    if service_type == 'calendar':
        scope = ['https://www.googleapis.com/auth/calendar']
        pickle_file = 'token.pickle.calendar'
        cred_file = 'credentials-cal.json'
    else:
        scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        # Select pickle file according to the service type
        pickle_file = 'token.pickle.sheets'
        cred_file = 'credentials-sheets.json'

    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                cred_file, scope)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(pickle_file, 'wb') as token:
            pickle.dump(creds, token)
    return creds

def deleteEvent(calendarId, eventId, sendNotifications, sendUpdates):
    print("Deleting ", eventId)
    service = build('calendar', 'v3', credentials=authorize('calendar'))
    service.events().delete(calendarId = calendarId, eventId = eventId, sendNotifications = sendNotifications, sendUpdates = sendUpdates).execute()

def createEvent(id,summary, location, description, startDateTime, endDateTime):
    newEvent = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': startDateTime,
            'timeZone': 'America/Los_Angeles',
        },
        'end': {
            'dateTime': endDateTime,
            'timeZone': 'America/Los_Angeles',
        },
        # 'recurrence': ['RRULE:FREQ=DAILY;UNTIL=20200429T065959Z'],
        # 'attendees': [
        #     {'email': 'justin.gwle@gmail.com'},
        # ],
        'reminders': {
        'useDefault': True,
        # 'overrides': [
        #   {'method': 'email', 'minutes': 24 * 60},
        #   {'method': 'popup', 'minutes': 10},
        # ],
        },
    }

    service = build('calendar', 'v3', credentials=authorize('calendar'))
    # # TODO: Handle comparisons elsewhere, likely cause of "Calendar usage limits"
    # events = getCalendarData(id = id, maxResults = 20)
    # now = datetime.datetime.now().isoformat() + 'Z' # 'Z' indicates UTC time
    # # Check if the event exists
    # eventExists = False
    # pastEvent = False
    # for event in events:
    #     # print("Calendar Event: ", event['summary'], " -----  New event: ", newEvent['summary'])
    #     if event['summary'] == newEvent['summary']:
    #         # print("Calendar Event: ", event['summary'], " -----  Existing event found: ", newEvent['summary'])
    #         eventExists = True
    #         break
    #     if newEvent['start']['dateTime'] < now:
    #         # print("Calendar Event: ", event['summary'], " -----  Existing event found: ", newEvent['summary'])
    #         print('Lets not create an event in the past')    
    #         pastEvent = True
    #         break

    # if eventExists or pastEvent:
    #     print('Event already exists!')
    # else:
    #     print('Creating event: ',newEvent['summary'])
    #     service = build('calendar', 'v3', credentials=authorize('calendar'))
    createEvent = service.events().insert(calendarId=id, body=newEvent).execute()

def showCalendar(id, maxResults):
    events = getCalendarData(id = id, maxResults = maxResults)
    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))

def showSheet(id, range):
    service = build('sheets', 'v4', credentials=authorize(service_type='sheets'))
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=id, range=range).execute()
    values = result.get('values', [])
    for value in values:
        print(value)

# Returns a list of events from google calendar
def getCalendarData(id, maxResults):
    service = build('calendar', 'v3', credentials=authorize('calendar'))

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    # print('Getting the upcoming', maxResults, 'events')
    events_result = service.events().list(calendarId=id, timeMin=now,
                                        maxResults=maxResults, singleEvents=True,
                                        orderBy='startTime').execute()
    return events_result.get('items', [])

# Returns a list of rows from google sheet
def getSheetData(id, range):
    service = build('sheets', 'v4', credentials=authorize(service_type='sheets'))
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=id, range=range).execute()
    return result.get('values', [])


if __name__ == '__main__':
    main()
