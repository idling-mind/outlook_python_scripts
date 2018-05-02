"""
Shows basic usage of the Google Calendar API. Creates a Google Calendar API
service object and outputs a list of the next 10 events on the user's calendar.
"""
from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import datetime
import win32com.client

# Setup the Calendar API
SCOPES = 'https://www.googleapis.com/auth/calendar'
store = file.Storage('credentials.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('calendar', 'v3', http=creds.authorize(Http()))

tomorrow = datetime.date.today() + datetime.timedelta(days=1)
dayafter = datetime.date.today() + datetime.timedelta(days=2)
# Refer to the Python quickstart on how to setup the environment:
# https://developers.google.com/calendar/quickstart/python
# Change the scope to 'https://www.googleapis.com/auth/calendar' and delete any
# stored credentials.

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

calendar = outlook.GetDefaultFolder(9)
calendar_items = calendar.Items
calendar_items.IncludeRecurrences = True
calendar_items.Sort("[Start]")
condition = "[Start] >= \"{}\" and [Start] <= \"{}\"".format(
    tomorrow.strftime("%d-%m-%Y"), dayafter.strftime("%d-%m-%Y")
    )
item = calendar_items.Find(condition)

while True:
    event = {
        'summary': '{}'.format(item.Subject),
        'location': '{}'.format(item.Location),
        'description': 'Meeting sent by {}.'.format(item.Organizer),
        'start': {
            'dateTime': '{}'.format(str(item.StartUTC.isoformat())),
            'timeZone': 'Asia/Calcutta',
        },
        'end': {
            'dateTime': '{}'.format(str(item.EndUTC.isoformat())),
            'timeZone': 'Asia/Calcutta',
        },
        'reminders': {
        'useDefault': False,
        'overrides': [
            {'method': 'email', 'minutes': 24 * 60},
            {'method': 'popup', 'minutes': 10},
        ],
        },
    }
    event = service.events().insert(calendarId='pj7h9i1gvk7s88ge9patn55c1g@group.calendar.google.com', body=event).execute()
    print('Event created: {} - {}'.format(item.Subject, event.get('htmlLink')))
    item = calendar_items.FindNext()
    if not item:
        break
