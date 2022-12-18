import openpyxl
import os.path
import tkinter as tk
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tkinter import filedialog

# Set up the tkinter file dialog
root = tk.Tk()
root.withdraw()

# Open the file dialog and get the file path
file_path = filedialog.askopenfilename()

# Read the excel file and store the data in a list
wb = openpyxl.load_workbook(file_path)
sheet = wb['Sheet1']# Read the excel file and store the data in a list
workbook = openpyxl.load_workbook(file_path)
sheet = workbook['Sheet1']

# Get the dates and names from the excel sheet
dates = []
names = []
for row in sheet.iter_rows(min_row=2):
    name = row[0].value
    date = row[1].value
    dates.append(date)
    names.append(name)

# Convert the dates to a datetime object
dates = [datetime.strftime(date, '%Y-%m-%d') for date in dates]

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())# If modifying these scopes, delete the file token.json.



# Set up the Google Calendar API
service = build('calendar', 'v3', credentials=creds, static_discovery=False)

# Create the events in the calendar
calendar_id = 'primary'
for date, name in zip(dates, names):
    manga = {
        'summary': name,
        'start': {
            'date': date
        },
        'end': {
            'date': date
        },
    }
    
    # get a list of events with freesearch matching name of mangas
    results = service.events().list(calendarId=calendar_id, q=name).execute()
    events = results.get('items', [])

    existingMangas = None
    for event in events:
        existingMangas = event['summary']

    if existingMangas is None:
        # Create the event using the events().insert() method
        created_event = service.events().insert(calendarId=calendar_id, body=manga).execute()
        print('Event created: ' + manga['summary'] + 'for date:' + date)
    else:
        print('event already exists:' + manga['summary'])

exit()
