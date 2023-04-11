import pandas as pd
import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build

# Replace the file path with the actual path to your Excel file
excel_file = 'eye2022.xlsx'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file)

# Print the DataFrame
print(df)

# Google Calendar API settings
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Function to authenticate and get the Google Calendar API client
def get_calendar_service():
    creds = None
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        import os

        client_secret_path = 'client_secret.json'
        flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)

        creds = flow.run_local_server(port=0)
    return build('calendar', 'v3', credentials=creds)

# Add events to Google Calendar
def add_events_to_calendar(service, events):
    calendar_id = 'primary'
    for event in events:
        try:
            event = service.events().insert(calendarId=calendar_id, body=event).execute()
            print(f'Event created: {event.get("htmlLink")}')
        except HttpError as error:
            print(f'An error occurred: {error}')
            event = None
    return event

def main():
    # Authenticate and get the calendar service
    service = get_calendar_service()

    # Prepare events from the Excel file
    events = []
    for _, row in df.iterrows():
        if pd.isnull(row['Date ']) or pd.isnull(row['Time']):
            continue

        time_obj = datetime.datetime.strptime(str(row['Time']), '%H:%M:%S').time()
        start_time = datetime.datetime.combine(row['Date '], time_obj)
        end_time = start_time + datetime.timedelta(hours=1)  # Adjust the duration as needed

        event = {
            'summary': row['Name '],
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': 'America/New_York'  # Change the timezone as needed
            },
            'end': {
                'dateTime': end_time.isoformat(),
                'timeZone': 'America/New_York'  # Change the timezone as needed
            },
        }
        events.append(event)

    # Add events to Google Calendar
    add_events_to_calendar(service, events)

if __name__ == '__main__':
    main()