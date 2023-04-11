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

def delete_events_by_summary(service, calendar_id, start_date, end_date, summary):
    events_result = service.events().list(calendarId=calendar_id, timeMin=start_date, timeMax=end_date, singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    for event in events:
        event_summary = event['summary']
        if event_summary == summary:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            print(f'Event deleted: {event_summary}')

def main():
    # Authenticate and get the calendar service
    service = get_calendar_service()

    # Prepare events from the Excel file
    start_date = df['Date '].min().strftime('%Y-%m-%dT00:00:00Z')
    end_date = df['Date '].max().strftime('%Y-%m-%dT23:59:59Z')

    for _, row in df.iterrows():
        if pd.isnull(row['Date ']) or pd.isnull(row['Time']):
            continue

        summary = row['Name ']
        delete_events_by_summary(service, 'primary', start_date, end_date, summary)

if __name__ == '__main__':
    main()
