import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = '1kAtn-w62aKq7lS1r_mmW3aZTcOlnrN1vFec5HXq25rU'

def main():
    credentials = None
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r'C:\Users\joaop\Área de Trabalho\Controle Oficios\Credentials\credentials.json', SCOPES)
            credentials = flow.run_local_server(port = 0)
        with open('token.json','w') as token:
            token.write(credentials.to_json())

    try:
        service = build('sheets', 'v4', credentials=credentials)
        sheets = service.spreadsheets()

        result = sheets.values().get(spreadsheetId= SPREADSHEET_ID,range ='Página1!A1:C6').execute()

        values = result.get('values', [])

        for row in values:
            print(values)
    except HttpError as error:
        print(error)

if __name__ == '__main__':
    main()