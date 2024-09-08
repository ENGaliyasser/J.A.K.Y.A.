import gspread
from google.oauth2.service_account import Credentials

def init_google_sheet():
    # Define the scope
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    # Add your service account file path
    creds = Credentials.from_service_account_file('jakya-audit-6fd3f63aae96.json', scopes=scope)

    # Authorize the Client
    client = gspread.authorize(creds)

    # List all spreadsheets
    spreadsheets = client.openall()
    for sheet in spreadsheets:
        print("Available Sheets: ",sheet.title)

    url = 'https://docs.google.com/spreadsheets/d/1VfH1b097fqMxjAanHsr1E3fXzz3FK4m271cdcmUtCKI/edit?gid=0#gid=0'
    spreadsheet = client.open_by_url(url)

    # Select the first sheet
    sheet = spreadsheet.sheet1
    return sheet

