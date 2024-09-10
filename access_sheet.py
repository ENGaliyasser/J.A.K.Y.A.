# Author: Khaled Waleed
# Date: 8 September 2024
# Description: This script initializes a connection to Google Sheets using gspread and Google OAuth2 credentials. 
#              It lists all available spreadsheets and selects the first sheet from a specified spreadsheet URL 
#              (this can be changed to select the sheet by its name using:  spreadsheet = client.open(spread_sheet_name)

import gspread
from google.oauth2.service_account import Credentials

def init_google_sheet():
    # Define the scope
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    # add JSON HERE
    # Add your service account file path
    creds = Credentials.from_service_account_info(service_account_info, scopes=scope)

    # Authorize the Client
    client = gspread.authorize(creds)

    # List all spreadsheets
    spreadsheets = client.openall()
    for sheet in spreadsheets:
        print("Available Sheets: ", sheet.title)
    
    # Open sheet by URL
    url = 'https://docs.google.com/spreadsheets/d/1VfH1b097fqMxjAanHsr1E3fXzz3FK4m271cdcmUtCKI/edit?gid=0#gid=0'
    spreadsheet = client.open_by_url(url)

    # Open by name
    # spreadsheet = client.open("Test")

    # Select the first sheet
    sheet = spreadsheet.sheet1
    return sheet

