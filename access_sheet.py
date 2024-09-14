# Author: Khaled Waleed
# Date: 8 September 2024
# Description: This script initializes a connection to Google Sheets using gspread and Google OAuth2 credentials. 
#              It lists all available spreadsheets and selects the first sheet from a specified spreadsheet URL 
#              or spreadsheet name. Only one of the parameters (sheet_url or sheet_name) should be provided at a time.

import gspread
from google.oauth2.service_account import Credentials

def init_google_sheet(sheet_name = None, sheet_url = None):
    # Define the scope
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    service_account_info = {
        # Add JSON creds here
    }


    # Add your service account file path
    creds = Credentials.from_service_account_info(service_account_info, scopes=scope)

    # Authorize the Client
    client = gspread.authorize(creds)

    # Open sheet by URL or name
    if sheet_name and sheet_url:
        raise ValueError("Error in init_google_sheet function provide only the sheet name or the sheet URL")
    elif sheet_url:
        spreadsheet = client.open_by_url(sheet_url)
    elif sheet_name:
        spreadsheet = client.open(sheet_name)
    else:
        raise ValueError("Error in init_google_sheet function either sheet_name or sheet_url must be provided")

    # Select the first sheet
    sheet = spreadsheet.sheet1
    return sheet