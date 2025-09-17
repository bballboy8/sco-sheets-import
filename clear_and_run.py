import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

SHEET_ID = os.environ["SHEET_ID"]
GOOGLE_CREDENTIALS = os.environ["GOOGLE_CREDENTIALS"]

def gs_client():
    info = json.loads(GOOGLE_CREDENTIALS)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

def clear_sheet():
    print("Clearing existing sheet data...")
    gc = gs_client()
    sh = gc.open_by_key(SHEET_ID)
    
    try:
        ws = sh.worksheet("Records")
        # Clear all data
        ws.clear()
        print("✅ Sheet cleared successfully!")
    except gspread.WorksheetNotFound:
        print("Records worksheet not found - will be created on first run")

if __name__ == "__main__":
    clear_sheet()
