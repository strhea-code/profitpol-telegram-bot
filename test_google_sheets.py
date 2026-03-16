import gspread
from google.oauth2.service_account import Credentials

from config import GOOGLE_CREDS_FILE, SPREADSHEET_NAME

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    GOOGLE_CREDS_FILE,
    scopes=SCOPES
)

client = gspread.authorize(creds)

spreadsheet = client.open(SPREADSHEET_NAME)

print("Подключение успешно!")
print("Таблица открыта:", spreadsheet.title)
print("Листы в таблице:")

for sheet in spreadsheet.worksheets():
    print("-", sheet.title)