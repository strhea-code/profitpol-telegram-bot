import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# путь к json ключу
CREDS_FILE = "google_creds.json"

creds = Credentials.from_service_account_file(
    CREDS_FILE,
    scopes=SCOPES
)

client = gspread.authorize(creds)

# вывести список всех таблиц, к которым есть доступ
files = client.list_spreadsheet_files()

print("Таблицы, доступные боту:\n")

for file in files:
    print(file["name"])