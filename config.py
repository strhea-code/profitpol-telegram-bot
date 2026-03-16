import os

TOKEN = os.getenv("TOKEN")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME")

GOOGLE_CREDS_FILE = "google_creds.json"

GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON")

if GOOGLE_CREDS_JSON:
    with open(GOOGLE_CREDS_FILE, "w", encoding="utf-8") as f:
        f.write(GOOGLE_CREDS_JSON)