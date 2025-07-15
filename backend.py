import os
import io
from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz

app = Flask(__name__)
INDIA_TZ = pytz.timezone("Asia/Kolkata")

# Load Google Sheets creds from env
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
if not GOOGLE_JSON:
    raise Exception("❌ GOOGLE_JSON not found in environment.")

creds_dict = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
gc = gspread.authorize(credentials)

@app.route("/track")
def track():
    try:
        sheet_name = request.args.get("sheet")
        row = int(request.args.get("row"))
        email = request.args.get("email", "").strip().lower()

        if not sheet_name or not row or not email:
            return "", 204

        # Validate real browser (skip proxy preloads)
        ua = request.headers.get("User-Agent", "")
        if "GoogleImageProxy" in ua or "ggpht" in ua:
            print("⚠️ Proxy hit ignored.")
            return "", 204

        # Update only if all parameters are present
        spreadsheet = gc.open_by_key(creds_dict["spreadsheet_id"])
        worksheet = spreadsheet.worksheet(sheet_name)
        headers = worksheet.row_values(1)

        col_open = headers.index("Open?") + 1
        col_open_time = headers.index("Open Timestamp") + 1
        col_email = headers.index("Email ID") + 1

        existing_email = worksheet.cell(row, col_email).value.strip().lower()
        if existing_email != email:
            print("⚠️ Email mismatch in row.")
            return "", 204

        # Update open status and timestamp
        worksheet.update_cell(row, col_open, "Yes")
        worksheet.update_cell(row, col_open_time, datetime.now(INDIA_TZ).strftime("%d-%m-%Y %H:%M:%S"))
        print(f"✅ Row {row}: Open tracked.")
        return "", 204

    except Exception as e:
        print(f"[❌ ERROR] Tracking error: {str(e)}")
        return "Internal Server Error", 500

@app.route("/")
def home():
    return "✅ Email tracker live!", 200
