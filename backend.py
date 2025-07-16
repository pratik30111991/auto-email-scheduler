from flask import Flask, request, make_response
import os
import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)
CREDS_FILE = "credentials.json"
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
IST = pytz.timezone("Asia/Kolkata")

def get_spreadsheet():
    if not os.path.exists(CREDS_FILE):
        creds_json = os.environ.get("GOOGLE_JSON", "")
        if not creds_json.strip():
            raise Exception("❌ GOOGLE_JSON is empty")
        with open(CREDS_FILE, "w") as f:
            f.write(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

@app.route("/track", methods=["GET"])
def track():
    try:
        sheet_name = request.args.get("sheet")
        row = int(request.args.get("row", "0"))
        email = request.args.get("email", "").strip().lower()
        ua = request.headers.get("User-Agent", "").lower()
        now = datetime.datetime.now(IST)

        if not sheet_name or not row or not email:
            print("[❌ Missing params]", sheet_name, row, email)
            return "", 204

        if "googleimageproxy" in ua or "googleusercontent" in ua:
            print(f"[SKIP: PROXY] {email}")
            return "", 204

        sheet = get_spreadsheet()
        ws = sheet.worksheet(sheet_name)
        sheet_email = ws.cell(row, 3).value or ""
        if sheet_email.strip().lower() != email:
            print(f"[SKIP: Email mismatch] Row {row}: {sheet_email} vs {email}")
            return "", 204

        open_status = ws.cell(row, 10).value
        if open_status and open_status.strip().lower() == "yes":
            print(f"[ALREADY OPENED] Row {row}")
            return "", 204

        sent_time_str = ws.cell(row, 9).value or ""
        try:
            sent_time = datetime.datetime.strptime(sent_time_str, "%d-%m-%Y %H:%M:%S").replace(tzinfo=IST)
            if (now - sent_time).total_seconds() < 10:
                print(f"[TOO EARLY] Delay <10s — {email}")
                return "", 204
        except:
            print(f"[WARN] Bad timestamp in row {row}: {sent_time_str}")
            return "", 204

        ws.update_cell(row, 10, "Yes")
        ws.update_cell(row, 11, now.strftime("%d-%m-%Y %H:%M:%S"))
        print(f"[✅ TRACKED] Row {row} for {email}")
    except Exception as e:
        print("[❌ Tracking error]", e)
        return make_response("Error", 500)

    return "", 204

@app.route("/")
def home():
    return "✅ Email tracking backend is live"
