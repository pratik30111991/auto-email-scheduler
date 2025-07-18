from flask import Flask, request, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import time

app = Flask(__name__)

# Setup Google Sheets
GOOGLE_JSON = os.environ.get("GOOGLE_JSON", "")
SHEET_ID = os.environ.get("SHEET_ID", "")
if not GOOGLE_JSON or not SHEET_ID:
    raise Exception("Missing GOOGLE_JSON or SHEET_ID")

data = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
client = gspread.authorize(creds)
print("✅ Connected to Google Sheets")

@app.route("/")
def home():
    return "📡 Email Tracking Pixel Server Running"

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row_number = request.args.get("row")
    email = request.args.get("email")
    sent_ts = request.args.get("t")

    print(f"👁️ Tracking Request: sheet={sheet_name}, row={row_number}, email={email}, sent_ts={sent_ts}")

    try:
        # Avoid proxy preloads
        if sent_ts:
            if time.time() - int(sent_ts) < 10:
                print("⚠️ Skipping — request too fast (proxy prefetch).")
                return Response(status=204)

        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        headers = sheet.row_values(1)

        open_col = headers.index("Open?") + 1
        open_time_col = headers.index("Open Timestamp") + 1
        email_col = headers.index("Email ID") + 1

        row_index = int(row_number)

        row_values = sheet.row_values(row_index)
        row_email = row_values[email_col - 1].strip().lower() if email_col <= len(row_values) else ""

        if row_email != email.strip().lower():
            print(f"❌ Email mismatch. Expected: {row_email}, Got: {email}")
            return Response(status=204)

        sheet.update_cell(row_index, open_col, "Yes")
        sheet.update_cell(row_index, open_time_col, time.strftime("%d-%m-%Y %H:%M:%S"))
        print(f"✅ Updated: Row {row_index} marked open.")

    except Exception as e:
        print(f"❌ Error: {e}")

    # Return transparent pixel
    pixel = b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF!' \
            b'\xF9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02' \
            b'\x4C\x01\x00;'
    return Response(pixel, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
