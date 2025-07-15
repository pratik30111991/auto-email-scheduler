from flask import Flask, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import os
from gspread.utils import rowcol_to_a1

app = Flask(__name__)

INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

if not os.path.exists(JSON_FILE):
    with open(JSON_FILE, "w") as f:
        f.write(os.environ["GOOGLE_JSON"])

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)

@app.route("/track")
def track():
    try:
        sheet_name = request.args.get("sheet", "").strip()
        row = int(request.args.get("row", "0"))
        email_param = request.args.get("email", "").strip().lower()
        ua = request.headers.get("User-Agent", "").lower()

        if not sheet_name or not row or not email_param:
            return "", 400

        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
        values = sheet.row_values(row)

        if len(values) < 9:
            print(f"⚠️ Row {row} too short — skipping.")
            return "", 204

        stored_email = values[1].strip().lower()  # Col B
        stored_status = values[7].strip().lower()  # Col H
        stored_timestamp = values[8].strip()       # Col I

        print(f"📩 Tracking pixel hit → sheet={sheet_name}, row={row}, email={stored_email}, UA={ua}")

        if stored_email != email_param:
            print(f"⚠️ Email mismatch: {email_param} != {stored_email} — skipping.")
            return "", 204

        try:
            sent_time = INDIA_TZ.localize(datetime.strptime(stored_timestamp, "%d-%m-%Y %H:%M:%S"))
            now = datetime.now(INDIA_TZ)
            delta = (now - sent_time).total_seconds()
            if delta < 5:
                print(f"⏳ Too soon after sent time ({delta:.2f}s) — skipping.")
                return "", 204
        except Exception as e:
            print(f"⚠️ Invalid timestamp: {stored_timestamp} — {e}")
            return "", 204

        # ✅ Update "Open?" column (Col J = 10) and "Open Timestamp" (Col K = 11)
        sheet.update_acell(rowcol_to_a1(row, 10), "Yes")
        sheet.update_acell(rowcol_to_a1(row, 11), now.strftime("%d-%m-%Y %H:%M:%S"))
        print(f"✅ SUCCESS: Open? = 'Yes' and Open Timestamp updated in row {row}")
        return "", 200

    except Exception as e:
        print("❌ ERROR in /track:", e)
        return "", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
