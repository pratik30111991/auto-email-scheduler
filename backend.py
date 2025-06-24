# === ✅ backend.py ===

from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pytz
import io
import os
import json

app = Flask(__name__)
INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# Write the credentials file at runtime
if not os.path.exists(JSON_FILE):
    with open(JSON_FILE, "w") as f:
        f.write(os.environ.get("GOOGLE_JSON", ""))

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ])
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID)
        ws = sheet.worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")
        print(f"✅ Open tracked for row {row} in {sheet_name}")
    except Exception as e:
        print("❌ Track error:", e)

    pixel = io.BytesIO(b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04' +
                       b'\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;')
    return send_file(pixel, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
