# ================= backend.py =================

from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pytz
import io
import os

app = Flask(__name__)
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# ✅ Write credentials.json on startup
with open(JSON_FILE, "w") as f:
    f.write(os.environ.get("GOOGLE_JSON", ""))
print("💾 credentials.json written from environment")

@app.route("/")
def home():
    return "✅ Email Tracker is live"

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    print(f"📩 Tracking → sheet={sheet_name}, row={row}")

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ])
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID)
        ws = sheet.worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")  # Column J = 10 = Open?
        print(f"✅ Updated row {row} in '{sheet_name}' to 'Yes'")
    except Exception as e:
        print(f"❌ Error updating sheet {sheet_name}, row {row} — {e}")

    # 1x1 Transparent GIF
    pixel = io.BytesIO(b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04'
                       b'\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;')
    return send_file(pixel, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
