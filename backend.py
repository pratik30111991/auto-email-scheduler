from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import os

app = Flask(__name__)

SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# ✅ Save credentials.json from environment secret
with open(JSON_FILE, "w") as f:
    f.write(os.environ.get("GOOGLE_JSON", ""))
print("💾 credentials.json written from env")

@app.route("/")
def home():
    return "✅ Email Tracker is running"

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    print(f"📩 /track triggered → sheet={sheet_name}, row={row}")

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ])
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID)
        ws = sheet.worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")  # Column 10 = 'Open?'
        print(f"✅ Row {row} in sheet '{sheet_name}' marked as OPENED")
    except Exception as e:
        print(f"❌ ERROR: Unable to update sheet → {e}")

    # 1x1 transparent pixel
    pixel = io.BytesIO(
        b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!'
        b'\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'
    )
    return send_file(pixel, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
