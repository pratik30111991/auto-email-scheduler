from flask import Flask, request, send_file, make_response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pytz
import io
import os

app = Flask(__name__)
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# Write GOOGLE_JSON from Env to file at startup
if not os.path.exists(JSON_FILE):
    with open(JSON_FILE, "w") as f:
        f.write(os.environ.get("GOOGLE_JSON", ""))
    print("💾 credentials.json file created")

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    print(f"📩 /track called → sheet={sheet_name}, row={row}")

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ])
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID)
        ws = sheet.worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")  # Column J = 10
        print(f"✅ SUCCESS: Row {row} updated to 'Yes' in '{sheet_name}'")
    except Exception as e:
        print("❌ ERROR in /track:", e)

    # Return 1x1 transparent tracking GIF with proper headers
    pixel = io.BytesIO(b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04' +
                       b'\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;')
    response = make_response(send_file(pixel, mimetype='image/gif'))
    response.headers['Content-Disposition'] = 'inline; filename="track.gif"'
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route("/")
def index():
    return "✅ Email Tracker backend is live. Use /track?sheet=...&row=..."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
