from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pytz
import io
import os

app = Flask(__name__)
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# 💾 Write GOOGLE_JSON from env to credentials file (Render will set this from secret)
if not os.path.exists(JSON_FILE):
    with open(JSON_FILE, "w") as f:
        f.write(os.environ.get("GOOGLE_JSON", ""))
    print("✅ credentials.json created")

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
        ws.update_cell(int(row), 10, "Yes")  # 'Open?' is column J (10)
        print(f"✅ Row {row} updated to 'Yes' in sheet '{sheet_name}'")
    except Exception as e:
        print("❌ ERROR in /track route:", e)

    # ✅ Return an actual invisible 1x1 transparent pixel (GIF)
    pixel = (
        b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00'
        b'\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00'
        b'\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b'
    )

    return send_file(
        io.BytesIO(pixel),
        mimetype='image/gif',
        as_attachment=False,
        download_name="pixel.gif"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
