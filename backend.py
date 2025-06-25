# ================= backend.py =================
from flask import Flask, request, send_file, make_response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io, os

app = Flask(__name__)
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# write credentials on startup
with open(JSON_FILE, "w") as f:
    f.write(os.environ.get("GOOGLE_JSON", ""))

@app.route("/")
def home():
    return "✅ Email Tracker is live"

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    print(f"📩 Tracking → sheet={sheet_name}, row={row}")

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            JSON_FILE,
            ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive",
             "https://www.googleapis.com/auth/spreadsheets"]
        )
        client = gspread.authorize(creds)
        ws = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")          # column J
        print(f"✅ Updated '{sheet_name}' row {row}")
    except Exception as e:
        print(f"❌ Sheet update failed – {e}")

    # 1×1 transparent GIF + no-cache headers
    pixel = io.BytesIO(
        b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!'
        b'\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01'
        b'\x00\x00\x02\x02D\x01\x00;'
    )
    resp = make_response(send_file(pixel, mimetype="image/gif"))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
