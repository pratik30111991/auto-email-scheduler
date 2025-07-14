from flask import Flask, request, send_file, make_response
import gspread
import io
from oauth2client.service_account import ServiceAccountCredentials
import os
import logging

app = Flask(__name__)

SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def send_pixel():
    pixel = io.BytesIO(
        b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04'
        b'\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'
    )
    response = make_response(send_file(pixel, mimetype='image/gif'))
    response.headers['Content-Disposition'] = 'inline; filename="track.gif"'
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    user_agent = request.headers.get("User-Agent", "").lower()

    logger.info(f"\U0001F4E9 Tracking pixel hit → sheet={sheet_name}, row={row}, UA={user_agent}")

    # Avoid bots triggering open tracking
    bot_keywords = ["bot", "crawler", "curl", "wget", "python", "uptime"]
    # Don't skip Gmail proxies (they include 'google' in UA), allow open tracking
    if any(k in user_agent for k in bot_keywords):
        logger.warning(f"🤖 Ignored bot hit on tracking pixel. UA: {user_agent}")
        return send_pixel()

    try:
        with open(JSON_FILE, "w") as f:
            f.write(os.environ["GOOGLE_JSON"])

        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ])
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID)
        ws = sheet.worksheet(sheet_name)
        ws.update_cell(int(row), 10, "Yes")  # Column 10 = 'Open?'
        logger.info(f"✅ SUCCESS: Updated Open? to 'Yes' in sheet '{sheet_name}', row {row}")
    except Exception as e:
        logger.error(f"❌ ERROR updating Open? in sheet '{sheet_name}', row {row}: {e}")

    return send_pixel()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
