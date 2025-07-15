from flask import Flask, request, send_file, make_response
import gspread
import io
from oauth2client.service_account import ServiceAccountCredentials
import os
import logging
from datetime import datetime, timedelta
import pytz

app = Flask(__name__)
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"
EMAIL_COL_INDEX = 3
TIMESTAMP_COL_INDEX = 9  # 'Timestamp'
OPEN_COL_INDEX = 10      # 'Open?'

INDIA_TZ = pytz.timezone("Asia/Kolkata")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def send_pixel():
    pixel = io.BytesIO(
        b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!'
        b'\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'
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
    email_param = request.args.get("email", "").strip().lower()
    user_agent = request.headers.get("User-Agent", "").lower()

    logger.info(f"📩 Tracking pixel hit → sheet={sheet_name}, row={row}, email={email_param}, UA={user_agent}")

    # Basic bot filtering
    if any(k in user_agent for k in ["bot", "crawler", "curl", "wget", "python", "uptime"]):
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

        actual_email = ws.cell(int(row), EMAIL_COL_INDEX).value.strip().lower()
        if actual_email != email_param:
            logger.warning(f"⛔ Email mismatch: sheet={actual_email}, pixel={email_param} — skipping Open? update")
            return send_pixel()

        # Check sent timestamp to enforce minimum delay
        timestamp_str = ws.cell(int(row), TIMESTAMP_COL_INDEX).value
        if not timestamp_str:
            logger.warning("⏱️ No timestamp found — skipping Open? update")
            return send_pixel()

        try:
            sent_time = datetime.strptime(timestamp_str, "%d-%m-%Y %H:%M:%S")
            sent_time = INDIA_TZ.localize(sent_time)
        except Exception as e:
            logger.warning(f"⏱️ Timestamp parse failed: {e} — skipping")
            return send_pixel()

        now = datetime.now(INDIA_TZ)
        if (now - sent_time) < timedelta(seconds=15):
            logger.warning("⏱️ TOO EARLY — possible Gmail preload — skipping")
            return send_pixel()

        ws.update_cell(int(row), OPEN_COL_INDEX, "Yes")
        logger.info(f"✅ SUCCESS: Open? marked 'Yes' in sheet '{sheet_name}', row {row}")
    except Exception as e:
        logger.error(f"❌ ERROR updating Open? in sheet '{sheet_name}', row {row}: {e}")

    return send_pixel()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
