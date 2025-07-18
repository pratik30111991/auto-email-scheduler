from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
import os
from datetime import datetime
import pytz

app = Flask(__name__)

# Authenticate using GOOGLE_JSON environment variable
GOOGLE_JSON = os.environ.get("GOOGLE_JSON")

if not GOOGLE_JSON:
    raise Exception("GOOGLE_JSON not found in environment variables!")

creds_dict = json.loads(GOOGLE_JSON)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# Route to check if server is running
@app.route("/live", methods=["GET"])
def live_check():
    return "Tracker running", 200

# Main pixel tracking endpoint
@app.route("/track", methods=["GET"])
def track_email_open():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return "Missing parameters", 400

    try:
        sheet = client.open_by_key(os.environ.get("SHEET_ID")).worksheet(sheet_name)
        email_cell = sheet.acell(f"C{row}").value  # Column C = Email ID
        if email_cell.strip().lower() != email_param.strip().lower():
            return "Email mismatch", 403

        open_status = sheet.acell(f"I{row}").value  # Column I = Open?
        if open_status != "Yes":
            sheet.update_acell(f"I{row}", "Yes")
            ist_now = datetime.now(pytz.timezone("Asia/Kolkata"))
            sheet.update_acell(f"J{row}", ist_now.strftime("%d/%m/%Y %H:%M:%S"))

    except Exception as e:
        return str(e), 500

    # Return 1x1 transparent pixel
    pixel = b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00' \
            b'\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00' \
            b'\x01\x00\x01\x00\x00\x02\x02\x4c\x01\x00\x3b'
    return send_file(io.BytesIO(pixel), mimetype='image/gif')

# Optional root route (will return 404 if not defined)
@app.route("/", methods=["GET"])
def root():
    return "Tracking backend is live. Use /track?sheet=...&row=...&email=...", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
