from flask import Flask, request, send_file
from datetime import datetime
import pytz
import gspread
import os
import io
import base64
import json
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Timezone setup
IST = pytz.timezone('Asia/Kolkata')

# Authorize with Google Sheets using GOOGLE_JSON from environment
google_json = os.environ.get("GOOGLE_JSON")
if not google_json:
    raise Exception("❌ GOOGLE_JSON environment variable not set.")

service_account_info = json.loads(google_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

# Your actual Sheet ID
sheet_id = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
sh = client.open_by_key(sheet_id)

@app.route("/track")
def track():
    sheet = request.args.get("sheet")
    row_str = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet or not row_str or not email_param:
        return '', 204

    try:
        worksheet = sh.worksheet(sheet)
        row = int(row_str)
        row_data = worksheet.row_values(row)

        # Check: column 3 = Email ID
        sheet_email = row_data[2].strip()
        if sheet_email != email_param:
            print(f"❌ Email mismatch: expected {sheet_email}, got {email_param}")
            return '', 204

        # Only update if not already opened
        open_col = 10  # J column = Open?
        timestamp_col = 11  # K column = Open Timestamp
        open_status = row_data[open_col - 1] if len(row_data) >= open_col else ""

        if open_status != "Yes":
            worksheet.update_acell(f'J{row}', "Yes")
            open_time = datetime.now(IST).strftime("%d/%m/%Y %H:%M:%S")
            worksheet.update_acell(f'K{row}', open_time)
            print(f"✅ Marked row {row} as opened at {open_time}")
        else:
            print(f"ℹ️ Row {row} already marked as opened.")

    except Exception as e:
        print(f"❌ Error in /track: {e}")

    # Send 1x1 transparent GIF
    gif_base64 = "R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=="
    gif_data = base64.b64decode(gif_base64)
    return send_file(io.BytesIO(gif_data), mimetype='image/gif')
