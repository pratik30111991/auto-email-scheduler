from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import io
import os
import logging

app = Flask(__name__)

# Setup logging for Render
logging.basicConfig(level=logging.INFO)

# Setup Google Sheets credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_json = os.environ.get("GOOGLE_JSON")
if not credentials_json:
    raise ValueError("GOOGLE_JSON environment variable not set")

creds = ServiceAccountCredentials.from_json_keyfile_dict(eval(credentials_json), scope)
client = gspread.authorize(creds)

@app.route("/track")
def track_email():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return "", 204

    try:
        sheet = client.open_by_key("1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q").worksheet(sheet_name)
        row_num = int(row)

        headers = sheet.row_values(1)
        lower_headers = [h.strip().lower() for h in headers]

        try:
            open_col = lower_headers.index("open?") + 1
            open_time_col = lower_headers.index("open timestamp") + 1
            email_col = lower_headers.index("email id") + 1
        except ValueError as e:
            logging.warning(f"Header missing in sheet: {e}")
            return "", 204

        row_data = sheet.row_values(row_num)
        actual_email = row_data[email_col - 1].strip().lower() if len(row_data) >= email_col else ""

        if email_param.strip().lower() != actual_email:
            logging.info(f"Ignored: email mismatch for row {row_num} — expected {actual_email}, got {email_param}")
            return "", 204

        # Optional: Ignore Gmail proxy preloads using delay logic (5 seconds after send time)
        now = datetime.now()
        try:
            timestamp_col = lower_headers.index("timestamp") + 1
            sent_time_str = row_data[timestamp_col - 1]
            if sent_time_str:
                sent_time = datetime.strptime(sent_time_str, "%d/%m/%Y %H:%M:%S")
                if now < sent_time + timedelta(seconds=5):
                    logging.info(f"Ignored early open (likely Gmail proxy) for row {row_num}")
                    return "", 204
        except Exception as e:
            logging.warning(f"Send time check skipped: {e}")

        # If already marked open, skip re-writing
        open_status = row_data[open_col - 1].strip().lower() if len(row_data) >= open_col else ""
        if open_status == "yes":
            logging.info(f"Row {row_num} already marked as opened.")
            return "", 204

        # Update "Open?" and "Open Timestamp"
        sheet.update_cell(row_num, open_col, "Yes")
        sheet.update_cell(row_num, open_time_col, now.strftime("%d/%m/%Y %H:%M:%S"))

        logging.info(f"Marked row {row_num} as opened.")
    except Exception as e:
        logging.error(f"Error processing tracking for row {row}: {e}")

    # Return 1x1 transparent PNG
    pixel = b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00' \
            b'\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00' \
            b'\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b'
    return send_file(io.BytesIO(pixel), mimetype='image/gif')

@app.route("/")
def home():
    return "Email tracker running!"

