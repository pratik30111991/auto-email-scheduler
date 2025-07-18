from flask import Flask, request, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import os
import pytz
import logging

app = Flask(__name__)

# Setup logging
logging.basicConfig(level=logging.INFO)

# Use IST (India Standard Time)
IST = pytz.timezone('Asia/Kolkata')

# Google Sheets setup
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
GOOGLE_JSON = os.environ.get("GOOGLE_JSON")

if not GOOGLE_JSON:
    raise Exception("GOOGLE_JSON env variable missing")

with open("credentials.json", "w") as f:
    f.write(GOOGLE_JSON)

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

@app.route('/track', methods=['GET'])
def track_email_open():
    sheet_name = request.args.get('sheet')
    row = request.args.get('row')
    email_param = request.args.get('email')

    if not sheet_name or not row or not email_param:
        return "Missing parameters", 400

    try:
        sheet = client.open_by_key("1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q").worksheet(sheet_name)
        row = int(row)

        # Read current values
        open_status = sheet.cell(row, 10).value  # "Open?" column
        sheet_email = sheet.cell(row, 3).value   # Email ID column
        open_timestamp = sheet.cell(row, 11).value  # "Open Timestamp"

        # Validate email
        if sheet_email.strip().lower() != email_param.strip().lower():
            logging.warning(f"Email mismatch on row {row}: Sheet={sheet_email}, Param={email_param}")
            return Response(status=204)

        # Skip if already opened
        if open_status == "Yes" and open_timestamp:
            logging.info(f"Already marked as opened for row {row}")
            return Response(status=204)

        # Get current IST time
        now = datetime.datetime.now(IST)
        open_time_str = now.strftime("%d/%m/%Y %H:%M:%S")

        # Update "Open?" and "Open Timestamp"
        sheet.update_cell(row, 10, "Yes")
        sheet.update_cell(row, 11, open_time_str)

        logging.info(f"Marked as opened: row={row}, email={email_param}, time={open_time_str}")
        return Response(status=204)

    except Exception as e:
        logging.error(f"Error processing pixel for row={row}, email={email_param} → {str(e)}")
        return "Error", 500

if __name__ == '__main__':
    app.run()
