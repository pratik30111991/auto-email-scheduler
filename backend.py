from flask import Flask, request, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import os
import pytz
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

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

        email_from_sheet = sheet.cell(row, 3).value
        open_status = sheet.cell(row, 10).value
        open_timestamp = sheet.cell(row, 11).value

        # Check email matches
        if email_from_sheet.strip().lower() != email_param.strip().lower():
            logging.warning(f"Email mismatch: Sheet={email_from_sheet}, Param={email_param}")
            return Response(status=204)

        # Skip if already opened
        if open_status == "Yes" and open_timestamp:
            logging.info(f"Row {row} already marked opened.")
            return Response(status=204)

        # Update timestamp with current IST
        now = datetime.datetime.now(IST)
        open_time_str = now.strftime("%d/%m/%Y %H:%M:%S")

        # Update Google Sheet
        sheet.update_cell(row, 10, "Yes")               # Open?
        sheet.update_cell(row, 11, open_time_str)       # Open Timestamp

        logging.info(f"Marked row {row} as opened at {open_time_str}")
        return Response(status=200)

    except Exception as e:
        logging.error(f"Error in /track → {str(e)}")
        return "Error", 500

if __name__ == '__main__':
    app.run()
