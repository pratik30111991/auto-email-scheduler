from flask import Flask, request, send_file
import datetime
import pytz
import os
import gspread
import io
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Authenticate with Google Sheets
def get_sheet(sheet_name):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_json = os.getenv('GOOGLE_JSON')

    if not creds_json:
        raise Exception("GOOGLE_JSON environment variable not set")

    creds_dict = eval(creds_json)  # securely convert string to dict
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    # Use your Sheet ID here (hardcoded or from env)
    SHEET_ID = os.getenv("SHEET_ID")  # Set this in Render
    sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
    return sheet

@app.route('/track')
def track_email():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return "Missing required parameters", 400

    try:
        row = int(row)
        sheet = get_sheet(sheet_name)

        # Get all values of the row
        row_values = sheet.row_values(row)
        open_col_index = row_values.index("Yes") + 1 if "Yes" in row_values else None

        # Always use fixed column indices
        OPEN_COL = 9      # "Open?" (I column)
        OPEN_TIME_COL = 10  # "Open Timestamp" (J column)
        EMAIL_COL = 3     # "Email ID" (C column)

        sheet_email = sheet.cell(row, EMAIL_COL).value.strip().lower()

        if sheet_email != email_param.strip().lower():
            return "", 204  # email mismatch (don't mark open)

        open_status = sheet.cell(row, OPEN_COL).value

        if open_status != "Yes":
            # Get current time in IST
            ist = pytz.timezone('Asia/Kolkata')
            current_time = datetime.datetime.now(ist).strftime('%d-%m-%Y %H:%M:%S')

            # Update "Open?" and "Open Timestamp"
            sheet.update_cell(row, OPEN_COL, "Yes")
            sheet.update_cell(row, OPEN_TIME_COL, current_time)

    except Exception as e:
        print(f"Error: {e}")
        return "Internal Error", 500

    # Return a transparent 1x1 pixel
    pixel = b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\xff\x00\xc0\xc0\xc0\x00\x00\x00\x21\xf9\x04' \
            b'\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3b'
    return send_file(io.BytesIO(pixel), mimetype='image/gif')

if __name__ == '__main__':
    app.run(debug=True)
