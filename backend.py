from flask import Flask, request, make_response
import os
import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

app = Flask(__name__)

# Google Sheet setup
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
CREDS_FILE = 'credentials.json'  # Upload this to Render
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# Load Google Sheets API
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPES)
client = gspread.authorize(creds)

# Timezone
IST = pytz.timezone("Asia/Kolkata")

# Helper: Validate GoogleImageProxy or not
def is_proxy(user_agent):
    return 'googleimageproxy' in user_agent.lower() or 'googleusercontent' in user_agent.lower()

# Helper: Force Capital Email
def normalize_email(email):
    return email.strip().lower()

@app.route('/track', methods=['GET'])
def track():
    sheet_name = request.args.get('sheet')
    row = request.args.get('row')
    email = request.args.get('email')

    if not (sheet_name and row and email):
        return make_response('Missing parameters', 400)

    row = int(row)
    email = normalize_email(email)

    ua = request.headers.get('User-Agent', '')
    is_gmail_proxy = is_proxy(ua)

    # Get sheet
    try:
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    except Exception as e:
        print(f"❌ Error: Unable to open sheet: {e}")
        return make_response('Sheet error', 500)

    # Read target row data
    try:
        row_data = sheet.row_values(row)
        if len(row_data) < 1:
            print(f"⚠️ Empty row {row} in sheet {sheet_name}")
            return ('', 204)

        header = sheet.row_values(1)
        col_map = {col.strip().lower(): idx+1 for idx, col in enumerate(header)}

        email_col = col_map.get("email")
        open_col = col_map.get("open?")
        timestamp_col = col_map.get("open timestamp")

        if not (email_col and open_col and timestamp_col):
            print("⚠️ Required columns missing")
            return ('', 204)

        sheet_email = sheet.cell(row, email_col).value.strip().lower()

        if sheet_email != email:
            print(f"⚠️ Email mismatch: {sheet_email} vs {email}")
            return ('', 204)

        # Optional: prevent proxy triggers if needed
        if is_gmail_proxy:
            print(f"📩 Proxy hit → sheet={sheet_name}, row={row}, email={email}, UA={ua}")
            # You can skip or allow proxies here depending on your logic
            # return ('', 204)  # If you want to block proxy hits completely
            # continue to allow real users who click "Show images"
        
        # Update only if not already opened
        open_status = sheet.cell(row, open_col).value
        if open_status.strip().lower() != 'yes':
            now = datetime.datetime.now(IST).strftime("%d/%m/%Y %I:%M:%S %p")
            sheet.update_cell(row, open_col, 'Yes')
            sheet.update_cell(row, timestamp_col, now)
            print(f"✅ SUCCESS: Open? marked 'Yes' in sheet '{sheet_name}', row {row}")
        else:
            print(f"ℹ️ Already marked open for row {row}, skipping update")

        return ('', 200)

    except Exception as e:
        print(f"❌ Error processing row {row}: {e}")
        return make_response('Internal error', 500)

@app.route('/')
def root():
    return 'Tracking pixel service is running.'

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
