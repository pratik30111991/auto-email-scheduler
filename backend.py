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
def track_email_open():
    sheet = request.args.get('sheet')
    row = request.args.get('row')
    email_param = request.args.get('email', '').strip().lower()

    user_agent = request.headers.get('User-Agent', '')
    ip = request.remote_addr
    now = datetime.datetime.now()

    # Only allow update if:
    # 1. Valid email
    # 2. Not a bot or Google Proxy
    # 3. More than 45 seconds after email was sent (to skip preloading)
    if not sheet or not row or not email_param:
        return '', 204

    if "googleimageproxy" in user_agent.lower() or "google" in ip:
        print(f"[IGNORED] Google proxy hit from {ip}")
        return '', 204

    sheet = SHEET_CLIENT.open(SHEET_NAME).worksheet(sheet)
    row_index = int(row)
    current_status = sheet.cell(row_index, COLUMN_MAP['Open?']).value

    if current_status.strip().upper() == 'YES':
        return '', 204

    # Optional: Delay check logic (e.g. wait at least 60 seconds after sending)
    send_time_str = sheet.cell(row_index, COLUMN_MAP['Last Sent At']).value
    try:
        send_time = datetime.datetime.strptime(send_time_str, "%d/%m/%Y %H:%M:%S")
        if (now - send_time).total_seconds() < 60:
            print(f"[IGNORED] Too soon after send: {email_param}")
            return '', 204
    except:
        pass

    print(f"[OPEN] Row {row} marked opened by {email_param}")
    sheet.update_cell(row_index, COLUMN_MAP['Open?'], 'Yes')
    return '', 204

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
