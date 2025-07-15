from flask import Flask, request, make_response
import os
import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
CREDS_FILE = 'credentials.json'
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"

creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

IST = pytz.timezone("Asia/Kolkata")

@app.route('/track', methods=['GET'])
def track_email_open():
    sheet_name = request.args.get('sheet')
    row = request.args.get('row')
    email_param = request.args.get('email', '').strip().lower()

    user_agent = request.headers.get('User-Agent', '')
    ip = request.remote_addr
    now = datetime.datetime.now(IST)

    if not sheet_name or not row or not email_param:
        return '', 204

    # Block known Google Proxy hits
    if 'googleimageproxy' in user_agent.lower() or 'googleusercontent' in user_agent.lower():
        print(f"[IGNORED] Proxy hit by {email_param} (User-Agent: {user_agent})")
        return '', 204

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        row = int(row)
        open_status = worksheet.cell(row, 10).value  # Column J = "Open?"
        if open_status.strip().lower() == "yes":
            return '', 204

        # Optional: Only allow after 60 seconds to block proxy preload
        send_time_str = worksheet.cell(row, 9).value  # Column I = "Timestamp"
        try:
            send_time = datetime.datetime.strptime(send_time_str, "%d-%m-%Y %H:%M:%S")
            if (now - send_time).total_seconds() < 60:
                print(f"[IGNORED] Too soon after sending for {email_param}")
                return '', 204
        except Exception as e:
            print(f"[WARN] Could not parse send time: {e}")

        worksheet.update_cell(row, 10, "Yes")  # "Open?"
        worksheet.update_cell(row, 11, now.strftime("%d-%m-%Y %H:%M:%S"))  # "Open Timestamp"
        print(f"[OPENED] Email opened by {email_param}, row {row}")

    except Exception as e:
        print(f"❌ Error in tracking: {e}")
        return make_response('Internal Server Error', 500)

    return '', 204

@app.route('/')
def home():
    return '✅ Tracking backend is live.'

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
