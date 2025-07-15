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

# Auth + Spreadsheet
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
    now = datetime.datetime.now(IST)

    if not sheet_name or not row or not email_param:
        print("[❌ MISSING PARAMS]", sheet_name, row, email_param)
        return '', 204

    if 'googleimageproxy' in user_agent.lower() or 'googleusercontent' in user_agent.lower():
        print(f"[IGNORED: PROXY] {email_param} ({user_agent})")
        return '', 204

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        row = int(row)

        sheet_email = worksheet.cell(row, 3).value  # Column C = Email
        if not sheet_email:
            print(f"[SKIP] Empty email in row {row}")
            return '', 204

        sheet_email = sheet_email.strip().lower()
        if sheet_email != email_param:
            print(f"[SKIP] Email mismatch at row {row}: sheet='{sheet_email}' vs param='{email_param}'")
            return '', 204

        open_status = worksheet.cell(row, 10).value  # Column J = Open?
        if open_status and open_status.strip().lower() == "yes":
            print(f"[ALREADY OPENED] Row {row}")
            return '', 204

        send_time_str = worksheet.cell(row, 9).value  # Column I = Timestamp
        if not send_time_str:
            print(f"[SKIP] No timestamp in row {row}")
            return '', 204

        try:
            send_time = datetime.datetime.strptime(send_time_str, "%d-%m-%Y %H:%M:%S")
            delay = (now - send_time).total_seconds()
            if delay < 10:
                print(f"[IGNORED] Only {delay:.1f}s since sent for {email_param} — skipping as possible proxy preload")
                return '', 204
        except Exception as e:
            print(f"[WARN] Invalid timestamp in row {row}: {send_time_str} ({e})")
            return '', 204

        # ✅ Update Open? and Timestamp
        worksheet.update_cell(row, 10, "Yes")
        worksheet.update_cell(row, 11, now.strftime("%d-%m-%Y %H:%M:%S"))
        print(f"[✅ MARKED OPEN] Row {row} for {email_param}")

    except Exception as e:
        print(f"[❌ ERROR] Tracking error for {email_param}: {e}")
        return make_response('Internal Server Error', 500)

    return '', 204


@app.route('/')
def home():
    return '✅ Email tracking backend is live (Gunicorn expected).'
