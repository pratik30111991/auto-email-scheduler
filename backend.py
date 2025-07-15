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
    now = datetime.datetime.now(IST)

    if not sheet_name or not row or not email_param:
        return '', 204

    # BLOCK: Google Proxy preloading
    if 'googleimageproxy' in user_agent.lower() or 'googleusercontent' in user_agent.lower():
        print(f"[IGNORED: PROXY] {email_param} ({user_agent})")
        return '', 204

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        row = int(row)

        sheet_email = worksheet.cell(row, 3).value  # Column C = Email
        if not sheet_email or sheet_email.strip().lower() != email_param:
            print(f"[SKIP] Email mismatch at row {row}: sheet={sheet_email}, param={email_param}")
            return '', 204

        open_status = worksheet.cell(row, 10).value  # Column J = Open?
        if open_status and open_status.strip().lower() == "yes":
            return '', 204

        send_time_str = worksheet.cell(row, 9).value  # Column I = Timestamp
        allow_update = True

        if send_time_str:
            try:
                send_time = datetime.datetime.strptime(send_time_str, "%d-%m-%Y %H:%M:%S")
                delay = (now - send_time).total_seconds()
                if delay < 10:  # Delay is less than 10 seconds
                    print(f"[IGNORED] Delay only {delay:.1f}s for {email_param} — suspicious")
                    allow_update = False
            except Exception as e:
                print(f"[WARN] Invalid send timestamp for row {row}: {e}")
                allow_update = False
        else:
            print(f"[SKIP] No timestamp in row {row}")
            allow_update = False

        if allow_update:
            worksheet.update_cell(row, 10, "Yes")
            worksheet.update_cell(row, 11, now.strftime("%d-%m-%Y %H:%M:%S"))
            print(f"[✅ UPDATED] Open tracked for {email_param} (row {row})")
        else:
            print(f"[SKIPPED] Valid request but skipped update for {email_param} (row {row})")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        return make_response('Internal Server Error', 500)

    return '', 204

@app.route('/')
def home():
    return '✅ Tracking backend is live.'
