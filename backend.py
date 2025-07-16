from flask import Flask, request, make_response
import os, datetime, pytz, gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
CREDS_FILE = 'credentials.json'
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
IST = pytz.timezone("Asia/Kolkata")

def get_spreadsheet():
    if not os.path.exists(CREDS_FILE):
        google_json = os.environ.get("GOOGLE_JSON", "")
        if not google_json.strip():
            raise Exception("❌ GOOGLE_JSON not found.")
        with open(CREDS_FILE, "w") as f:
            f.write(google_json)
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

@app.route('/track', methods=['GET'])
def track_email_open():
    try:
        spreadsheet = get_spreadsheet()
        sheet_name = request.args.get('sheet')
        row = request.args.get('row')
        email_param = request.args.get('email', '').strip().lower()
        user_agent = request.headers.get('User-Agent', '')
        now = datetime.datetime.now(IST)

        if not sheet_name or not row or not email_param:
            return '', 204
        if 'googleimageproxy' in user_agent.lower() or 'googleusercontent' in user_agent.lower():
            return '', 204

        worksheet = spreadsheet.worksheet(sheet_name)
        row = int(row)
        sheet_email = worksheet.cell(row, 3).value  # Column C
        if not sheet_email or sheet_email.strip().lower() != email_param:
            return '', 204

        open_status = worksheet.cell(row, 10).value  # Column J
        if open_status and open_status.strip().lower() == "yes":
            return '', 204

        send_time_str = worksheet.cell(row, 9).value  # Column I
        if not send_time_str:
            return '', 204

        try:
            send_time = datetime.datetime.strptime(send_time_str, "%d-%m-%Y %H:%M:%S").replace(tzinfo=IST)
            delay = (now - send_time).total_seconds()
            if delay < 10:
                return '', 204
        except:
            return '', 204

        worksheet.update_cell(row, 10, "Yes")
        worksheet.update_cell(row, 11, now.strftime("%d-%m-%Y %H:%M:%S"))
    except Exception as e:
        print(f"❌ Error in /track: {e}")
        return make_response('Internal Server Error', 500)
    return '', 204

@app.route('/')
def home():
    return '✅ Email tracking backend is live'
