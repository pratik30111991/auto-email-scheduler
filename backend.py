import os, json, pytz
from datetime import datetime
from flask import Flask, request, Response
import gspread
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__)

# Load Google credentials
creds = json.loads(os.environ['GOOGLE_JSON'])
gc = gspread.service_account_from_dict(creds)
SHEET_ID = os.environ['SHEET_ID']

def load_domains():
    sh = gc.open_by_key(SHEET_ID)
    dd = sh.worksheet("Domain Details").get_all_records()
    return {
        r['SubSheet Name'].strip(): {
            'smtp': r['SMTP Server'],
            'port': int(r['Port']),
            'email': r['Email ID'],
            'pass': r['Password']
        } for r in dd
    }

@app.route("/send")
def send_sched():
    tz = pytz.timezone("Asia/Kolkata")
    sh = gc.open_by_key(SHEET_ID)
    domains = load_domains()
    now = datetime.now(tz)
    for ws in sh.worksheets():
        if ws.title == "Domain Details": continue
        rows = ws.get_all_records()
        updates = []
        for idx, row in enumerate(rows, start=2):
            if row.get('Status','').lower() == 'sent': continue
            sched = row.get('Schedule Date & Time')
            if not sched: continue
            try:
                st = tz.localize(datetime.strptime(sched, "%d/%m/%Y %H:%M:%S"))
            except:
                updates.append({'range': f'H{idx}', 'values': [['Skipped:Invalid Date']]})
                continue
            if now < st: continue
            domain = domains.get(ws.title)
            if not domain:
                updates.append({'range': f'H{idx}', 'values': [['Skipped:No Domain Config']]})
                continue
            msg = MIMEMultipart("alternative")
            msg['From'] = f"Unlisted Radar <{domain['email']}>"
            msg['To'] = row['Email ID']
            msg['Subject'] = row['Subject']
            track = f"{os.getenv('TRACKING_BACKEND_URL')}/track?sheet={ws.title}&row={idx}&email={row['Email ID']}"
            html = f"{row['Message']}<img src='{track}' width='1' height='1' style='display:none;'>"
            msg.attach(MIMEText(html,'html'))
            try:
                with smtplib.SMTP_SSL(domain['smtp'], domain['port']) as s:
                    s.login(domain['email'], domain['pass'])
                    s.sendmail(domain['email'], row['Email ID'], msg.as_string())
                updates.append({'range': f'H{idx}:I{idx}', 'values': [['Sent', now.strftime("%d/%m/%Y %H:%M:%S")]]})
            except Exception as e:
                updates.append({'range': f'H{idx}', 'values': [[f"Failed: {e}"]]})
        if updates:
            ws.batch_update(updates)
    return "✅ Emails processed"

@app.route("/track")
def track():
    sheet = request.args.get('sheet')
    row = int(request.args.get('row'))
    email = request.args.get('email').strip().lower()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(sheet)
    rv = ws.row_values(1)
    cm = {v.strip(): idx+1 for idx,v in enumerate(rv)}
    if ws.cell(row, cm['Email ID']).value.strip().lower() != email: return Response(status=204)
    if ws.cell(row, cm['Open?']).value == 'Yes': return Response(status=204)
    now = datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%d/%m/%Y %H:%M:%S")
    ws.batch_update([
        {'range': f"{gspread.utils.rowcol_to_a1(row, cm['Open?'])}", 'values': [['Yes']]},
        {'range': f"{gspread.utils.rowcol_to_a1(row, cm['Open Timestamp'])}", 'values': [[now]]}
    ])
    gif = b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00" \
          b"\xFF\xFF\xFF!\xF9\x04\x01\x00\x00\x00\x00," \
          b"\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02" \
          b"D\x01\x00;"
    return Response(gif, mimetype='image/gif')

@app.route("/")
def home():
    return "🚀 Email Scheduler Active"

if __name__=="__main__":
    app.run()
