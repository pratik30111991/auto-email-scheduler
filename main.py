import os
import time
import smtplib
import imaplib
import email.utils
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Load environment variables
TRACKING_BASE = os.getenv("TRACKING_BACKEND_URL")
IS_MANUAL = os.getenv("IS_MANUAL", "false").lower() == "true"

SMTP_CREDENTIALS = {
    "Dilshad_Mails": os.getenv("SMTP_DILSHAD"),
    "Nana_Mails": os.getenv("SMTP_NANA"),
    "Gaurav_Mails": os.getenv("SMTP_GAURAV"),
    "Info_Mails": os.getenv("SMTP_INFO"),
}

# Setup Google Sheets API
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q")

# Send email function
def send_email(smtp_str, sender, receiver, subject, body):
    smtp_host, smtp_port, smtp_user, smtp_pass = smtp_str.split("|")
    msg = MIMEMultipart("alternative")
    msg['From'] = sender
    msg['To'] = receiver
    msg['Date'] = email.utils.formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    with smtplib.SMTP_SSL(smtp_host, int(smtp_port)) as server:
        server.login(smtp_user, smtp_pass)
        server.sendmail(sender, receiver, msg.as_string())

# Process each sub-sheet
for sub_sheet_name in SMTP_CREDENTIALS:
    smtp_cred = SMTP_CREDENTIALS[sub_sheet_name]
    if not smtp_cred:
        print(f"❌ SMTP credentials missing or invalid for {sub_sheet_name}")
        continue

    try:
        ws = sheet.worksheet(sub_sheet_name)
    except Exception as e:
        print(f"❌ Error accessing sheet {sub_sheet_name}: {e}")
        continue

    records = ws.get_all_records()
    now = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    for i, row in enumerate(records, start=2):
        schedule_str = row.get('Schedule Date & Time', '').strip()
        status = row.get('Status', '').strip()
        opened = row.get('Open?', '').strip()

        if status.startswith("Mail Sent"):
            continue

        try:
            scheduled_time = datetime.strptime(schedule_str, '%d/%m/%Y %H:%M:%S')
        except:
            print(f"⛔ Row {i} invalid datetime: '{schedule_str}' — skipping")
            continue

        if scheduled_time.strftime('%d/%m/%Y %H:%M:%S') > now and not IS_MANUAL:
            continue

        name = row.get("Name", "").strip()
        email_to = row.get("Email ID", "").strip()
        subject = row.get("Subject", "No Subject").strip()
        message = row.get("Message", "").strip()
        first_name = name.split()[0]

        tracking_pixel = (
            f'<img src="{TRACKING_BASE}/track?sheet={sub_sheet_name}&row={i}" '
            'alt="" width="1" height="1" style="opacity:0;position:absolute;left:-9999px;">'
        )

        if "</body>" in message.lower():
            full_body = message.replace("</body>", tracking_pixel + "</body>")
        else:
            full_body = message + tracking_pixel

        full_body = f"Hi <b>{first_name}</b>,<br><br>{full_body}"

        try:
            send_email(smtp_cred, smtp_cred.split('|')[2], email_to, subject, full_body)
            ws.update_acell(f'G{i}', 'Mail Sent Successfully')
            ws.update_acell(f'H{i}', datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
            print(f"✅ Mail sent to {email_to} in sheet {sub_sheet_name}")
        except Exception as e:
            ws.update_acell(f'G{i}', f"Error: {str(e)}")
            print(f"❌ Failed to send mail to {email_to}: {e}")
