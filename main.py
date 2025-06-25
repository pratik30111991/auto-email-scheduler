# ================= main.py =================

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib, ssl, imaplib
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import time
import os

INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"
TRACKING_BASE = os.getenv("TRACKING_BACKEND_URL", "")
IS_MANUAL = os.getenv("IS_MANUAL", "false").lower() == "true"

# ✅ Write credentials
with open(JSON_FILE, "w") as f:
    f.write(os.environ["GOOGLE_JSON"])

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID)

domain_sheet = sheet.worksheet("Domain Details")
domain_configs = domain_sheet.get_all_records()

key_map = {
    "Dilshad_Mails": "SMTP_DILSHAD",
    "Nana_Mails": "SMTP_NANA",
    "Gaurav_Mails": "SMTP_GAURAV",
    "Info_Mails": "SMTP_INFO"
}

def send_email(smtp_server, port, sender_email, password, recipient, subject, body, imap_server=""):
    msg = MIMEText(body, "html")
    msg["Subject"] = subject.strip()
    msg["From"] = sender_email.strip()
    msg["To"] = recipient.strip()
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient, msg.as_string())

        imap = imaplib.IMAP4_SSL(imap_server or smtp_server)
        imap.login(sender_email, password)
        imap.append("Sent", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
        return True
    except Exception as e:
        print(f"❌ Email to {recipient} failed: {e}")
        return False

for domain in domain_configs:
    sub_sheet_name = domain["SubSheet Name"]
    smtp_server = domain["SMTP Server"]
    imap_server = domain.get("IMAP Server", smtp_server)
    port = int(domain["Port"])
    sender_email = domain["Email ID"]
    password = os.environ.get(key_map.get(sub_sheet_name))

    if not password:
        print(f"❌ Missing password for {sub_sheet_name}")
        continue

    try:
        subsheet = sheet.worksheet(sub_sheet_name)
        rows = subsheet.get_all_records()
    except Exception as e:
        print(f"⚠️ Cannot open sheet '{sub_sheet_name}': {e}")
        continue

    for i, row in enumerate(rows, start=2):
        status = row.get("Status", "").strip().lower()
        schedule = row.get("Schedule Date & Time", "").strip()

        if status not in ["", "pending"]:
            continue

        parsed = False
        for fmt in ["%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S"]:
            try:
                schedule_dt = INDIA_TZ.localize(datetime.strptime(schedule, fmt))
                parsed = True
                break
            except:
                continue

        if not parsed:
            print(f"⛔ Row {i} invalid datetime: '{schedule}' — skipping")
            continue

        now = datetime.now(INDIA_TZ)
        diff = (now - schedule_dt).total_seconds()

        if diff < 0:
            print(f"⏳ Too early for row {i} — Scheduled at {schedule_dt}, now is {now}")
            continue
        if diff > 300 and not IS_MANUAL:
            print(f"❌ Row {i} skipped due to delay >5 min ({diff}s)")
            continue

        name = row.get("Name", "")
        email = row.get("Email ID", "").strip()
        subject = row.get("Subject", "")
        message = row.get("Message", "")
        first_name = name.split()[0] if name else "Friend"

        tracking_pixel = f'<img src="{TRACKING_BASE}/track?sheet={sub_sheet_name}&row={i}" width="1" height="1" style="display:none;">'
        if "</body>" in message:
            full_body = f"Hi <b>{first_name}</b>,<br><br>{message.replace('</body>', tracking_pixel + '</body>')}"
        else:
            full_body = f"Hi <b>{first_name}</b>,<br><br>{message}{tracking_pixel}"

        success = send_email(smtp_server, port, sender_email, password, email, subject, full_body, imap_server)
        timestamp = now.strftime("%d-%m-%Y %H:%M:%S")

        if success:
            subsheet.update_cell(i, 8, "Mail Sent Successfully")
            subsheet.update_cell(i, 9, timestamp)
        else:
            subsheet.update_cell(i, 8, "Failed to Send")
