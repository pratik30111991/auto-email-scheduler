import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib, ssl, imaplib
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import time
import os
import json

INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
TRACKING_BASE = os.getenv("TRACKING_BACKEND_URL", "")
JSON_FILE = "credentials.json"

with open(JSON_FILE, "w") as f:
    f.write(os.environ["GOOGLE_JSON"])

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID)
domain_sheet = sheet.worksheet("Domain Details")
domain_configs = domain_sheet.get_all_records()

def send_email(smtp_server, port, sender_email, password, recipient, subject, body, imap_server=""):
    msg = MIMEText(body, "html")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient
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
    except:
        return False

for domain in domain_configs:
    subsheet = sheet.worksheet(domain["SubSheet Name"])
    smtp_server = domain["SMTP Server"]
    imap_server = domain.get("IMAP Server", smtp_server)
    port = int(domain["Port"])
    sender_email = domain["Email ID"]
    key_map = {
        "Dilshad_Mails": "SMTP_DILSHAD",
        "Nana_Mails": "SMTP_NANA",
        "Gaurav_Mails": "SMTP_GAURAV",
        "Info_Mails": "SMTP_INFO"
    }
    password = os.environ.get(key_map.get(domain["SubSheet Name"]))
    rows = subsheet.get_all_records()

    for i, row in enumerate(rows, start=2):
        status = row.get("Status", "").strip().lower()
        schedule = row.get("Schedule Date & Time", "").strip()

        if status not in ["", "pending"]:
            continue
        try:
            schedule_dt = INDIA_TZ.localize(datetime.strptime(schedule, "%d/%m/%Y %H:%M:%S"))
        except:
            continue

        now = datetime.now(INDIA_TZ)
        diff = (now - schedule_dt).total_seconds()
        if diff < 0 or diff > 300:
            continue

        name = row["Name"]
        email = row["Email ID"]
        subject = row["Subject"]
        message = row["Message"]
        first_name = name.split()[0] if name else "Friend"

        tracking_pixel = f'<img src="{TRACKING_BASE}?sheet={domain["SubSheet Name"]}&row={i}" width="1" height="1">'
        full_body = f"Hi <b>{first_name}</b>,<br>{message}<br>{tracking_pixel}"

        success = send_email(smtp_server, port, sender_email, password, email, subject, full_body, imap_server)
        timestamp = now.strftime("%d-%m-%Y %H:%M:%S")
        if success:
            subsheet.update_cell(i, 8, "Mail Sent Successfully")
            subsheet.update_cell(i, 9, timestamp)
