import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib, ssl, imaplib
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import time
import os
import json

# === CONSTANTS ===
INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# === WRITE JSON SECRET TO FILE ===
if not os.environ.get("GOOGLE_JSON"):
    print("\u274c GOOGLE_JSON not found in environment")
    exit(1)

with open(JSON_FILE, "w") as f:
    f.write(os.environ["GOOGLE_JSON"])
print("\u2705 credentials.json written")

# === CONNECT TO SHEET ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID)
    print("\u2705 Google Sheet connected")
except Exception as e:
    print("\u274c Error connecting to Google Sheet:", e)
    exit(1)

# === GET DOMAIN CONFIGS ===
try:
    domain_sheet = sheet.worksheet("Domain Details")
    domain_configs = domain_sheet.get_all_records()
    print(f"\ud83d\udcc4 Found {len(domain_configs)} domain config(s)")
except Exception as e:
    print("\u274c Failed to read 'Domain Details':", e)
    exit(1)

# === EMAIL SENDER FUNCTION ===
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
        print(f"\u2705 Email sent to {recipient}")

        imap = imaplib.IMAP4_SSL(imap_server or smtp_server)
        imap.login(sender_email, password)
        imap.append("Sent", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
        return True
    except Exception as e:
        print(f"\u274c Failed to send email to {recipient}: {e}")
        return False

# === PROCESS EACH SUBSHEET ===
for domain in domain_configs:
    sub_sheet_name = domain["SubSheet Name"]
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
    env_key = key_map.get(sub_sheet_name)
    password = os.environ.get(env_key)

    if not password:
        print(f"\u274c No password found for {sub_sheet_name}")
        continue

    try:
        subsheet = sheet.worksheet(sub_sheet_name)
        rows = subsheet.get_all_records()
    except Exception as e:
        print(f"\u26a0\ufe0f Could not access subsheet '{sub_sheet_name}': {e}")
        continue

    for i, row in enumerate(rows, start=2):
        status = row.get("Status", "").strip().lower()
        schedule = row.get("Schedule Date & Time", "").strip()

        if status not in ["", "pending"]:
            continue

        parsed = False
        for fmt in ["%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M"]:
            try:
                schedule_dt = INDIA_TZ.localize(datetime.strptime(schedule, fmt))
                parsed = True
                break
            except:
                continue

        if not parsed:
            subsheet.update_cell(i, 8, "Skipped: Invalid Date Format")
            continue

        now = datetime.now(INDIA_TZ)
        diff = (now - schedule_dt).total_seconds()

    if diff < 0:
        print(f"⏳ Not time yet for row {i}. Scheduled: {schedule_dt}, Now: {now}")
    continue
    # ✅ Send if now >= schedule_dt — no more 5-min skip


        name = row.get("Name", "").strip()
        email = row.get("Email ID", "").strip()
        subject = row.get("Subject", "").strip()
        message = row.get("Message", "").strip()
        first_name = name.split()[0] if name else "Friend"
        full_message = f"Hi {first_name},<br><br>{message}"

        success = send_email(
            smtp_server=smtp_server,
            port=port,
            sender_email=sender_email,
            password=password,
            recipient=email,
            subject=subject,
            body=full_message,
            imap_server=imap_server
        )

        timestamp = now.strftime("%d-%m-%Y %H:%M:%S")
        if success:
            subsheet.update_cell(i, 8, "Mail Sent Successfully")
            subsheet.update_cell(i, 9, timestamp)
        else:
            subsheet.update_cell(i, 8, "Error sending mail")
