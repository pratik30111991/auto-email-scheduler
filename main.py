import os
import smtplib
import imaplib
import email.message
import pytz
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
from dotenv import load_dotenv

load_dotenv()

TRACKING_BASE = os.getenv("TRACKING_BACKEND_URL")
IS_MANUAL = os.getenv("IS_MANUAL", "false").lower() == "true"

# Google Sheet auth
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q")

# Subsheet mapping
SHEET_SMTP_MAPPING = {
    "Dilshad_Mails": "SMTP_DILSHAD",
    "Nana_Mails": "SMTP_NANA",
    "Gaurav_Mails": "SMTP_GAURAV",
    "Info_Mails": "SMTP_INFO",
}

# IST timezone
IST = pytz.timezone("Asia/Kolkata")
now_ist = datetime.now(IST)

# Function to send email and return success

def send_email(smtp_user, smtp_pass, to_email, subject, html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = to_email

    html_part = MIMEText(html, "html")
    msg.attach(html_part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, to_email, msg.as_string())

# Loop through all subsheets
for sub_sheet_name in SHEET_SMTP_MAPPING:
    smtp_credentials = os.getenv(SHEET_SMTP_MAPPING[sub_sheet_name])
    if not smtp_credentials or ":" not in smtp_credentials:
        print(f"❌ SMTP credentials missing or invalid for {sub_sheet_name}")
        continue

    smtp_user, smtp_pass = smtp_credentials.split(":", 1)

    try:
        worksheet = sheet.worksheet(sub_sheet_name)
    except:
        print(f"❌ Sheet {sub_sheet_name} not found")
        continue

    data = worksheet.get_all_records()
    for i, row in enumerate(data, start=2):
        schedule_str = row.get("Schedule Date & Time")
        status = row.get("Status")

        if not schedule_str or status:
            continue

        try:
            schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_dt = IST.localize(schedule_dt)
        except:
            print(f"⛔ Row {i} invalid datetime: '{schedule_str}' — skipping")
            continue

        if schedule_dt > now_ist:
            continue

        name = row.get("Name")
        email_id = row.get("Email ID")
        subject = row.get("Subject")
        message = row.get("Message")

        if not email_id or not subject or not message:
            worksheet.update_acell(f"H{i}", "Skipped: Missing required fields")
            continue

        # Inject pixel
        tracking_pixel = (
            f'<img src="{TRACKING_BASE}/track?sheet={sub_sheet_name}&row={i}" '
            'alt="" width="1" height="1" style="opacity:0;position:absolute;left:-9999px;">
        )

        full_body = message
        if "</body>" in message.lower():
            full_body = message.replace("</body>", tracking_pixel + "</body>")
        else:
            full_body = message + tracking_pixel

        first_name = name.split()[0] if name else "there"
        full_body = f"Hi <b>{first_name}</b>,<br><br>{full_body}"

        try:
            send_email(smtp_user, smtp_pass, email_id, subject, full_body)
            worksheet.update_acell(f"H{i}", "Mail Sent Successfully")
            worksheet.update_acell(f"I{i}", now_ist.strftime("%d-%m-%Y %H:%M:%S"))
        except Exception as e:
            print(f"❌ Failed to send to row {i}: {e}")
            worksheet.update_acell(f"H{i}", f"Failed: {str(e)}")

        if IS_MANUAL:
            sleep(1)
