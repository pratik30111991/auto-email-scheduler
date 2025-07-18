# main.py
import os
import json
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pytz

print("🔧 Debug — SHEET_ID:", os.getenv("SHEET_ID"))

google_json = os.getenv("GOOGLE_JSON")
if not google_json:
    print("❌ Missing GOOGLE_JSON in environment")
    exit(1)

try:
    creds_dict = json.loads(google_json)
except Exception as e:
    print("❌ Invalid GOOGLE_JSON format:", str(e))
    exit(1)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
gc = gspread.authorize(credentials)

sheet_id = os.getenv("SHEET_ID")
if not sheet_id:
    print("❌ Missing SHEET_ID in environment")
    exit(1)

try:
    sh = gc.open_by_key(sheet_id)
    worksheet = sh.worksheet("Sales_Mails")
    domain_sheet = sh.worksheet("Domain Details")
except Exception as e:
    print("❌ Error loading worksheet:", str(e))
    exit(1)

domain_data = domain_sheet.get_all_records()
domain_map = {
    row["SubSheet Name"].strip(): {
        "smtp_server": row["SMTP Server"],
        "port": int(row["Port"]),
        "email": row["Email ID"],
        "password": row["Password"]
    }
    for row in domain_data
}

rows = worksheet.get_all_records()
timezone = pytz.timezone("Asia/Kolkata")
now = datetime.now(timezone)

for i, row in enumerate(rows, start=2):
    schedule_str = row.get("Schedule Date & Time")
    email = row.get("Email ID")
    subject = row.get("Subject")
    message = row.get("Message")
    status = row.get("Status")

    if not email or not subject or not message:
        continue

    if status and status.strip().lower() == "sent":
        continue

    if not schedule_str:
        continue

    try:
        scheduled_time = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
        scheduled_time = timezone.localize(scheduled_time)
    except:
        worksheet.update_cell(i, 8, "Skipped: Invalid Date Format")
        continue

    if now < scheduled_time:
        continue

    sub_sheet_name = "Sales_Mails"
    creds = domain_map.get(sub_sheet_name)

    if not creds:
        worksheet.update_cell(i, 8, "Skipped: SMTP Config Missing")
        continue

    smtp_server = creds["smtp_server"]
    smtp_port = creds["port"]
    from_email = creds["email"]
    smtp_password = creds["password"]

    if not smtp_password:
        worksheet.update_cell(i, 8, "Skipped: Missing SMTP Password")
        continue

    tracking_url = f"{os.getenv('TRACKING_BACKEND_URL')}/track?email={email}&sheet={sub_sheet_name}&row={i}"
    html_message = f"{message}<img src='{tracking_url}' width='1' height='1' style='display:none;'>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Unlisted Radar <{from_email}>"
    msg["To"] = email
    part = MIMEText(html_message, "html")
    msg.attach(part)

    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(from_email, smtp_password)
            server.sendmail(from_email, email, msg.as_string())

        worksheet.update_cell(i, 8, "Sent")
        worksheet.update_cell(i, 9, now.strftime("%d/%m/%Y %H:%M:%S"))
    except Exception as e:
        worksheet.update_cell(i, 8, f"Failed: {str(e)}")
