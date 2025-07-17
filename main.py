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

# Load GOOGLE_JSON
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
except Exception as e:
    print("❌ Error loading worksheet:", str(e))
    exit(1)

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

    smtp_password = None
    from_email = None

    if "dilshad" in email.lower():
        from_email = "dilshad@unlistedradar.in"
        smtp_password = os.getenv("SMTP_DILSHAD")
    elif "nana" in email.lower():
        from_email = "nana@unlistedradar.in"
        smtp_password = os.getenv("SMTP_NANA")
    elif "gaurav" in email.lower():
        from_email = "gaurav@unlistedradar.in"
        smtp_password = os.getenv("SMTP_GAURAV")
    elif "info" in email.lower():
        from_email = "info@unlistedradar.in"
        smtp_password = os.getenv("SMTP_INFO")
    else:
        from_email = "sales@unlistedradar.in"
        smtp_password = os.getenv("SMTP_SALES")

    if not smtp_password:
        worksheet.update_cell(i, 8, "Skipped: Missing SMTP Password")
        continue

    tracking_url = f"{os.getenv('TRACKING_BACKEND_URL')}/track?email={email}&sheet=Sales_Mails&row={i}"
    html_message = f"{message}<img src='{tracking_url}' width='1' height='1' style='display:none;'>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Unlisted Radar <{from_email}>"
    msg["To"] = email

    part = MIMEText(html_message, "html")
    msg.attach(part)

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(from_email, smtp_password)
            server.sendmail(from_email, email, msg.as_string())

        worksheet.update_cell(i, 8, "Sent")
        worksheet.update_cell(i, 9, now.strftime("%d/%m/%Y %H:%M:%S"))
    except Exception as e:
        worksheet.update_cell(i, 8, f"Failed: {str(e)}")
