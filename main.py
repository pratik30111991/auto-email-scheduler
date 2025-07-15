import os
import json
import gspread
import smtplib
import imaplib
import email.utils
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pytz

# --- Constants ---
INDIA_TZ = pytz.timezone("Asia/Kolkata")
TRACKING_URL = os.getenv("TRACKING_BACKEND_URL", "").strip("/")

# --- Load Google Sheets credentials ---
with open("credentials.json", "r") as f:
    creds_data = json.load(f)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_data, scope)
gc = gspread.authorize(credentials)

# --- Sheet details ---
spreadsheet = gc.open_by_key(creds_data["spreadsheet_id"])
sheet_name = "Sales_Mails"
worksheet = spreadsheet.worksheet(sheet_name)
rows = worksheet.get_all_records()

# --- Column map (0-based index + 1 for gspread) ---
col_map = {
    "schedule": "Schedule Date & Time",
    "status": "Status",
    "timestamp": "Timestamp",
    "email": "Email ID",
    "name": "Name",
    "open": "Open?",
    "open_time": "Open Timestamp"
}

headers = worksheet.row_values(1)
col_indices = {k: headers.index(v) + 1 for k, v in col_map.items()}

now = datetime.now(INDIA_TZ)

for idx, row in enumerate(rows, start=2):
    email_id = row.get(col_map["email"], "").strip()
    schedule_raw = row.get(col_map["schedule"], "").strip()
    status = row.get(col_map["status"], "").strip().lower()

    if not email_id or not schedule_raw:
        print(f"ℹ️ Row {idx} skipped — no email or schedule time.")
        continue

    try:
        schedule_dt = datetime.strptime(schedule_raw, "%d/%m/%Y %H:%M:%S")
        schedule_dt = INDIA_TZ.localize(schedule_dt)
    except Exception:
        print(f"⚠️ Row {idx} skipped — invalid date format: {schedule_raw}")
        continue

    if status == "mail sent successfully":
        continue

    if now < schedule_dt:
        print(f"⏳ SKIP Row {idx} — Time not matched: now={now}, schedule={schedule_dt}")
        continue

    # --- Construct HTML email ---
    name = row.get(col_map["name"], "").strip()
    subject = f"Hello {name or 'there'}"
    body = f"""
    <html>
    <body>
        <p>Dear {name},<br><br>This is a scheduled email.<br><br>Thanks!</p>
        <img src="{TRACKING_URL}/track?sheet={sheet_name}&row={idx}&email={email_id}" width="1" height="1" />
    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["From"] = "info@example.com"
    msg["To"] = email_id
    msg["Date"] = email.utils.format_datetime(now)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    try:
        smtp_server = smtplib.SMTP("smtp.gmail.com", 587)
        smtp_server.starttls()
        smtp_server.login("your@gmail.com", "your_password")  # Replace with app password or env
        smtp_server.sendmail(msg["From"], [msg["To"]], msg.as_string())
        smtp_server.quit()

        worksheet.update_cell(idx, col_indices["status"], "Mail Sent Successfully")
        worksheet.update_cell(idx, col_indices["timestamp"], now.strftime("%d-%m-%Y %H:%M:%S"))
        print(f"📤 Row {idx}: Mail Sent Successfully")

    except Exception as e:
        worksheet.update_cell(idx, col_indices["status"], f"Failed: {str(e)}")
        print(f"❌ Row {idx} failed — {str(e)}")
