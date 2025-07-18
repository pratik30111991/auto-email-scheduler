import os
import json
import time
import pytz
import smtplib
import gspread
from datetime import datetime
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials

# Load credentials
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")
TRACKING_BACKEND_URL = os.getenv("TRACKING_BACKEND_URL")
if not all([GOOGLE_JSON, SHEET_ID, TRACKING_BACKEND_URL]):
    raise Exception("GOOGLE_JSON, SHEET_ID, or TRACKING_BACKEND_URL not set in environment.")

# Write credentials file
with open("credentials.json", "w") as f:
    f.write(GOOGLE_JSON)

# Connect to Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID)

# Sheet and column config
all_sheets = sheet.worksheets()
for worksheet in all_sheets:
    sheet_name = worksheet.title
    if sheet_name == "Domain Details":
        continue
    print(f"\n📄 Processing Sheet: {sheet_name}")
    records = worksheet.get_all_records()
    if not records:
        continue

    # Map column names
    header_row = worksheet.row_values(1)
    col_map = {name.strip(): i+1 for i, name in enumerate(header_row)}
    required_cols = ["Name", "Email ID", "Schedule Date & Time", "Status", "Subject", "Message"]
    if not all(k in col_map for k in required_cols):
        print(f"❌ Skipping {sheet_name}: Missing columns.")
        continue

    # Get sender config
    domain_sheet = sheet.worksheet("Domain Details")
    domain_data = domain_sheet.get_all_records()
    sender_email = next((r["Sender Email ID"] for r in domain_data if r["Sheet Name"] == sheet_name), None)
    smtp_password = os.getenv(f'SMTP_{sender_email.split("@")[0].upper()}')

    if not sender_email or not smtp_password:
        print(f"❌ Skipping {sheet_name}: Sender Email/Password missing.")
        continue

    for idx, row in enumerate(records, start=2):
        name = row.get("Name", "").strip()
        to_email = row.get("Email ID", "").strip()
        subject = str(row.get("Subject", "")).strip().replace('\n', ' ').replace('\r', '')
        message = row.get("Message", "").strip()
        schedule_str = row.get("Schedule Date & Time", "").strip()
        status = row.get("Status", "").strip()

        if not name or not to_email:
            worksheet.update_cell(idx, col_map["Status"], "Failed: Name/Email missing")
            continue
        if status and ("Successfully" in status or status.startswith("Failed")):
            continue

        try:
            schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_dt = pytz.timezone("Asia/Kolkata").localize(schedule_dt)
        except:
            worksheet.update_cell(idx, col_map["Status"], "Failed: Invalid Date Format")
            continue

        now = pytz.timezone("Asia/Kolkata").localize(datetime.now())
        if now < schedule_dt:
            continue  # Not yet time

        try:
            # Construct email
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = f"Unlisted Radar <{sender_email}>"
            msg["To"] = to_email
            msg.set_content("This email contains HTML content.")
            tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={idx}&email={to_email}"
            html_with_pixel = message + f'<img src="{tracking_url}" width="1" height="1" />'
            msg.add_alternative(html_with_pixel, subtype='html')

            with smtplib.SMTP_SSL("smtp.zoho.in", 465) as server:
                server.login(sender_email, smtp_password)
                server.send_message(msg)

            now_str = now.strftime("%d/%m/%Y %H:%M:%S")
            worksheet.update_cell(idx, col_map["Status"], "Mail Sent Successfully")
            worksheet.update_cell(idx, col_map["Timestamp"], now_str)
            print(f"✅ Sent: {to_email}")

        except Exception as e:
            worksheet.update_cell(idx, col_map["Status"], f"Failed: {str(e)}")
            print(f"❌ Failed to send to {to_email}: {e}")
