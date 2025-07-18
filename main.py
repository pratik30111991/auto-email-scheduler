import os
import base64
import json
import smtplib
import pytz
import time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from gspread.utils import rowcol_to_a1

# Set timezone
INDIA_TZ = pytz.timezone('Asia/Kolkata')

# Load credentials from environment variable
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
with open("credentials.json", "w") as f:
    f.write(GOOGLE_JSON)

# Google Sheets auth
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Google Sheet ID
sheet_id = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"

# Tracking backend
TRACKING_BACKEND = "https://email-tracking-backend-17rs.onrender.com/track"

# Load domain & SMTP config
domain_sheet = client.open_by_key(sheet_id).worksheet("Domain Details")
domain_data = domain_sheet.get_all_records()

domain_config = {}
for row in domain_data:
    domain_config[row["Sheet Name"]] = {
        "smtp_server": row["SMTP Server"],
        "smtp_port": row["SMTP Port"],
        "email": row["Email"],
        "password": row["App Password"]
    }

# Process each sheet
spreadsheet = client.open_by_key(sheet_id)
sheet_list = spreadsheet.worksheets()

for sheet in sheet_list:
    sheet_name = sheet.title
    if sheet_name == "Domain Details":
        continue

    if sheet_name not in domain_config:
        print(f"Skipping unknown sheet: {sheet_name}")
        continue

    print(f"🔧 Processing Sheet: {sheet_name}")
    config = domain_config[sheet_name]
    from_email = config["email"]
    smtp_server = config["smtp_server"]
    smtp_port = config["smtp_port"]
    smtp_password = config["password"]

    data = sheet.get_all_values()
    headers = data[0]
    records = data[1:]

    header_map = {header: idx for idx, header in enumerate(headers)}

    for i, row in enumerate(records, start=2):
        try:
            name = row[header_map.get("Name", -1)].strip()
            to_email = row[header_map.get("Email ID", -1)].strip()
            subject = row[header_map.get("Subject", -1)].strip()
            message = row[header_map.get("Message", -1)].strip()
            schedule_str = row[header_map.get("Schedule Date & Time", -1)].strip()
            status = row[header_map.get("Status", -1)].strip()

            if not to_email or not name:
                continue
            if status:
                continue
            if not schedule_str:
                continue

            schedule_time = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_time = INDIA_TZ.localize(schedule_time)
            current_time = datetime.now(INDIA_TZ)

            if current_time < schedule_time:
                continue

            # Send email
            msg = MIMEMultipart("alternative")
            msg["From"] = f"Unlisted Radar <{from_email}>"
            msg["To"] = to_email
            msg["Subject"] = subject

            tracking_url = f"{TRACKING_BACKEND}?email={to_email}&sheet={sheet_name}"
            html_with_tracking = f'{message}<br><img src="{tracking_url}" width="1" height="1" />'

            msg.attach(MIMEText(html_with_tracking, "html"))

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(from_email, smtp_password)
                server.sendmail(from_email, to_email, msg.as_string())

            now_str = current_time.strftime("%d/%m/%Y %H:%M:%S")

            # Update status, timestamp
            sheet.update_acell(rowcol_to_a1(i, header_map["Status"] + 1), "Mail Sent Successfully")
            sheet.update_acell(rowcol_to_a1(i, header_map["Timestamp"] + 1), now_str)

            print(f"✅ Email sent to {to_email} ({sheet_name} row {i})")

            time.sleep(1)

        except Exception as e:
            print(f"❌ Error in row {i} of sheet {sheet_name}: {e}")
            sheet.update_acell(rowcol_to_a1(i, header_map["Status"] + 1), f"Failed to Send")
