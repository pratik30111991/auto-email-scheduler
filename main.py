import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz

# ✅ Check required environment variables
GOOGLE_JSON = os.environ.get("GOOGLE_JSON")
SHEET_ID = os.environ.get("SHEET_ID")
TRACKING_BACKEND_URL = os.environ.get("TRACKING_BACKEND_URL")

if not GOOGLE_JSON or not SHEET_ID or not TRACKING_BACKEND_URL:
    raise Exception("Missing GOOGLE_JSON, SHEET_ID, or TRACKING_BACKEND_URL")

# ✅ Save credentials to file from raw JSON
with open('credentials.json', 'w') as f:
    f.write(GOOGLE_JSON)

# ✅ Setup Google Sheets client
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# ✅ Get all sheets
sheet = client.open_by_key(SHEET_ID)
all_sheets = sheet.worksheets()

# ✅ Get SMTP creds from Domain Details
smtp_sheet = sheet.worksheet("Domain Details")
smtp_data = smtp_sheet.get_all_records()
smtp_config = {row["SubSheet Name"]: row for row in smtp_data}

# ✅ Process each subsheet
for ws in all_sheets:
    sheet_name = ws.title
    if sheet_name == "Domain Details":
        continue

    print(f"📄 Processing Sheet: {sheet_name}")

    # Skip if SMTP not configured
    if sheet_name not in smtp_config:
        print(f"⚠️ Skipping {sheet_name} (no SMTP config)")
        continue

    smtp = smtp_config[sheet_name]
    smtp_server = smtp['SMTP Server']
    smtp_port = smtp['Port']
    from_email = smtp['Email ID']
    smtp_password = smtp['Password']

    # Load sheet data
    data = ws.get_all_values()
    headers = data[0]
    rows = data[1:]

    col_map = {name.strip(): i for i, name in enumerate(headers)}

    for i, row in enumerate(rows, start=2):
        try:
            name = row[col_map.get("Name", -1)].strip()
            to_email = row[col_map.get("Email ID", -1)].strip()
            subject = row[col_map.get("Subject", -1)].strip()
            message = row[col_map.get("Message", -1)].strip()
            schedule_str = row[col_map.get("Schedule Date & Time", -1)].strip()
            status = row[col_map.get("Status", -1)].strip()
            open_status = row[col_map.get("Open?", -1)].strip()

            if status == "Mail Sent Successfully":
                continue

            if not name or not to_email:
                ws.update_cell(i, col_map["Status"] + 1, "Failed: Name/Email missing")
                continue

            if not schedule_str:
                ws.update_cell(i, col_map["Status"] + 1, "Skipped: No Schedule")
                continue

            try:
                schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
                now = datetime.now(pytz.timezone("Asia/Kolkata"))
                if now < schedule_dt:
                    continue
            except:
                ws.update_cell(i, col_map["Status"] + 1, "Skipped: Invalid Date Format")
                continue

            # Construct email with tracking pixel
            tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={i}&email={to_email}"
            html = f"""
                <html>
                    <body>
                        {message}
                        <img src="{tracking_url}" width="1" height="1" style="display:none;">
                    </body>
                </html>
            """

            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = from_email
            msg["To"] = to_email
            msg.attach(MIMEText(html, "html"))

            # Send email
            with smtplib.SMTP_SSL(smtp_server, int(smtp_port)) as server:
                server.login(from_email, smtp_password)
                server.send_message(msg)

            now_str = datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%d/%m/%Y %H:%M:%S")
            ws.update_cell(i, col_map["Status"] + 1, "Mail Sent Successfully")
            ws.update_cell(i, col_map["Timestamp"] + 1, now_str)

        except Exception as e:
            print(f"❌ Error in row {i}: {e}")
            ws.update_cell(i, col_map["Status"] + 1, f"Failed: {str(e)}")
