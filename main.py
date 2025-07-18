import os
import json
import smtplib
import pytz
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from gspread.utils import rowcol_to_a1

# Load environment variables
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")
TRACKING_BACKEND_URL = os.getenv("TRACKING_BACKEND_URL")

if not GOOGLE_JSON or not SHEET_ID or not TRACKING_BACKEND_URL:
    raise Exception("Missing GOOGLE_JSON, SHEET_ID, or TRACKING_BACKEND_URL")

# Authenticate with Google Sheets
creds = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(SHEET_ID)

# Timezone setup
ist = pytz.timezone("Asia/Kolkata")

def send_from_sheet(sheet, row_index, row):
    # Skip completely empty rows
    if not any(str(value).strip() for value in row.values()):
        return

    status = row.get("Status", "").strip()
    if status and "Mail Sent Successfully" in status:
        return

    name = row.get("Name", "").strip()
    email = row.get("Email ID", "").strip()
    subj = row.get("Subject", "").strip()
    msg_body = row.get("Message", "").strip()
    sched = row.get("Schedule Date & Time", "").strip()

    if not name or not email:
        sheet.update_cell(row_index, headers_map["Status"], "Failed: Name/Email missing")
        return

    try:
        sched_clean = sched.replace("\xa0", " ").replace("\u200b", "").strip()
        sched_dt = datetime.strptime(sched_clean, "%d/%m/%Y %H:%M:%S")
        sched_dt = ist.localize(sched_dt)
    except Exception:
        sheet.update_cell(row_index, headers_map["Status"], "Skipped: Invalid Date Format")
        return

    if datetime.now(ist) < sched_dt:
        return

    sender = "sales@unlistedradar.in"
    smtp_pass = os.getenv("SMTP_SALES")
    if not smtp_pass:
        sheet.update_cell(row_index, headers_map["Status"], "Failed: SMTP missing")
        return

    # Build email with tracking
    msg = MIMEMultipart("alternative")
    msg["From"] = f"Unlisted Radar <{sender}>"
    msg["To"] = email
    msg["Subject"] = subj.replace("\n", " ").strip()

    tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet.title}&row={row_index}&email={email}"
    html = msg_body + f"<img src='{tracking_url}' width='1' height='1' />"
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("md-114.webhostbox.net", 465) as server:
            server.login(sender, smtp_pass)
            server.send_message(msg)

        now_str = datetime.now(ist).strftime("%d/%m/%Y %H:%M:%S")
        sheet.update_cell(row_index, headers_map["Status"], "Mail Sent Successfully")
        sheet.update_cell(row_index, headers_map["Timestamp"], now_str)
        print(f"✅ Sent {email}")
    except Exception as e:
        print(f"❌ Error sending to {email}: {e}")
        sheet.update_cell(row_index, headers_map["Status"], f"Failed: {str(e)}")

# Process all relevant sheets
for sheet in spreadsheet.worksheets():
    if sheet.title == "Domain Details":
        continue

    print(f"📄 Processing Sheet: {sheet.title}")
    headers = sheet.row_values(1)
    headers_map = {h.strip(): i + 1 for i, h in enumerate(headers)}

    required = ["Name", "Email ID", "Schedule Date & Time", "Status", "Subject", "Message", "Timestamp", "Open?", "Open Timestamp"]
    if not all(col in headers_map for col in required):
        print(f"ⓘ Skipping invalid sheet {sheet.title}")
        continue

    data = sheet.get_all_records()
    for idx, rec in enumerate(data, start=2):
        send_from_sheet(sheet, idx, rec)
