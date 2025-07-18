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

# ENV Variables
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")
TRACKING_BACKEND_URL = os.getenv("TRACKING_BACKEND_URL")
if not GOOGLE_JSON or not SHEET_ID or not TRACKING_BACKEND_URL:
    raise Exception("Missing GOOGLE_JSON, SHEET_ID, or TRACKING_BACKEND_URL")

# Google Auth
creds = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(SHEET_ID)

# Timezone
ist = pytz.timezone("Asia/Kolkata")

# Load SMTP accounts from Domain Details sheet
smtp_sheet = spreadsheet.worksheet("Domain Details")
smtp_data = smtp_sheet.get_all_records()
smtp_accounts = {}
for row in smtp_data:
    smtp_accounts[row["SubSheet Name"]] = row

def send_email(row_index, row, sheet, smtp_info, headers_map):
    name = row.get("Name", "").strip()
    email = row.get("Email ID", "").strip()
    subject = row.get("Subject", "").strip()
    message = row.get("Message", "").strip()
    schedule_str = row.get("Schedule Date & Time", "").strip()
    status = row.get("Status", "").strip()

    # Skip empty rows
    if not name or not email or not schedule_str:
        return

    # Skip if already sent or failed
    if status:
        return

    # Parse schedule datetime
    try:
        schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
        schedule_dt = ist.localize(schedule_dt)
    except:
        sheet.update_cell(row_index, headers_map["Status"], "Failed: Invalid Date Format")
        return

    # Skip if not yet scheduled
    if datetime.now(ist) < schedule_dt:
        return

    # Check content
    if not subject or not message:
        sheet.update_cell(row_index, headers_map["Status"], "Failed: Subject/Message missing")
        return

    # SMTP Info
    sender = smtp_info["Email ID"].strip()
    smtp_pass = smtp_info["Password"].strip()
    smtp_server = smtp_info["SMTP Server"].strip()
    smtp_port = int(smtp_info["Port"])

    # Compose email
    msg = MIMEMultipart("alternative")
    msg["From"] = f"{sender}"
    msg["To"] = email
    msg["Subject"] = subject

    # Add tracking pixel
    track_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet.title}&row={row_index}&email={email}"
    html = message + f"<img src='{track_url}' width='1' height='1' />"
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(sender, smtp_pass)
            server.send_message(msg)

        # Success: Update sheet
        now_str = datetime.now(ist).strftime("%d/%m/%Y %H:%M:%S")
        sheet.update_cell(row_index, headers_map["Status"], "Mail Sent Successfully")
        sheet.update_cell(row_index, headers_map["Timestamp"], now_str)
        print(f"✅ Sent: {email} from {sender} via {smtp_server}")
    except Exception as e:
        print(f"❌ Failed: {email} from {sender} via {smtp_server}: {e}")
        sheet.update_cell(row_index, headers_map["Status"], f"Failed: {str(e).splitlines()[0]}")

# Process each sheet
for sheet in spreadsheet.worksheets():
    title = sheet.title
    if title == "Domain Details":
        continue
    print(f"📄 Processing Sheet: {title}")

    smtp_info = smtp_accounts.get(title)
    if not smtp_info:
        print(f"⚠️ Skipping {title}: No SMTP config in Domain Details")
        continue

    headers = sheet.row_values(1)
    headers_map = {h.strip(): i + 1 for i, h in enumerate(headers)}

    required = ["Name", "Email ID", "Schedule Date & Time", "Status", "Subject", "Message", "Timestamp", "Open?", "Open Timestamp"]
    if not all(h in headers_map for h in required):
        print(f"ⓘ Skipping invalid sheet {title}")
        continue

    records = sheet.get_all_records()
    for idx, row in enumerate(records, start=2):
        send_email(idx, row, sheet, smtp_info, headers_map)
