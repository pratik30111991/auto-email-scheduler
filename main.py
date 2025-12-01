import os
import json
import smtplib
import pytz
from datetime import datetime, timedelta
from time import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from gspread.utils import rowcol_to_a1

# --- Setup ---
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")

# ✅ HARD-CODED WORKING BACKEND URL
TRACKING_BACKEND_URL = "https://auto-email-scheduler.onrender.com"

if not GOOGLE_JSON or not SHEET_ID:
    raise Exception("Missing GOOGLE_JSON or SHEET_ID")

creds = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(SHEET_ID)

ist = pytz.timezone("Asia/Kolkata")

# SMTP Mapping for each sheet
domain_map = {
    "Dilshad_Mails": ("dilshad@ticketingplatform.live", "tuesday.mxrouting.net", 465),
    "Nana_Mails": ("nana_kante@ticketingplatform.live", "md-114.webhostbox.net", 465),
    "Gaurav_Mails": ("gaurav@ticketingplatform.live", "tuesday.mxrouting.net", 465),
    "Info_Mails": ("info@ticketingplatform.live", "tuesday.mxrouting.net", 465),
    "Sales_Mails": ("sales@unlistedradar.in", "md-114.webhostbox.net", 465),
    "Yatix_Mails": ("sales@yatix.co.in", "md-114.webhostbox.net", 465),
    "Yatix_Mails1": ("sales@yatix.co.in", "md-114.webhostbox.net", 465),
    "Yatix_Mails2": ("sales@yatix.co.in", "md-114.webhostbox.net", 465),
}

def send_from_sheet(sheet, row_index, row, headers_map):
    status = row.get("Status", "").strip()
    if status:
        return

    name = row.get("Name", "").strip()
    email = row.get("Email ID", "").strip()
    subject = row.get("Subject", "").strip()
    message = row.get("Message", "").strip()
    sched = row.get("Schedule Date & Time", "").strip()

    if not name or not email:
        sheet.update_cell(row_index, headers_map["Status"], "Missing Name or Email ID")
        return

    if not sched:
        return  # ⛔ Skip silently if no schedule

    try:
        sched_dt = datetime.strptime(sched, "%d/%m/%Y %H:%M:%S")
        sched_dt = ist.localize(sched_dt)
    except:
        sheet.update_cell(row_index, headers_map["Status"], "Invalid Schedule Date & Time")
        return

    now = datetime.now(ist)
    if now < sched_dt:
        return  # ⛔ Skip silently if scheduled for future

    if now - sched_dt > timedelta(minutes=30):
        sheet.update_cell(row_index, headers_map["Status"], "Skipped – Schedule Too Old")
        return

    sheet_name = sheet.title
    if sheet_name not in domain_map:
        print(f"❌ Missing SMTP mapping for {sheet_name}")
        sheet.update_cell(row_index, headers_map["Status"], "Missing SMTP mapping")
        return

    from_email, smtp_server, smtp_port = domain_map[sheet_name]
    smtp_env_key = f'SMTP_{sheet_name.split("_")[0].upper()}'
    smtp_pass = os.getenv(smtp_env_key)
    if not smtp_pass:
        print(f"❌ Missing SMTP password for {sheet_name}")
        sheet.update_cell(row_index, headers_map["Status"], "Missing SMTP password")
        return

    msg = MIMEMultipart("alternative")
    msg["From"] = from_email
    msg["To"] = email
    msg["Subject"] = subject

    # ✅ Tracking pixel
    tracking_pixel = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={row_index}&email={email}&t={int(time())}"
    html = f"""
    <html>
      <body>
        {message}
        <img src="{tracking_pixel}" width="1" height="1" style="display:none;" />
      </body>
    </html>
    """
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(from_email, smtp_pass)
            server.send_message(msg)

        now_str = now.strftime("%d-%m-%Y %H:%M:%S")
        sheet.update_cell(row_index, headers_map["Status"], "Mail Sent Successfully")
        sheet.update_cell(row_index, headers_map["Timestamp"], now_str)
        print(f"✅ Sent: {email} from {from_email} via {smtp_server}")
    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")
        sheet.update_cell(row_index, headers_map["Status"], f"Failed to Send - {str(e)}")

# Process all sheets
for sheet in spreadsheet.worksheets():
    sheet_name = sheet.title
    if sheet_name == "Domain Details" or sheet_name not in domain_map:
        continue

    print(f"📄 Processing Sheet: {sheet_name}")
    headers = sheet.row_values(1)
    headers_map = {h.strip(): i + 1 for i, h in enumerate(headers)}

    required_cols = [
        "Name", "Email ID", "Subject", "Message",
        "Schedule Date & Time", "Status", "Timestamp",
        "Open?", "Open Timestamp"
    ]
    if not all(col in headers_map for col in required_cols):
        print(f"ⓘ Skipping invalid sheet {sheet_name}")
        continue

    data = sheet.get_all_records()
    for idx, row in enumerate(data, start=2):
        send_from_sheet(sheet, idx, row, headers_map)
