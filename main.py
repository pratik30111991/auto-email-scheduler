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
ENABLED_SHEETS = os.getenv("ENABLED_SHEETS", "")

if not GOOGLE_JSON or not SHEET_ID or not TRACKING_BACKEND_URL:
    raise Exception("Missing GOOGLE_JSON, SHEET_ID, or TRACKING_BACKEND_URL")

enabled_sheet_list = [s.strip() for s in ENABLED_SHEETS.split(",") if s.strip()]

# Authenticate with Google Sheets
creds = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(SHEET_ID)

# Timezone
ist = pytz.timezone("Asia/Kolkata")

# SMTP cache
smtp_cache = {}

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
        sheet.update_cell(row_index, headers_map["Status"], "Failed: Name/Email missing")
        return

    try:
        sched_dt = datetime.strptime(sched, "%d/%m/%Y %H:%M:%S")
        sched_dt = ist.localize(sched_dt)
    except:
        sheet.update_cell(row_index, headers_map["Status"], "Skipped: Invalid Date Format")
        return

    if datetime.now(ist) < sched_dt:
        return

    sender = None
    smtp_host = None
    smtp_port = None
    smtp_pass = None

    # Use correct sender based on sheet
    sheet_name = sheet.title
    env_key = "SMTP_" + sheet_name.split("_")[0].upper()
    smtp_pass = os.getenv(env_key)

    domain_map = {
        "Dilshad_Mails": ("dilshad@ticketingplatform.live", "tuesday.mxrouting.net", 465),
        "Nana_Mails": ("nana_kante@ticketingplatform.live", "md-114.webhostbox.net", 465),
        "Gaurav_Mails": ("gaurav@ticketingplatform.live", "tuesday.mxrouting.net", 465),
        "Info_Mails": ("info@ticketingplatform.live", "tuesday.mxrouting.net", 465),
        "Sales_Mails": ("sales@unlistedradar.in", "md-114.webhostbox.net", 465),
    }

    if sheet_name in domain_map and smtp_pass:
        sender, smtp_host, smtp_port = domain_map[sheet_name]
    else:
        sheet.update_cell(row_index, headers_map["Status"], "Failed: SMTP config missing")
        return

    # Compose email
    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = email
    msg["Subject"] = subject
    track_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={row_index}&email={email}"
    html = message + f"<img src='{track_url}' width='1' height='1' />"
    msg.attach(MIMEText(html, "html"))

    try:
        if (smtp_host, sender) not in smtp_cache:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port)
            server.login(sender, smtp_pass)
            smtp_cache[(smtp_host, sender)] = server
        else:
            server = smtp_cache[(smtp_host, sender)]

        server.send_message(msg)
        now_str = datetime.now(ist).strftime("%d/%m/%Y %H:%M:%S")
        sheet.update_cell(row_index, headers_map["Status"], "Mail Sent Successfully")
        sheet.update_cell(row_index, headers_map["Timestamp"], now_str)
        print(f"✅ Sent: {email} from {sender} via {smtp_host}")
    except Exception as e:
        print(f"❌ Error sending to {email}: {e}")
        sheet.update_cell(row_index, headers_map["Status"], "Failed to Send")

# Go through each sheet
for sheet in spreadsheet.worksheets():
    if sheet.title == "Domain Details":
        continue
    if enabled_sheet_list and sheet.title not in enabled_sheet_list:
        print(f"⚠️ Skipping sheet not in ENABLED_SHEETS: {sheet.title}")
        continue

    print(f"📄 Processing Sheet: {sheet.title}")
    headers = sheet.row_values(1)
    headers_map = {h: i + 1 for i, h in enumerate(headers)}

    required = ["Name", "Email ID", "Schedule Date & Time", "Status", "Subject", "Message", "Timestamp", "Open?", "Open Timestamp"]
    if not all(col in headers_map for col in required):
        print(f"ⓘ Skipping invalid sheet {sheet.title}")
        continue

    data = sheet.get_all_records()
    for idx, rec in enumerate(data, start=2):
        send_from_sheet(sheet, idx, rec, headers_map)

# Close SMTP connections
for server in smtp_cache.values():
    server.quit()
