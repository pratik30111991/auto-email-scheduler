import os
import base64
import smtplib
import imaplib
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# === Setup ===
IST = pytz.timezone('Asia/Kolkata')
TRACKING_BASE = os.environ.get("TRACKING_BACKEND_URL")
IS_MANUAL = os.environ.get("IS_MANUAL", "false").lower() == "true"

smtp_details = {
    'Dilshad_Mails': os.environ.get("SMTP_DILSHAD"),
    'Nana_Mails': os.environ.get("SMTP_NANA"),
    'Gaurav_Mails': os.environ.get("SMTP_GAURAV"),
    'Info_Mails': os.environ.get("SMTP_INFO"),
}

# === Google Sheet Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_json = os.getenv("GOOGLE_JSON")

if not credentials_json:
    raise ValueError("GOOGLE_JSON is not set.")

creds_dict = eval(credentials_json) if isinstance(credentials_json, str) else credentials_json
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(credentials)

spreadsheet_id = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
spreadsheet = client.open_by_key(spreadsheet_id)
sub_sheets = spreadsheet.worksheets()

now = datetime.now(IST)

for sheet in sub_sheets:
    sub_sheet_name = sheet.title
    if sub_sheet_name not in smtp_details or smtp_details[sub_sheet_name] is None:
        continue

    smtp_auth = smtp_details[sub_sheet_name]
    if ':' not in smtp_auth:
        print(f"❌ SMTP credentials missing or invalid for {sub_sheet_name}")
        continue

    smtp_email, smtp_password = smtp_auth.split(":", 1)
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    # Get column indexes
    col_idx = {h: i for i, h in enumerate(headers)}
    required = ['Email ID', 'Subject', 'Message', 'Schedule Date & Time', 'Status']
    if not all(k in col_idx for k in required):
        print(f"⛔ Missing required columns in {sub_sheet_name}")
        continue

    for i, row in enumerate(rows, start=2):  # row index in sheet = i
        try:
            status = row[col_idx['Status']].strip().lower()
            if status.startswith("mail sent") or status.startswith("skipped"):
                continue

            date_str = row[col_idx['Schedule Date & Time']].strip()
            if not date_str:
                print(f"⛔ Row {i} invalid datetime: '' — skipping")
                continue

            try:
                scheduled_dt = datetime.strptime(date_str, "%d/%m/%Y %H:%M:%S")
                scheduled_dt = IST.localize(scheduled_dt)
            except:
                print(f"⛔ Row {i} invalid datetime format: {date_str}")
                continue

            if scheduled_dt > now and not IS_MANUAL:
                continue

            to_email = row[col_idx['Email ID']].strip()
            subject = row[col_idx['Subject']].strip()
            message = row[col_idx['Message']].strip()
            name = row[col_idx.get('Name', -1)].strip() if 'Name' in col_idx else ""
            first_name = name.split()[0] if name else ""

            # === Compose email ===
            tracking_pixel = (
                f'<img src="{TRACKING_BASE}/track?sheet={sub_sheet_name}&row={i}" '
                'width="1" height="1" style="display:none;" alt="tracker">'
            )

            lower_msg = message.lower()
            if "</body>" in lower_msg:
                insert_index = lower_msg.rindex("</body>")
                full_body = message[:insert_index] + tracking_pixel + message[insert_index:]
            else:
                full_body = message + tracking_pixel

            if first_name:
                full_body = f"Hi <b>{first_name}</b>,<br><br>{full_body}"

            msg = MIMEMultipart("alternative")
            msg['Subject'] = subject
            msg['From'] = smtp_email
            msg['To'] = to_email
            msg['Date'] = email.utils.formatdate(localtime=True)

            mime_text = MIMEText(full_body, 'html')
            msg.attach(mime_text)

            # === Send email ===
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(smtp_email, smtp_password)
                server.sendmail(smtp_email, to_email, msg.as_string())

            # === Update status ===
            timestamp = datetime.now(IST).strftime("%d-%m-%Y %H:%M:%S")
            sheet.update_cell(i, col_idx['Status'] + 1, "Mail Sent Successfully")
            sheet.update_cell(i, col_idx['Timestamp'] + 1, timestamp)

            print(f"✅ Sent to {to_email} | Sheet: {sub_sheet_name} | Row: {i}")

        except Exception as e:
            print(f"❌ Row {i} failed: {e}")
