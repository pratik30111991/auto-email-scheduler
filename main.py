import os
import json
import pytz
import smtplib
import base64
import imaplib
import email.utils
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Load credentials
with open("credentials.json", "r") as f:
    creds = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)

sheet_id = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
spreadsheet = client.open_by_key(sheet_id)
timezone = pytz.timezone("Asia/Kolkata")

# Read domain SMTP credentials
SMTP_ACCOUNTS = {
    "dilshad": os.getenv("SMTP_DILSHAD"),
    "nana": os.getenv("SMTP_NANA"),
    "gaurav": os.getenv("SMTP_GAURAV"),
    "info": os.getenv("SMTP_INFO"),
    "sales": os.getenv("SMTP_SALES"),
}

tracking_backend_url = os.getenv("TRACKING_BACKEND_URL")

def send_email(smtp_cred, to_email, subject, html_content, from_name, from_email):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{from_name} <{from_email}>"
    msg["To"] = to_email

    msg.attach(MIMEText(html_content, "html"))

    smtp_server = smtplib.SMTP("smtp.gmail.com", 587)
    smtp_server.starttls()
    smtp_server.login(from_email, smtp_cred)
    smtp_server.sendmail(from_email, to_email, msg.as_string())
    smtp_server.quit()

    # IMAP append to Sent (optional)
    try:
        imap_server = imaplib.IMAP4_SSL("imap.gmail.com")
        imap_server.login(from_email, smtp_cred)
        imap_server.append('"[Gmail]/Sent Mail"', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap_server.logout()
    except Exception as e:
        print(f"⚠️ IMAP append failed: {e}")

def rowcol_to_a1(row, col):
    col_str = ""
    while col:
        col, remainder = divmod(col - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return f"{col_str}{row}"

def process_sheet(sheet_name):
    worksheet = spreadsheet.worksheet(sheet_name)
    data = worksheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    col_idx = {key: headers.index(key) for key in ["Name", "Email ID", "Subject", "Message", "Schedule Date & Time", "Status", "Timestamp", "Open?", "Open Timestamp"]}

    for i, row in enumerate(rows, start=2):
        name = row[col_idx["Name"]].strip()
        email_id = row[col_idx["Email ID"]].strip()
        subject = row[col_idx["Subject"]].strip()
        message = row[col_idx["Message"]].strip()
        schedule_str = row[col_idx["Schedule Date & Time"]].strip()
        status = row[col_idx["Status"]].strip()

        if not name or not email_id:
            print(f"⛔ Row {i} skipped — missing name/email.")
            continue

        if status == "Mail Sent Successfully":
            continue

        try:
            schedule_time = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_time = timezone.localize(schedule_time)
        except Exception as e:
            print(f"⚠️ Row {i} skipped — invalid datetime format: {schedule_str}")
            continue

        now = datetime.now(timezone)
        if now < schedule_time:
            print(f"⏳ Row {i} skipped — scheduled for future.")
            continue

        # Extract domain from From email
        domain_part = email_id.split("@")[0].lower()
        smtp_password = SMTP_ACCOUNTS.get(domain_part)

        if not smtp_password:
            print(f"❌ Missing SMTP details for domain {domain_part.upper()}")
            continue

        # Tracking pixel
        pixel_url = f"{tracking_backend_url}/track?sheet={sheet_name}&row={i}&email={email_id}"
        html_with_tracking = message + f'<img src="{pixel_url}" width="1" height="1" style="display:none"/>'

        try:
            send_email(smtp_password, email_id, subject, html_with_tracking, name, email_id)
            print(f"✅ Row {i} email sent to {email_id}")

            # Update Status and Timestamp
            worksheet.update_acell(rowcol_to_a1(i, col_idx["Status"] + 1), "Mail Sent Successfully")
            worksheet.update_acell(rowcol_to_a1(i, col_idx["Timestamp"] + 1), now.strftime("%d/%m/%Y %H:%M:%S"))
        except Exception as e:
            print(f"❌ Failed to send to {email_id} (Row {i}): {e}")

# Process all relevant sheets
sheet_names = ["Sales_Mails"]
for sheet in sheet_names:
    process_sheet(sheet)
