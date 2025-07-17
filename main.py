import os
import json
import base64
import pytz
import time
import smtplib
import imaplib
import email.utils
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Load credentials
with open("credentials.json") as f:
    creds_dict = json.load(f)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(credentials)

# Constants
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
TIMEZONE = pytz.timezone("Asia/Kolkata")
TRACKING_BACKEND_URL = os.getenv("TRACKING_BACKEND_URL")

def get_domain_credentials(domain):
    smtp_env = f"SMTP_{domain.upper()}"
    if smtp_env in os.environ:
        creds = os.environ[smtp_env].split(",")
        if len(creds) >= 4:
            return {
                "email": creds[0],
                "password": creds[1],
                "smtp_server": creds[2],
                "imap_server": creds[3]
            }
    return None

def send_email(smtp_details, sender_name, to_email, subject, html_content, message_id):
    msg = MIMEMultipart('alternative')
    msg['From'] = f"{sender_name} <{smtp_details['email']}>"
    msg['To'] = to_email
    msg['Subject'] = subject
    msg['Message-ID'] = message_id
    msg['Date'] = email.utils.formatdate(localtime=True)

    msg.attach(MIMEText(html_content, 'html'))

    server = smtplib.SMTP(smtp_details['smtp_server'], 587)
    server.starttls()
    server.login(smtp_details['email'], smtp_details['password'])
    server.sendmail(smtp_details['email'], to_email, msg.as_string())
    server.quit()

    # Append to Sent folder (IMAP)
    imap = imaplib.IMAP4_SSL(smtp_details['imap_server'])
    imap.login(smtp_details['email'], smtp_details['password'])
    imap.append('Sent', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
    imap.logout()

def process_sheet(sheet_name):
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    data = sheet.get_all_values()
    headers = data[0]

    try:
        idx = lambda col: headers.index(col)
    except ValueError as e:
        print(f"❌ Missing required column: {e}")
        return

    for i, row in enumerate(data[1:], start=2):  # Row index starts from 2
        try:
            name = row[idx("Name")].strip()
            email_id = row[idx("Email ID")].strip()
            subject = row[idx("Subject")].strip()
            message = row[idx("Message")].strip()
            schedule_str = row[idx("Schedule Date & Time")].strip()
            status = row[idx("Status")].strip()

            if not name or not email_id:
                print(f"⛔ Row {i} skipped — missing name/email.")
                sheet.update_cell(i, idx("Status") + 1, "Failed to Send")
                sheet.update_cell(i, idx("Timestamp") + 1, datetime.now(TIMEZONE).strftime("%d-%m-%Y %H:%M:%S"))
                continue

            if not schedule_str:
                continue

            try:
                schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
                schedule_dt = TIMEZONE.localize(schedule_dt)
            except ValueError:
                print(f"⚠️ Row {i} skipped — invalid date format: {schedule_str}")
                continue

            if status == "Mail Sent Successfully":
                continue

            now = datetime.now(TIMEZONE)
            if now < schedule_dt:
                continue

            # Determine domain and credentials
            domain_sheet = sheet_name.split("_")[0].upper()
            smtp_details = get_domain_credentials(domain_sheet)
            if not smtp_details:
                print(f"❌ Missing SMTP details for domain {domain_sheet}")
                continue

            # Create tracking pixel
            message_id = email.utils.make_msgid()
            tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={i}&email={email_id}"
            full_html = f"{message}<img src='{tracking_url}' width='1' height='1' style='display:none;'>"

            send_email(smtp_details, "Unlisted Radar", email_id, subject, full_html, message_id)

            # Update sheet
            sheet.update_cell(i, idx("Status") + 1, "Mail Sent Successfully")
            sheet.update_cell(i, idx("Timestamp") + 1, now.strftime("%d-%m-%Y %H:%M:%S"))
            print(f"✅ Row {i} — Email sent to {email_id}")

        except Exception as e:
            print(f"❌ Error in row {i}: {e}")

# Sheets to process
sheet_list = ["Sales_Mails"]

for sheet_name in sheet_list:
    process_sheet(sheet_name)
