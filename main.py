import os
import smtplib
import pytz
import json
import gspread
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import quote

# Setup
IST = pytz.timezone('Asia/Kolkata')
TRACKING_BACKEND_URL = os.getenv("TRACKING_BACKEND_URL", "").rstrip("/")
SHEET_ID = os.getenv("SHEET_ID")

with open('credentials.json', 'w') as f:
    f.write(os.getenv('GOOGLE_JSON', ''))

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

def get_smtp_details(sheet):
    smtp_data = {}
    records = sheet.get_all_records()
    for row in records:
        email = row.get('Email')
        password = row.get('Password')
        if email and password:
            smtp_data[email.lower()] = password
    return smtp_data

def send_email(row, headers, smtp_passwords, sheet, row_index, sheet_name):
    try:
        email = row.get("Email ID")
        name = row.get("Name")
        subject = row.get("Subject", "").replace('\n', ' ').strip()
        message = row.get("Message", "")
        schedule_str = row.get("Schedule Date & Time")
        status = row.get("Status")
        from_email = headers[1]  # Assuming second column is sender
        smtp_password = smtp_passwords.get(from_email.lower())

        # Skip if already sent
        if status.strip().lower() == "mail sent successfully":
            return

        if not (email and name and from_email and smtp_password):
            sheet.update_cell(row_index, headers.index("Status") + 1, "Failed: Name/Email missing")
            return

        # Parse schedule time
        try:
            schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_dt = IST.localize(schedule_dt)
        except:
            sheet.update_cell(row_index, headers.index("Status") + 1, "Skipped: Invalid Date Format")
            return

        now = datetime.now(IST)
        if now < schedule_dt:
            return  # Not time yet

        msg = MIMEMultipart('alternative')
        msg['From'] = f"Unlisted Radar <{from_email}>"
        msg['To'] = email
        msg['Subject'] = subject

        # Add tracking pixel
        tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={quote(sheet_name)}&row={row_index}&email={quote(email)}"
        html_content = message + f'<img src="{tracking_url}" width="1" height="1" />'
        msg.attach(MIMEText(html_content, 'html'))

        server = smtplib.SMTP("smtp.zoho.in", 587)
        server.starttls()
        server.login(from_email, smtp_password)
        server.sendmail(from_email, email, msg.as_string())
        server.quit()

        now_str = now.strftime("%d/%m/%Y %H:%M:%S")
        sheet.update_cell(row_index, headers.index("Status") + 1, "Mail Sent Successfully")
        sheet.update_cell(row_index, headers.index("Timestamp") + 1, now_str)

    except smtplib.SMTPAuthenticationError:
        sheet.update_cell(row_index, headers.index("Status") + 1, "Failed: Authentication Error")
    except smtplib.SMTPConnectError:
        sheet.update_cell(row_index, headers.index("Status") + 1, "Failed: SMTP Connection Error")
    except Exception as e:
        error_msg = f"Failed: {str(e).strip()[:100]}"
        sheet.update_cell(row_index, headers.index("Status") + 1, error_msg)

def process_sheet(sheet_name, smtp_passwords):
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        records = sheet.get_all_records()
        headers = sheet.row_values(1)

        for idx, row in enumerate(records, start=2):  # Data starts at row 2
            send_email(row, headers, smtp_passwords, sheet, idx, sheet_name)

    except Exception as e:
        print(f"❌ Error processing sheet {sheet_name}: {e}")

# === MAIN ===
try:
    domain_sheet = client.open_by_key(SHEET_ID).worksheet("Domain Details")
    smtp_passwords = get_smtp_details(domain_sheet)

    all_sheets = client.open_by_key(SHEET_ID).worksheets()
    mail_sheets = [sheet.title for sheet in all_sheets if sheet.title.endswith("_Mails")]

    print(f"📄 Sheets Detected: {mail_sheets}")
    for sheet_name in mail_sheets:
        print(f"\n📄 Processing Sheet: {sheet_name}")
        process_sheet(sheet_name, smtp_passwords)

except Exception as e:
    print(f"❌ Global Error: {e}")
