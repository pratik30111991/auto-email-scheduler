import os
import smtplib
import time
import json
import pytz
import imaplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.utils import rowcol_to_a1
from gspread.exceptions import APIError

# Load credentials and sheet
with open("credentials.json") as f:
    creds = json.load(f)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)

sheet_id = os.environ.get("SHEET_ID")
backend_url = os.environ.get("TRACKING_BACKEND_URL")

sheet = client.open_by_key(sheet_id)

for worksheet in sheet.worksheets():
    sheet_title = worksheet.title
    print(f"\n📄 Processing Sheet: {sheet_title}")
    all_data = worksheet.get_all_records()

    for i, row in enumerate(all_data, start=2):
        try:
            name = row.get("Name", "").strip()
            email = row.get("Email ID", "").strip()
            subject = row.get("Subject", "").strip()
            message = row.get("Message", "").strip()
            scheduled_time_str = row.get("Schedule Date & Time", "").strip()
            status = row.get("Status", "").strip()

            if not email or not name:
                worksheet.update_cell(i, 8, "Failed: Name/Email missing")
                continue

            if status:
                continue  # Already processed

            if not scheduled_time_str:
                worksheet.update_cell(i, 8, "Skipped: No Schedule Time")
                continue

            try:
                scheduled_time = datetime.strptime(scheduled_time_str, "%d/%m/%Y %H:%M:%S")
                scheduled_time = pytz.timezone("Asia/Kolkata").localize(scheduled_time)
                now = datetime.now(pytz.timezone("Asia/Kolkata"))
                if now < scheduled_time:
                    continue  # Not yet time
            except ValueError:
                worksheet.update_cell(i, 8, "Skipped: Invalid Date Format")
                continue

            # Get domain (email sender) from 'Domain Details' sheet
            domain_sheet = sheet.worksheet("Domain Details")
            domain_data = domain_sheet.get_all_records()
            from_email = None
            smtp_server = "smtp.gmail.com"
            smtp_port = 465
            smtp_password = None

            for domain_row in domain_data:
                if domain_row.get("Sheet Name", "").strip() == sheet_title:
                    from_email = domain_row.get("Email", "").strip()
                    smtp_password = os.environ.get(domain_row.get("SMTP_ENV", "").strip())
                    break

            if not from_email or not smtp_password:
                worksheet.update_cell(i, 8, "Failed: Sender Email/Password missing")
                continue

            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = f"Unlisted Radar <{from_email}>"
            msg["To"] = email

            # Tracking pixel
            tracking_url = f"{backend_url}/track?sheet={sheet_title}&row={i}&email={email}"
            html_with_pixel = message + f'<img src="{tracking_url}" alt="" width="1" height="1" style="display:none;">'
            msg.attach(MIMEText(html_with_pixel, "html"))

            # Send email
            try:
                with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                    server.login(from_email, smtp_password)
                    server.sendmail(from_email, email, msg.as_string())
            except Exception as e:
                worksheet.update_cell(i, 8, f"Failed: {str(e)}")
                continue

            # Log to IMAP Sent folder
            try:
                imap = imaplib.IMAP4_SSL("imap.gmail.com")
                imap.login(from_email, smtp_password)
                imap.append('"[Gmail]/Sent Mail"', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                imap.logout()
            except Exception as e:
                print(f"⚠️ IMAP log failed: {e}")

            # Timestamp
            now_str = now.strftime("%d/%m/%Y %H:%M:%S")
            def safe_update(cell, value, retries=3):
                for attempt in range(retries):
                    try:
                        worksheet.update_acell(cell, value)
                        return True
                    except APIError as api_err:
                        if "429" in str(api_err):
                            wait = 2 ** attempt
                            print(f"⚠️ Quota limit hit. Retrying in {wait}s...")
                            time.sleep(wait)
                        else:
                            raise
                return False

            safe_update(rowcol_to_a1(i, 8), "Sent")
            safe_update(rowcol_to_a1(i, 9), now_str)

            print(f"✅ Email sent to: {email}")

        except Exception as e:
            print(f"❌ Error in row {i}: {e}")
            try:
                worksheet.update_cell(i, 8, f"Failed: {str(e)}")
            except:
                print("⚠️ Failed to write error to sheet.")

