import os
import json
import smtplib
import pytz
from datetime import datetime
import gspread
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from oauth2client.service_account import ServiceAccountCredentials
from gspread.utils import rowcol_to_a1

# Timezone
INDIA_TZ = pytz.timezone("Asia/Kolkata")

# Load credentials
with open("credentials.json", "w") as f:
    f.write(os.environ["GOOGLE_JSON"])

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(credentials)

# Sheet ID from GitHub Secrets
SHEET_ID = os.environ["SHEET_ID"]
sheet = client.open_by_key(SHEET_ID)

# SMTP account mapping
SMTP_CREDENTIALS = {
    "Dilshad": os.environ.get("SMTP_DILSHAD"),
    "Nana": os.environ.get("SMTP_NANA"),
    "Gaurav": os.environ.get("SMTP_GAURAV"),
    "Info": os.environ.get("SMTP_INFO"),
    "Sales": os.environ.get("SMTP_SALES")
}

TRACKING_BACKEND_URL = os.environ["TRACKING_BACKEND_URL"]

def format_datetime(dt_str):
    try:
        return INDIA_TZ.localize(datetime.strptime(dt_str, "%d/%m/%Y %H:%M:%S"))
    except Exception:
        return None

def send_email(from_email, password, to_email, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.attach(MIMEText(html_body, "html"))

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(from_email, password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

def process_sheet(worksheet):
    rows = worksheet.get_all_values()
    header = rows[0]
    col_index = {col: idx for idx, col in enumerate(header)}

    for i, row in enumerate(rows[1:], start=2):
        name = row[col_index["Name"]].strip()
        email = row[col_index["Email ID"]].strip()
        subject = row[col_index["Subject"]].strip()
        message = row[col_index["Message"]].strip()
        schedule_str = row[col_index["Schedule Date & Time"]].strip()
        status = row[col_index["Status"]].strip()

        if not name or not email:
            if not status:
                worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", "Failed: Name/Email missing")
            continue

        if status and "Mail Sent Successfully" in status:
            continue  # Skip already sent

        scheduled_dt = format_datetime(schedule_str)
        if not scheduled_dt:
            if not status:
                worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", "Skipped: Invalid Date Format")
            continue

        if datetime.now(INDIA_TZ) < scheduled_dt:
            continue  # Not yet scheduled

        domain = worksheet.title.split("_")[0]  # e.g., "Sales"
        smtp_creds = SMTP_CREDENTIALS.get(domain)
        if not smtp_creds:
            worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", "Failed: Sender Email/Password missing")
            continue

        try:
            from_email, smtp_pass = smtp_creds.split(":")
        except:
            worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", "Failed: Invalid SMTP format")
            continue

        try:
            # Add tracking pixel
            pixel_url = f'{TRACKING_BACKEND_URL}/track?sheet={worksheet.title}&row={i}&email={email}'
            full_message = f"{message}<img src='{pixel_url}' width='1' height='1' />"

            send_email(from_email, smtp_pass, email, subject, full_message)

            now_str = datetime.now(INDIA_TZ).strftime("%d/%m/%Y %H:%M:%S")
            worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", "Mail Sent Successfully")
            worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Timestamp']+1)}", now_str)
        except Exception as e:
            worksheet.update_acell(f"{rowcol_to_a1(i, col_index['Status']+1)}", f"Failed: {str(e).splitlines()[0]}")

def main():
    all_sheets = sheet.worksheets()
    for ws in all_sheets:
        title = ws.title
        if title == "Domain Details":
            continue
        print(f"📄 Processing Sheet: {title}")
        try:
            process_sheet(ws)
        except Exception as e:
            print(f"❌ Error processing {title}: {e}")

if __name__ == "__main__":
    main()
