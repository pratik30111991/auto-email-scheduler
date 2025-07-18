import os
import smtplib
import pytz
import gspread
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import quote_plus

# === Setup ===
timezone = pytz.timezone("Asia/Kolkata")
backend_url = os.environ.get("TRACKING_BACKEND_URL")
sheet_id = os.environ.get("SHEET_ID")

# === Authorize Google Sheets ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(sheet_id)

# === Column Mapping ===
column_map = {
    "Name": "name",
    "Email ID": "email",
    "Subject": "subject",
    "Message": "message",
    "Schedule Date & Time": "schedule",
    "Status": "status",
    "Timestamp": "timestamp",
    "Open?": "open",
    "Open Timestamp": "open_ts",
}

# === SMTP Mapping ===
smtp_accounts = {
    "Dilshad_Mails": os.environ.get("SMTP_DILSHAD"),
    "Nana_Mails": os.environ.get("SMTP_NANA"),
    "Gaurav_Mails": os.environ.get("SMTP_GAURAV"),
    "Info_Mails": os.environ.get("SMTP_INFO"),
    "Sales_Mails": os.environ.get("SMTP_SALES"),
}

# === Helpers ===
def get_a1(row, col_index):
    return gspread.utils.rowcol_to_a1(row, col_index)

def get_column_indexes(header_row):
    return {col.strip(): i + 1 for i, col in enumerate(header_row)}

def send_email(smtp_user, smtp_pass, from_email, to_email, subject, html):
    server = smtplib.SMTP("smtp.zoho.in", 587)
    server.starttls()
    server.login(from_email, smtp_pass)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.attach(MIMEText(html, "html"))

    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

def generate_tracking_html(email, sheet_name, row):
    pixel_url = f"{backend_url}/track?sheet={quote_plus(sheet_name)}&row={row}&email={quote_plus(email)}"
    return f'<img src="{pixel_url}" alt="" width="1" height="1" style="display:none;" />'

# === Main Execution ===
for sheet_name, smtp_cred in smtp_accounts.items():
    print(f"📄 Processing Sheet: {sheet_name}")
    if not smtp_cred or ":" not in smtp_cred:
        print(f"ⓘ Skipping invalid sheet {sheet_name}")
        continue

    smtp_email, smtp_pass = smtp_cred.split(":", 1)
    worksheet = spreadsheet.worksheet(sheet_name)
    data = worksheet.get_all_values()
    if not data or len(data) < 2:
        print(f"⚠️  No data in sheet {sheet_name}")
        continue

    headers = data[0]
    indexes = get_column_indexes(headers)

    for row_num, row in enumerate(data[1:], start=2):
        def get_val(key): return row[indexes.get(key, -1) - 1].strip() if indexes.get(key, 0) <= len(row) else ""

        name = get_val("Name")
        email = get_val("Email ID")
        subject = get_val("Subject")
        message = get_val("Message")
        schedule_str = get_val("Schedule Date & Time")

        status_col = indexes.get("Status")
        timestamp_col = indexes.get("Timestamp")
        open_col = indexes.get("Open?")
        open_ts_col = indexes.get("Open Timestamp")

        status_cell = get_a1(row_num, status_col) if status_col else None
        timestamp_cell = get_a1(row_num, timestamp_col) if timestamp_col else None

        if not name or not email:
            if status_cell:
                worksheet.update(status_cell, "Missing Name or Email ID")
            continue

        try:
            schedule_dt = datetime.strptime(schedule_str, "%d/%m/%Y %H:%M:%S")
            schedule_dt = timezone.localize(schedule_dt)
        except Exception:
            if status_cell:
                worksheet.update(status_cell, "Invalid or Missing Schedule Date & Time")
            continue

        now = datetime.now(timezone)
        if now < schedule_dt:
            continue  # not time yet

        if now >= schedule_dt:
            print(f"⏩ Skipped past schedule: {email} (Scheduled: {schedule_dt}, Now: {now})")
            if status_cell:
                worksheet.update(status_cell, "Skipped - Scheduled time already passed")
            continue

        # Send email
        try:
            html_msg = message + generate_tracking_html(email, sheet_name, row_num)
            send_email(smtp_email, smtp_pass, smtp_email, email, subject, html_msg)
            if status_cell:
                worksheet.update(status_cell, "Mail Sent Successfully")
            if timestamp_cell:
                worksheet.update(timestamp_cell, now.strftime("%d/%m/%Y %H:%M:%S"))
            if open_col:
                worksheet.update(get_a1(row_num, open_col), "No")
            if open_ts_col:
                worksheet.update(get_a1(row_num, open_ts_col), "")
        except Exception as e:
            print(f"❌ Error sending to {email}: {e}")
            if status_cell:
                worksheet.update(status_cell, f"Failed to Send - {str(e)}")
