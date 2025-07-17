import os
import json
import base64
import smtplib
import pytz
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

TRACKING_BACKEND_URL = os.environ.get("TRACKING_BACKEND_URL", "").rstrip("/")
LOCAL_TZ = pytz.timezone("Asia/Kolkata")

def load_credentials():
    with open("credentials.json", "r") as f:
        return json.load(f)

def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(load_credentials(), scope)
    return gspread.authorize(creds)

def parse_datetime(dt_str):
    try:
        dt = datetime.strptime(dt_str, "%d/%m/%Y %H:%M:%S")
        return LOCAL_TZ.localize(dt)
    except:
        return None

def send_email(smtp_config, to_email, subject, html_body):
    msg = MIMEMultipart()
    msg['From'] = f"{smtp_config['email']}"
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html'))

    with smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port']) as server:
        server.login(smtp_config['email'], smtp_config['password'])
        server.sendmail(smtp_config['email'], to_email, msg.as_string())

def process_subsheet(sheet, smtp_config, sheet_name):
    headers = sheet.row_values(1)
    col_map = {h.strip(): i for i, h in enumerate(headers)}

    for idx, row in enumerate(sheet.get_all_values()[1:], start=2):
        try:
            name = row[col_map.get("Name", -1)].strip()
            to_email = row[col_map.get("Email ID", -1)].strip()
            subject = row[col_map.get("Subject", -1)].strip()
            message = row[col_map.get("Message", -1)].strip()
            schedule_str = row[col_map.get("Schedule Date & Time", -1)].strip()
            status = row[col_map.get("Status", -1)].strip()

            if not name or not to_email:
                print(f"⛔ Row {idx} skipped — missing name/email.")
                continue

            if status:
                continue

            scheduled_time = parse_datetime(schedule_str)
            if not scheduled_time:
                print(f"⛔ Row {idx} skipped — invalid date: {schedule_str}")
                continue

            if datetime.now(LOCAL_TZ) < scheduled_time:
                continue

            tracking_url = f"{TRACKING_BACKEND_URL}/track?sheet={sheet_name}&row={idx}&email={to_email}"
            full_message = f"{message}<br><img src='{tracking_url}' width='1' height='1'>"
            send_email(smtp_config, to_email, subject, full_message)

            sheet.update_cell(idx, col_map["Status"] + 1, "✅ Sent")
            sheet.update_cell(idx, col_map["Timestamp"] + 1, datetime.now(LOCAL_TZ).strftime("%d/%m/%Y %H:%M:%S"))
            print(f"✅ Row {idx} email sent to {to_email}")

        except Exception as e:
            print(f"⛔ Row {idx} failed: {e}")
            sheet.update_cell(idx, col_map["Status"] + 1, "❌ Failed")
            sheet.update_cell(idx, col_map["Timestamp"] + 1, datetime.now(LOCAL_TZ).strftime("%d/%m/%Y %H:%M:%S"))

def get_smtp_configs():
    smtp_configs = {}
    sheet_id = os.environ.get("SHEET_ID")
    client = get_gspread_client()
    sheet = client.open_by_key(sheet_id).worksheet("Domain Details")

    for row in sheet.get_all_records():
        name = row['SubSheet Name'].strip()
        smtp_env = name.split("_")[0].upper()
        password = os.environ.get(f"SMTP_{smtp_env}")
        if not password:
            print(f"❌ Missing SMTP details for domain {smtp_env}")
            continue

        smtp_configs[name] = {
            "server": row['SMTP Server'].strip(),
            "port": int(row['Port']),
            "email": row['Email ID'].strip(),
            "password": password
        }
    return smtp_configs

def main():
    client = get_gspread_client()
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        print("❌ Missing SHEET_ID in environment")
        return

    smtp_configs = get_smtp_configs()
    spreadsheet = client.open_by_key(sheet_id)
    for sub_sheet in spreadsheet.worksheets():
        name = sub_sheet.title
        if name == "Domain Details":
            continue
        if name not in smtp_configs:
            print(f"❌ Missing SMTP config for sub-sheet: {name}")
            continue
        process_subsheet(sub_sheet, smtp_configs[name], name)

if __name__ == "__main__":
    main()
