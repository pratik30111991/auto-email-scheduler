import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib, ssl, imaplib
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import time
import os
from gspread.utils import rowcol_to_a1

INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"
TRACKING_BASE = os.getenv("TRACKING_BACKEND_URL", "")

# Load credentials.json from env
with open(JSON_FILE, "w") as f:
    f.write(os.environ["GOOGLE_JSON"])

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID)
domain_sheet = sheet.worksheet("Domain Details")
domain_configs = domain_sheet.get_all_records()

def send_email(smtp_server, port, sender_email, password, recipient, subject, body, imap_server=""):
    msg = MIMEText(body, "html")
    msg["Subject"] = subject.strip().replace("\n", " ")
    msg["From"] = f"Unlisted Radar <{sender_email}>"
    msg["To"] = recipient
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient, msg.as_string())
        imap = imaplib.IMAP4_SSL(imap_server or smtp_server)
        imap.login(sender_email, password)
        imap.append("Sent", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
        return True
    except Exception as e:
        print("❌ Email sending failed:", e)
        return False

key_map = {
    "Dilshad_Mails": "SMTP_DILSHAD",
    "Nana_Mails": "SMTP_NANA",
    "Gaurav_Mails": "SMTP_GAURAV",
    "Info_Mails": "SMTP_INFO",
    "Sales_Mails": "SMTP_SALES"
}

for domain in domain_configs:
    sub_sheet_name = domain["SubSheet Name"]
    smtp_server = domain["SMTP Server"]
    imap_server = domain.get("IMAP Server", smtp_server)
    port = int(domain["Port"])
    sender_email = domain["Email ID"]
    password = os.environ.get(key_map.get(sub_sheet_name))

    if not password:
        print(f"❌ No password found for {sub_sheet_name}")
        continue

    try:
        subsheet = sheet.worksheet(sub_sheet_name)
        rows = subsheet.get_all_records()
    except Exception as e:
        print(f"⚠️ Could not access subsheet '{sub_sheet_name}': {e}")
        continue

    updates = []
    for i, row in enumerate(rows, start=2):
        name = row.get("Name", "").strip()
        email = row.get("Email ID", "").strip()
        status = row.get("Status", "").strip().lower()
        schedule = row.get("Schedule Date & Time", "").strip()

        now = datetime.now(INDIA_TZ)
        timestamp = now.strftime("%d-%m-%Y %H:%M:%S")

        if status not in ["", "pending"]:
            continue

        if not name or not email:
            updates.append((i, 8, "Failed to Send"))
            updates.append((i, 9, timestamp))
            print(f"⛔ Row {i} skipped — missing Name/Email.")
            continue

        if not schedule:
            print(f"ℹ️ Row {i} skipped — no schedule time.")
            continue  # clean skip for blank

        parsed = False
        for fmt in ["%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S"]:
            try:
                schedule_dt = INDIA_TZ.localize(datetime.strptime(schedule, fmt))
                parsed = True
                break
            except:
                continue

        if not parsed:
            updates.append((i, 8, "Skipped: Invalid Date Format"))
            updates.append((i, 9, timestamp))
            print(f"❌ Row {i} skipped — invalid date format: {schedule}")
            continue

        print(f"🕒 TIME CHECK Row {i}: now = {now.strftime('%d/%m/%Y %H:%M:%S')}, scheduled = {schedule_dt.strftime('%d/%m/%Y %H:%M:%S')}")
        if now.replace(second=0) != schedule_dt.replace(second=0):
            print(f"⏳ SKIP Row {i} — Not exact time match.")
            continue

        subject = row.get("Subject", "").strip()
        message = row.get("Message", "")
        tracking_pixel = (
            f'<img src="{TRACKING_BASE}/track?sheet={sub_sheet_name}&row={i}&email={email}" '
            'width="1" height="1" alt="." style="opacity:0;">'
        )
        full_body = f"{message}{tracking_pixel}"

        success = send_email(smtp_server, port, sender_email, password, email, subject, full_body, imap_server)
        time.sleep(1)

        status_text = "Mail Sent Successfully" if success else "Failed to Send"
        updates.append((i, 8, status_text))
        updates.append((i, 9, timestamp))
        print(f"📤 Row {i}: {status_text}")
        time.sleep(1)

    for row, col, val in updates:
        try:
            subsheet.update_acell(rowcol_to_a1(row, col), val)
            time.sleep(0.2)
        except Exception as e:
            print(f"❌ Update error row {row}, col {col}: {e}")
