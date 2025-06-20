import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib, ssl, imaplib
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import time
import os

# === CONSTANTS ===
INDIA_TZ = pytz.timezone("Asia/Kolkata")
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE = "credentials.json"

# === SAVE JSON SECRET TO FILE ===
with open(JSON_FILE, "w") as f:
    f.write(os.environ["GOOGLE_JSON"])

# === CONNECT TO SHEET ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID)

# === GET DOMAIN CONFIGS ===
domain_sheet = sheet.worksheet("Domain Details")
domain_configs = domain_sheet.get_all_records()

# === SEND EMAIL FUNCTION ===
def send_email(smtp_server, port, sender_email, password, recipient, subject, body, imap_server=""):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient, msg.as_string())
        print(f"✅ Sent to {recipient}")

        # Save to Sent folder
        imap = imaplib.IMAP4_SSL(imap_server or smtp_server)
        imap.login(sender_email, password)
        imap.append("Sent", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
        return True
    except Exception as e:
        print(f"❌ Failed to send to {recipient}: {e}")
        return False

# === TRY ALL SUBSHEETS ===
for domain in domain_configs:
    sub_sheet_name = domain["SubSheet Name"]
    smtp_server = domain["SMTP Server"]
    imap_server = domain.get("IMAP Server", smtp_server)
    port = int(domain["Port"])
    sender_email = domain["Email ID"]

    # ✅ Use correct password for each domain
    key_map = {
        "Dilshad_Mails": "SMTP_DILSHAD",
        "Nana_Mails": "SMTP_NANA",
        "Gaurav_Mails": "SMTP_GAURAV",
        "Info_Mails": "SMTP_INFO"
    }
    env_key = key_map.get(sub_sheet_name)
    password = os.environ.get(env_key)

    try:
        subsheet = sheet.worksheet(sub_sheet_name)
        rows = subsheet.get_all_records()
    except Exception as e:
        print(f"⚠️ Could not read sheet {sub_sheet_name}: {e}")
        continue

    for i, row in enumerate(rows, start=2):
        status = row.get("Status", "").strip().lower()
        schedule = row.get("Schedule Date & Time", "").strip()

        if "mail sent" in status:
            continue

        # Parse date/time
        parsed = False
        for fmt in ["%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M"]:
            try:
                schedule_dt = INDIA_TZ.localize(datetime.strptime(schedule, fmt))
                parsed = True
                break
            except:
                continue

        if not parsed:
            subsheet.update_cell(i, 8, "Skipped: Invalid Date Format")
            continue

        now = datetime.now(INDIA_TZ)
        diff = (now - schedule_dt).total_seconds()

        if diff < 0 or diff > 3600:
            continue  # not within current 1-minute window

        name = row.get("Name", "").strip()
        email = row.get("Email ID", "").strip()
        subject = row.get("Subject", "").strip()
        message = row.get("Message", "").strip()
        first_name = name.split()[0] if name else "Friend"
        full_message = f"Hi {first_name},\n\n{message}"

        success = send_email(
            smtp_server=smtp_server,
            port=port,
            sender_email=sender_email,
            password=password,
            recipient=email,
            subject=subject,
            body=full_message,
            imap_server=imap_server
        )

        if success:
            subsheet.update_cell(i, 8, "Mail Sent Successfully")
            subsheet.update_cell(i, 9, now.strftime("%d-%m-%Y %H:%M:%S"))
        else:
            subsheet.update_cell(i, 8, "Error sending mail")
