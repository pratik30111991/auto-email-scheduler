import os
import time
import gspread
import ssl
import smtplib
import imaplib
import pytz

from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

# === CONFIG ===
SPREADSHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"
JSON_FILE      = "credentials.json"
INDIA_TZ       = pytz.timezone("Asia/Kolkata")

# === 1) WRITE GOOGLE_JSON SECRET TO FILE ===
raw = os.environ.get("GOOGLE_JSON")
if not raw:
    print("❌ Missing GOOGLE_JSON in env")
    exit(1)
with open(JSON_FILE, "w") as f:
    f.write(raw)

# === 2) AUTHORIZE & OPEN SHEET ===
scope  = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
creds  = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
client = gspread.authorize(creds)
sheet  = client.open_by_key(SPREADSHEET_ID)

# === 3) LOAD DOMAIN DETAILS ===
domain_cfg = sheet.worksheet("Domain Details").get_all_records()

# === 4) EMAIL SENDER ===
def send_email(smtp_server, port, user, pwd, to_addr, subject, html_body, imap_server, image_files):
    msg = MIMEMultipart("related")
    msg["From"]    = user
    msg["To"]      = to_addr
    msg["Subject"] = subject

    # build HTML part
    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    alt.attach(MIMEText(html_body, "html"))

    # attach inline images
    for path, cid in image_files:
        try:
            img = MIMEImage(open(path,"rb").read())
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=os.path.basename(path))
            msg.attach(img)
        except Exception as e:
            print(f"⚠️ Could not attach {path}: {e}")

    # send SMTP
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=ctx) as s:
            s.login(user, pwd)
            s.sendmail(user, to_addr, msg.as_string())
        # save to Sent
        im = imaplib.IMAP4_SSL(imap_server or smtp_server)
        im.login(user, pwd)
        im.append("Sent", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        im.logout()
        return True
    except Exception as e:
        print(f"❌ send_email error for {to_addr}: {e}")
        return False

# === 5) LOOP SUBSHEETS & SEND ===
for dom in domain_cfg:
    sheet_name = dom["SubSheet Name"]
    smtp_srv   = dom["SMTP Server"]
    imap_srv   = dom.get("IMAP Server", smtp_srv)
    port       = int(dom["Port"])
    sender     = dom["Email ID"]
    pwd_env    = {
      "Dilshad_Mails":"SMTP_DILSHAD",
      "Nana_Mails":"SMTP_NANA",
      "Gaurav_Mails":"SMTP_GAURAV",
      "Info_Mails":"SMTP_INFO"
    }.get(sheet_name)

    pwd = os.environ.get(pwd_env or "")
    if not pwd:
        print(f"❌ No password for {sheet_name}")
        continue

    try:
        sub = sheet.worksheet(sheet_name)
        rows = sub.get_all_records()
    except Exception as e:
        print(f"⚠️ Cannot open {sheet_name}: {e}")
        continue

    for idx, row in enumerate(rows, start=2):
        status = row.get("Status","").lower()
        sched  = row.get("Schedule Date & Time","").strip()
        if "mail sent" in status:
            continue

        # parse schedule
        dt = None
        for fmt in ["%d-%m-%Y %H:%M:%S","%d/%m/%Y %H:%M:%S","%d-%m-%Y %H:%M","%d/%m/%Y %H:%M"]:
            try:
                dt = INDIA_TZ.localize(datetime.strptime(sched,fmt))
                break
            except:
                pass
        if not dt:
            sub.update_cell(idx,8,"Skipped: Invalid Date")
            continue

        now  = datetime.now(INDIA_TZ)
        diff = (now - dt).total_seconds()
        if diff < 0 or diff > 3600:
            continue  # only send if within the last hour

        # build personalized HTML from sheet Message
        name    = row.get("Name","").strip()
        to_addr = row.get("Email ID","").strip()
        subj    = row.get("Subject","").strip()
        msg_txt = row.get("Message","").strip()

        # prepend greeting & convert CR→<br>
        first   = name.split()[0] if name else "Friend"
        html    = f"<p>Hi {first},</p><p>{msg_txt.replace(chr(10),'<br>')}</p>"

        # append images after user's content
        html += """
            <br>
            <img src="cid:murderimg" style="max-width:300px;"><br><br>
            <img src="cid:bannerimg" style="max-width:400px;">
        """

        # send it
        ok = send_email(
            smtp_srv, port, sender, pwd, to_addr, subj,
            html, imap_srv,
            [("Murder on the Mind.png","murderimg"),
             ("mail_banner.png","bannerimg")]
        )

        ts = now.strftime("%d-%m-%Y %H:%M:%S")
        if ok:
            sub.update_cell(idx,8,"Mail Sent Successfully")
            sub.update_cell(idx,9,ts)
        else:
            sub.update_cell(idx,8,"Error sending mail")
