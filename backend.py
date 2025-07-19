from flask import Flask, request, Response
import gspread
import os
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import time

app = Flask(__name__)
INDIA_TZ = pytz.timezone('Asia/Kolkata')

# Google Sheets auth
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(os.environ['GOOGLE_JSON']), scope)
client = gspread.authorize(creds)

@app.route("/track", methods=["GET"])
def track():
    try:
        sheet_name = request.args.get("sheet")
        row = request.args.get("row")
        email_param = request.args.get("email")
        t_param = request.args.get("t")

        if not sheet_name or not row or not email_param or not t_param:
            print("[⚠️ Missing Parameter]", sheet_name, row, email_param, t_param)
            return Response(status=400)

        # Skip preloading proxies (if opened in under 5 seconds)
        try:
            t_param = int(t_param)
            now_ts = int(time.time())
            if now_ts - t_param < 5:
                print("[🛑 Skipped proxy preload – too fast]", now_ts - t_param, "seconds")
                return Response(status=204)
        except:
            print("[⚠️ Invalid timestamp param]")
            return Response(status=204)

        # Read headers
        user_agent = request.headers.get("User-Agent", "").lower()
        if "google" in user_agent and "image" in user_agent:
            print(f"[🤖 Skipping Google proxy UA] {user_agent}")
            return Response(status=204)

        sheet = client.open_by_key(os.environ['SHEET_ID']).worksheet(sheet_name)
        row = int(row)
        header = sheet.row_values(1)

        email_col = header.index("Email ID") + 1
        open_col = header.index("Open?") + 1
        timestamp_col = header.index("Open Timestamp") + 1

        email_in_sheet = sheet.cell(row, email_col).value.strip().lower()
        if email_in_sheet != email_param.strip().lower():
            print(f"[⛔ Proxy Email Mismatch] Row: {row}, Sheet: {sheet_name}")
            return Response(status=204)

        open_status = sheet.cell(row, open_col).value
        if open_status != "Yes":
            now = datetime.now(INDIA_TZ).strftime('%d-%m-%Y %H:%M:%S')
            sheet.update_cell(row, open_col, "Yes")
            sheet.update_cell(row, timestamp_col, now)
            print(f"[✅ Marked Open] Row: {row}, Sheet: {sheet_name}, Time: {now}")

    except Exception as e:
        print("[❌ Error]", str(e))
        return Response(status=500)

    # 1x1 transparent GIF
    pixel = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF!' \
            b'\xF9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01' \
            b'\x00\x00\x02\x02D\x01\x00;'
    return Response(pixel, mimetype='image/gif')

@app.route("/")
def home():
    return "✅ Tracking backend is live!", 200
