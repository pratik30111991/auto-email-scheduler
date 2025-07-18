from flask import Flask, request, Response
import gspread
import os
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz

app = Flask(__name__)
INDIA_TZ = pytz.timezone('Asia/Kolkata')

# Load credentials from Render environment
GOOGLE_JSON = os.environ.get('GOOGLE_JSON')
SHEET_ID = os.environ.get('SHEET_ID')

# Authorize Sheets client
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(GOOGLE_JSON), scope)
client = gspread.authorize(creds)

@app.route("/track", methods=["GET"])
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return Response(status=400)

    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        row = int(row)

        headers = sheet.row_values(1)
        if "Email ID" not in headers or "Open?" not in headers or "Open Timestamp" not in headers:
            return Response(status=400)

        email_col = headers.index("Email ID") + 1
        open_col = headers.index("Open?") + 1
        timestamp_col = headers.index("Open Timestamp") + 1

        actual_email = sheet.cell(row, email_col).value.strip().lower()
        if actual_email != email_param.strip().lower():
            return Response(status=204)  # Proxy mismatch

        # Update only if not already opened
        if sheet.cell(row, open_col).value != "Yes":
            now = datetime.now(INDIA_TZ).strftime('%d-%m-%Y %H:%M:%S')
            sheet.update_cell(row, open_col, "Yes")
            sheet.update_cell(row, timestamp_col, now)

    except Exception as e:
        print("Error:", e)
        return Response(status=500)

    # Return 1x1 transparent gif pixel
    pixel = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF!' \
            b'\xF9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01' \
            b'\x00\x00\x02\x02D\x01\x00;'
    return Response(pixel, mimetype='image/gif')

@app.route("/")
def home():
    return "Tracking backend is live!", 200
