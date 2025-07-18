from flask import Flask, request, send_file
import gspread
import io
from datetime import datetime, timedelta
import os
import pytz
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Set timezone to IST
tz_ist = pytz.timezone("Asia/Kolkata")

# Authenticate Google Sheets
def get_gsheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = os.getenv("GOOGLE_JSON")
    if not creds_dict:
        raise Exception("Missing GOOGLE_JSON environment variable.")
    creds_dict = eval(creds_dict)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(os.getenv("SHEET_ID"))
    return sheet

@app.route("/track", methods=["GET"])
def track():
    sheet_name = request.args.get("sheet")
    row = int(request.args.get("row", 0))
    email_param = request.args.get("email", "").strip().lower()

    if not sheet_name or not row or not email_param:
        return "", 204

    try:
        sheet = get_gsheet()
        worksheet = sheet.worksheet(sheet_name)

        email_cell = worksheet.cell(row, 3).value.strip().lower()  # Column C
        status = worksheet.cell(row, 8).value  # Column H
        timestamp_str = worksheet.cell(row, 9).value  # Column I

        # Confirm email matches and email was already sent
        if email_cell != email_param or "Mail Sent Successfully" not in status:
            return "", 204

        # Check if open happened too early (proxy preload)
        if timestamp_str:
            sent_time = datetime.strptime(timestamp_str, "%d-%m-%Y %H:%M:%S")
            now = datetime.now(tz_ist)
            if (now - sent_time).total_seconds() < 5:
                return "", 204

        # Only mark if not already opened
        open_status = worksheet.cell(row, 10).value  # Column J
        if open_status != "Yes":
            worksheet.update_cell(row, 10, "Yes")  # "Open?" column
            open_time = datetime.now(tz_ist).strftime("%d-%m-%Y %H:%M:%S")
            worksheet.update_cell(row, 11, open_time)  # "Open Timestamp" column

    except Exception as e:
        print("Error:", e)
        return "", 204

    # Return 1x1 transparent pixel
    img = io.BytesIO(b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x21\xf9\x04\x01\x00\x00\x00\x00\x2c\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x4c\x01\x00\x3b')
    return send_file(img, mimetype='image/gif')

@app.route("/")
def home():
    return "Tracking backend is live!", 200

if __name__ == "__main__":
    app.run()
