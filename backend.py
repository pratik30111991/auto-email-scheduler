from flask import Flask, request, make_response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os
from datetime import datetime
import pytz

app = Flask(__name__)

GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")

if not GOOGLE_JSON or not SHEET_ID:
    raise Exception("GOOGLE_JSON or SHEET_ID missing")

creds = json.loads(GOOGLE_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(SHEET_ID)

ist = pytz.timezone("Asia/Kolkata")

@app.route("/track", methods=["GET"])
def track_open():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email = request.args.get("email")

    if not sheet_name or not row or not email:
        return make_response("", 204)

    try:
        sheet = spreadsheet.worksheet(sheet_name)
        row = int(row)
        headers = sheet.row_values(1)
        headers_map = {h.strip(): i + 1 for i, h in enumerate(headers)}
        open_col = headers_map.get("Open?")
        open_time_col = headers_map.get("Open Timestamp")
        email_col = headers_map.get("Email ID")

        if not open_col or not open_time_col:
            return make_response("", 204)

        cell_email = sheet.cell(row, email_col).value.strip().lower()
        if cell_email != email.strip().lower():
            return make_response("", 204)

        sheet.update_cell(row, open_col, "Yes")
        sheet.update_cell(row, open_time_col, datetime.now(ist).strftime("%d-%m-%Y %H:%M:%S"))

    except Exception as e:
        print("❌ Error in /track:", e)

    # Return 1x1 transparent GIF (invisible pixel)
    gif = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF!\xF9' \
          b'\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02' \
          b'\x02D\x01\x00;'
    response = make_response(gif)
    response.headers.set('Content-Type', 'image/gif')
    return response

@app.route("/")
def index():
    return "✅ Email Tracker is live!"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
