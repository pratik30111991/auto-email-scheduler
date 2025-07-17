# backend.py
import os
from flask import Flask, request, Response
import gspread
from datetime import datetime
import pytz
import json

app = Flask(__name__)

# Load service account from environment variable
google_json = os.environ.get("GOOGLE_JSON")
if not google_json:
    raise ValueError("GOOGLE_JSON environment variable not set")

creds_dict = json.loads(google_json)
gc = gspread.service_account_from_dict(creds_dict)

SHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"

@app.route("/")
def index():
    return "📬 Tracker Running"

@app.route("/track", methods=["GET"])
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return Response(status=204)

    try:
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet(sheet_name)
        row = int(row)

        row_values = worksheet.row_values(row)
        headers = worksheet.row_values(1)
        col_map = {header.strip(): idx + 1 for idx, header in enumerate(headers)}

        open_col = col_map.get("Open?")
        timestamp_col = col_map.get("Open Timestamp")
        email_col = col_map.get("Email ID")

        if not open_col or not timestamp_col or not email_col:
            return Response(status=204)

        email_in_sheet = worksheet.cell(row, email_col).value.strip().lower()
        if email_in_sheet != email_param.strip().lower():
            return Response(status=204)

        already_opened = worksheet.cell(row, open_col).value
        if already_opened == "Yes":
            return Response(status=204)

        ist_now = datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%d/%m/%Y %H:%M:%S")
        worksheet.update_cell(row, open_col, "Yes")
        worksheet.update_cell(row, timestamp_col, ist_now)

    except Exception as e:
        print(f"❌ Error: {e}")
        return Response(status=500)

    # Return tracking pixel
    pixel_gif = b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00" \
                b"\xFF\xFF\xFF!\xF9\x04\x01\x00\x00\x00\x00," \
                b"\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02" \
                b"D\x01\x00;"
    return Response(pixel_gif, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
