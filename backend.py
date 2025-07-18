from flask import Flask, request, Response
import os
import gspread
import json
import pytz
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")

creds = gspread.service_account_from_dict(json.loads(GOOGLE_JSON))
sheet = creds.open_by_key(SHEET_ID)

TZ = pytz.timezone("Asia/Kolkata")

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = int(request.args.get("row", "0"))
    email = request.args.get("email", "")

    ws = sheet.worksheet(sheet_name)
    headers = ws.row_values(1)
    col = {h: i + 1 for i, h in enumerate(headers)}

    if ws.cell(row, col["Email ID"]).value.strip() != email:
        return "", 204
    if ws.cell(row, col["Open?"]).value.strip() == "Yes":
        return "", 204

    now = datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S")
    ws.update_cell(row, col["Open?"], "Yes")
    ws.update_cell(row, col["Open Timestamp"], now)

    pixel = b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00" \
            b"\xFF\xFF\xFF!\xF9\x04\x01\x00\x00\x00\x00," \
            b"\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02" \
            b"D\x01\x00;"
    return Response(pixel, mimetype="image/gif")
