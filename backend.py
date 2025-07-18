from flask import Flask, request, send_file
import os
import io
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import pytz

app = Flask(__name__)
GOOGLE_JSON = os.getenv("GOOGLE_JSON")
SHEET_ID = os.getenv("SHEET_ID")

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID)

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email", "").strip()

    if not all([sheet_name, row, email_param]):
        return "", 204

    try:
        worksheet = sheet.worksheet(sheet_name)
        headers = worksheet.row_values(1)
        col_map = {h: i+1 for i, h in enumerate(headers)}

        email_in_sheet = worksheet.cell(int(row), col_map["Email ID"]).value.strip()
        status = worksheet.cell(int(row), col_map["Open?"]).value.strip()

        if email_param != email_in_sheet or status == "Yes":
            return "", 204

        now_str = datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%d/%m/%Y %H:%M:%S")
        worksheet.update_cell(int(row), col_map["Open?"], "Yes")
        worksheet.update_cell(int(row), col_map["Open Timestamp"], now_str)
        print(f"✅ Tracked open for row {row} on sheet {sheet_name}")
    except Exception as e:
        print("Error in tracking:", e)

    return send_file(io.BytesIO(b""), mimetype="image/png")

if __name__ == "__main__":
    app.run(host="0.0.0.0")
