from flask import Flask, request, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import datetime

app = Flask(__name__)

# Setup Google Sheets access
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_file = "credentials.json"

if not os.path.exists(creds_file):
    with open(creds_file, "w") as f:
        f.write(os.environ["GOOGLE_JSON"])

creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
client = gspread.authorize(creds)

@app.route("/")
def home():
    return "✅ Email Tracking Backend Live"

@app.route("/track", methods=["GET"])
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return Response(status=400)

    try:
        sheet_id = os.environ.get("SHEET_ID")
        if not sheet_id:
            return Response(status=500)

        sheet = client.open_by_key(sheet_id)
        worksheet = sheet.worksheet(sheet_name)

        row_index = int(row)

        values = worksheet.row_values(row_index)
        headers = worksheet.row_values(1)

        email_col = headers.index("Email ID") + 1
        open_col = headers.index("Open?") + 1
        open_ts_col = headers.index("Open Timestamp") + 1

        email_value = worksheet.cell(row_index, email_col).value

        # Safety: Ensure email matches
        if email_value and email_value.strip().lower() != email_param.strip().lower():
            return Response(status=204)

        open_status = worksheet.cell(row_index, open_col).value
        if open_status != "Yes":
            now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            worksheet.update_cell(row_index, open_col, "Yes")
            worksheet.update_cell(row_index, open_ts_col, now_str)
            print(f"✅ Marked opened for row {row_index}, email={email_param}")
        else:
            print(f"🔁 Already marked open for row {row_index}")

    except Exception as e:
        print(f"❌ Error in /track: {e}")
        return Response(status=500)

    # Return 1x1 transparent GIF with no-cache headers
    pixel = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF!' \
            b'\xF9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01' \
            b'\x00\x00\x02\x02D\x01\x00;'
    return Response(pixel, mimetype='image/gif', headers={
        'Cache-Control': 'no-cache, no-store, must-revalidate'
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
