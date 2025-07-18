from flask import Flask, request, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import time
import json

app = Flask(__name__)

# Load credentials once
creds = None
gc = None
try:
    GOOGLE_JSON = os.environ.get("GOOGLE_JSON", "")
    if not GOOGLE_JSON:
        raise Exception("GOOGLE_JSON not set in environment")
    data = json.loads(GOOGLE_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
    gc = gspread.authorize(creds)
    print("[✅] Google Sheets connected.")
except Exception as e:
    print("[❌] Failed to connect to Sheets:", e)

@app.route("/")
def home():
    return "🎯 Tracking Pixel Backend Live"

@app.route("/track")
def track_open():
    sheet_name = request.args.get("sheet")
    row_number = request.args.get("row")
    email_param = request.args.get("email")
    t = request.args.get("t")  # timestamp sent

    print(f"[👀] Tracking pixel hit received: sheet={sheet_name}, row={row_number}, email={email_param}, time={t}")
    
    # Delay protection — ignore hits within 10 seconds of sending (e.g., Gmail proxy)
    try:
        t = int(t)
        if time.time() - t < 10:
            print("⏱️ Skipping hit — too soon (proxy preload).")
            return Response(status=204)
    except Exception as e:
        print("⚠️ Invalid timestamp:", e)
        return Response(status=204)

    if not sheet_name or not row_number or not email_param:
        print("❌ Missing required query params.")
        return Response(status=204)

    try:
        sheet = gc.open_by_key(os.environ.get("SHEET_ID")).worksheet(sheet_name)
        all_values = sheet.get_all_values()

        header = all_values[0]
        email_col = header.index("Email ID")
        open_col = header.index("Open?")
        open_time_col = header.index("Open Timestamp")

        row_index = int(row_number)

        row_values = all_values[row_index - 1]  # 1-based

        current_email = row_values[email_col].strip().lower()
        if current_email != email_param.strip().lower():
            print(f"⚠️ Email mismatch. Expected: {current_email}, Got: {email_param}")
            return Response(status=204)

        if row_values[open_col].strip().lower() != "yes":
            sheet.update_cell(row_index, open_col + 1, "Yes")
            sheet.update_cell(row_index, open_time_col + 1, time.strftime("%d/%m/%Y %H:%M:%S"))
            print(f"[✅] Marked opened for row {row_index}, email={email_param}")
        else:
            print(f"[ℹ️] Already marked open for row {row_index}")

    except Exception as e:
        print("❌ Error updating sheet:", e)

    # Transparent 1x1 pixel
    pixel_data = b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xFF\xFF\xFF\x21' \
                 b'\xF9\x04\x01\x00\x00\x00\x00\x2C\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x4C\x01\x00\x3B'
    return Response(pixel_data, mimetype='image/gif')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
