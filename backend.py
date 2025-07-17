from flask import Flask, request, Response
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import pytz

app = Flask(__name__)
TIMEZONE = pytz.timezone("Asia/Kolkata")

# Load credentials
creds_dict = json.loads(os.environ["GOOGLE_JSON"])
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(credentials)

@app.route("/track")
def track():
    try:
        sheet_name = request.args.get("sheet")
        row = int(request.args.get("row"))
        email_param = request.args.get("email")

        if not sheet_name or not row or not email_param:
            return Response(status=204)

        user_agent = request.headers.get("User-Agent", "")
        if "google" in user_agent.lower():
            return Response(status=204)  # Ignore proxy hits

        sheet = client.open_by_key("1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q").worksheet(sheet_name)
        data = sheet.get_all_values()
        headers = data[0]
        idx = lambda col: headers.index(col)

        current_email = data[row - 1][idx("Email ID")].strip()
        if current_email.lower() != email_param.lower():
            return Response(status=204)  # Email mismatch

        sheet.update_cell(row, idx("Open?") + 1, "Yes")
        sheet.update_cell(row, idx("Open Timestamp") + 1, datetime.now(TIMEZONE).strftime("%d-%m-%Y %H:%M:%S"))
        print(f"✅ Open tracked: Row {row} - {email_param}")
        return Response(status=204)

    except Exception as e:
        print(f"❌ Tracking error: {e}")
        return Response(status=204)

@app.route("/")
def home():
    return "📬 Tracker Running", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
