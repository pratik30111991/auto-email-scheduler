from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
from PIL import Image
import os

app = Flask(__name__)

# Use environment variable for Google credentials
GOOGLE_CREDS_JSON = os.getenv('GOOGLE_JSON')

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Load your Sheet ID from env or hardcode it
SHEET_ID = "1J7bS1MfkLh5hXnpBfHdx-uYU7Qf9gc965CdW-j9mf2Q"

@app.route("/track")
def track():
    sheet_name = request.args.get("sheet")
    row = request.args.get("row")
    email_param = request.args.get("email")

    if not sheet_name or not row or not email_param:
        return "Missing parameters", 400

    try:
        row = int(row)
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)

        email_col = 3  # Column C (Email ID)
        open_col = 10  # Column J (Open?)
        open_time_col = 11  # Column K (Open Timestamp)

        row_data = sheet.row_values(row)
        actual_email = row_data[email_col - 1].strip().lower()

        if actual_email != email_param.strip().lower():
            return "Invalid email for row", 403

        open_status = sheet.cell(row, open_col).value
        if open_status != "Yes":
            sheet.update_cell(row, open_col, "Yes")
            now_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            sheet.update_cell(row, open_time_col, now_time)

    except Exception as e:
        print(f"Error: {e}")
        return "Error processing tracking", 500

    # Transparent 1x1 GIF
    image = Image.new("RGBA", (1, 1), (255, 255, 255, 0))
    byte_io = io.BytesIO()
    image.save(byte_io, "PNG")
    byte_io.seek(0)
    return send_file(byte_io, mimetype='image/png')

@app.route("/")
def home():
    return "Pixel Tracker is live"

if __name__ == "__main__":
    app.run(debug=True)
