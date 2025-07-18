from flask import Flask, request, Response
import gspread
import os
import datetime
import json

app = Flask(__name__)

@app.route('/')
def index():
    return '✅ Email Tracking Backend is Live', 200

@app.route('/track')
def track():
    sheet_name = request.args.get('sheet')
    row = request.args.get('row')
    email_param = request.args.get('email')

    if not all([sheet_name, row, email_param]):
        print("❌ Missing parameters in tracking URL")
        return Response(status=204)

    try:
        gc = gspread.service_account_from_dict(json.loads(os.environ['GOOGLE_JSON']))
        sh = gc.open_by_key(os.environ['SHEET_ID'])
        worksheet = sh.worksheet(sheet_name)

        row = int(row)

        # Read headers to find "Email ID", "Open?", and "Open Timestamp"
        headers = worksheet.row_values(1)
        email_col = headers.index("Email ID") + 1
        open_col = headers.index("Open?") + 1
        timestamp_col = headers.index("Open Timestamp") + 1

        # Get email in sheet
        email_in_sheet = worksheet.cell(row, email_col).value
        if not email_in_sheet:
            print(f"❌ No email found in row {row}")
            return Response(status=204)

        email_in_sheet = email_in_sheet.strip().lower()
        if email_in_sheet != email_param.strip().lower():
            print(f"❌ Email mismatch: sheet='{email_in_sheet}' vs pixel='{email_param}'")
            return Response(status=204)

        # Check if already opened
        open_status = worksheet.cell(row, open_col).value
        if open_status and open_status.strip().lower() == "yes":
            print(f"🔁 Row {row} already marked as opened")
            return Response(status=204)

        # Update "Open?" and "Open Timestamp"
        worksheet.update_cell(row, open_col, "Yes")
        worksheet.update_cell(row, timestamp_col, datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        print(f"✅ Marked opened for row {row}, email={email_param}")
        return Response(status=204)

    except Exception as e:
        print(f"❌ Exception: {str(e)}")
        return Response(status=204)
