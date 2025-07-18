from flask import Flask, request, make_response
import datetime
import gspread
import os
import json
import pytz

app = Flask(__name__)

@app.route('/track', methods=['GET'])
def track_email():
    try:
        sheet_name = request.args.get('sheet')
        row_number = int(request.args.get('row'))
        email_param = request.args.get('email')

        if not sheet_name or not row_number or not email_param:
            return make_response("Missing required parameters", 400)

        # Load credentials
        creds = json.loads(os.environ['GOOGLE_JSON'])
        gc = gspread.service_account_from_dict(creds)
        sh = gc.open_by_key(os.environ['SHEET_ID'])
        worksheet = sh.worksheet(sheet_name)

        values = worksheet.row_values(row_number)

        if len(values) < 9:
            return make_response("Row data incomplete", 204)

        current_open = values[8].strip() if len(values) >= 9 else ''
        current_email = values[2].strip() if len(values) >= 3 else ''

        # Check if already opened
        if current_open == 'Yes':
            return pixel_response()

        # Ensure email matches
        if current_email.lower() != email_param.lower():
            return make_response("Email mismatch", 204)

        # Check for proxy User-Agent (Gmail image proxy)
        ua = request.headers.get('User-Agent', '').lower()
        if 'googleimageproxy' in ua or 'google' in request.remote_addr:
            return make_response("Ignored proxy request", 204)

        # Optionally: Check if opened too soon (e.g., < 60 seconds after sent)
        timestamp_str = values[7].strip()  # Timestamp (when mail was sent)
        if timestamp_str:
            try:
                tz = pytz.timezone('Asia/Kolkata')
                sent_time = datetime.datetime.strptime(timestamp_str, '%d/%m/%Y %H:%M:%S')
                sent_time = tz.localize(sent_time)
                now = datetime.datetime.now(tz)
                if (now - sent_time).total_seconds() < 60:
                    return make_response("Ignored: opened too soon", 204)
            except Exception as e:
                pass  # Continue even if timestamp can't be parsed

        # Update 'Open?' and 'Open Timestamp'
        worksheet.update_cell(row_number, 9, 'Yes')  # Open?
        open_time = datetime.datetime.now(pytz.timezone('Asia/Kolkata')).strftime('%d/%m/%Y %H:%M:%S')
        worksheet.update_cell(row_number, 10, open_time)  # Open Timestamp

        return pixel_response()

    except Exception as e:
        print(f"Error in /track: {str(e)}")
        return make_response("Error", 204)

def pixel_response():
    response = make_response('', 204)
    response.headers['Content-Type'] = 'image/gif'
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

if __name__ == '__main__':
    app.run()
