import csv
import io
from datetime import datetime
from openpyxl import Workbook

def convert_tsv_to_excel(tsv_text):
    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"

    headers = ["Date(Weekday)", "Time", "Location", "Name", "Email", "NetID", "Non-UA Email"]
    ws.append(headers)

    reader = csv.DictReader(io.StringIO(tsv_text), delimiter="\t")

    for row in reader:
        raw_dt = row.get("Date Time", "").strip()
        if not raw_dt:
            continue

        dt_obj = datetime.strptime(raw_dt, "%m/%d/%Y %I:%M %p")
        weekday = dt_obj.strftime("%A")
        formatted_time = f"@{dt_obj.strftime('%I:%M%p').lstrip('0')}"

        location = row.get("Location", "")
        if "(BASC)" in location:
            location = location.split("(BASC)")[0] + "(BASC)"

        name = row.get("Customer Name", "")
        email = row.get("Customer Email", "").strip()
        net_id = email.split("@")[0] if "@" in email else email
        non_ua_email = "TRUE" if email and not email.lower().endswith("@arizona.edu") else ""

        ws.append([weekday, formatted_time, location, name, email, net_id, non_ua_email])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()