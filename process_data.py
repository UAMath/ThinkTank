import csv
import io
from datetime import datetime
from collections import Counter
from openpyxl import Workbook

def make_first_last_key(name_str):
    """ Extracts only the VERY FIRST word and the VERY LAST word of a name string. """
    if not name_str:
        return ""
    
    cleaned = name_str.strip().lower()
    
    if "," in cleaned:
        parts = [p.strip() for p in cleaned.split(",")]
        last_part = parts[0].split()[0] if parts[0] else ""
        first_part = parts[1].split()[0] if len(parts) > 1 and parts[1] else ""
        return f"{first_part} {last_part}".strip()
    
    words = cleaned.split()
    if len(words) == 1:
        return words[0]
    
    first_word = words[0]
    last_word = words[-1]
    return f"{first_word} {last_word}"

def convert_tsv_to_excel(tsv_text):
    # 1. Load Master Sheet into lookup dictionary
    email_lookup = {}
    
    # Dictionaries to count individual first/last name occurrences
    first_name_counts = Counter()
    last_name_counts = Counter()
    first_name_to_email = {}
    last_name_to_email = {}

    try:
        with open('master_sheet.csv', mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                first_name = row.get('First Name', '').strip().lower()
                last_name = row.get('Last Name', '').strip().lower()
                email = row.get('Email', '').strip()

                if not email:
                    continue

                raw_name = f"{first_name} {last_name}"
                normalized_name = " ".join(raw_name.split())
                
                if normalized_name:
                    email_lookup[normalized_name] = email
                    short_key = make_first_last_key(normalized_name)
                    if short_key:
                        email_lookup[short_key] = email

                if first_name:
                    first_name_counts[first_name] += 1
                    first_name_to_email[first_name] = email

                if last_name:
                    last_name_counts[last_name] += 1
                    last_name_to_email[last_name] = email

        for fn, count in first_name_counts.items():
            if count == 1 and fn not in email_lookup:
                email_lookup[fn] = first_name_to_email[fn]

        for ln, count in last_name_counts.items():
            if count == 1 and ln not in email_lookup:
                email_lookup[ln] = last_name_to_email[ln]

    except Exception as e:
        print(f"Warning: Could not read master_sheet.csv: {e}")

    # 2. Set up Excel Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"

    headers = ["Date(Weekday)", "Time", "Location", "Name", "Email", "NetID", "Non-UA Email"]
    ws.append(headers)

    # 3. Read incoming TSV input
    tsv_reader = csv.DictReader(io.StringIO(tsv_text), delimiter="\t")

    for row in tsv_reader:
        raw_dt = row.get("Date Time", "").strip()
        if not raw_dt:
            continue

        dt_obj = datetime.strptime(raw_dt, "%m/%d/%Y %I:%M %p")
        
        # Shortened weekday and month date format (e.g., "Mon_Sep 28")
        formatted_date = dt_obj.strftime("%a_%b %d")

        # Time formatted with surrounding spaces (e.g., " @ 9:30AM ")
        time_str = dt_obj.strftime("%I:%M%p").lstrip('0')
        formatted_time = f" @ {time_str} "

        # Hardcoded fixed location as requested
        formatted_location = "in Math Room 101"

        name = row.get("Customer Name", "").strip()
        email = row.get("Customer Email", "").strip()

        # Check if email is non-UA and lookup replacement in master sheet
        if email and not email.lower().endswith("@arizona.edu"):
            # Normalize incoming search key (lowercase & single spaces)
            input_key = " ".join(name.lower().split())
            found_email = email_lookup.get(input_key)
            
            if not found_email:
                short_input_key = make_first_last_key(input_key)
                found_email = email_lookup.get(short_input_key)
            
            if found_email:
                email = found_email

        # Compute NetID and Non-UA status
        net_id = email.split("@")[0] if "@" in email else email
        non_ua_email = "TRUE" if email and not email.lower().endswith("@arizona.edu") else ""

        ws.append([formatted_date, formatted_time, formatted_location, name, email, net_id, non_ua_email])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()