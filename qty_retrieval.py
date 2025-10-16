from simple_salesforce import Salesforce
import pandas as pd
import xlwings as xw
import re
import os
import time

# --- Salesforce login ---
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# --- Load Excel workbook & table ---
wb = xw.Book(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx")
sheet_out = wb.sheets['test']
table_out = sheet_out.tables['test']

# --- Helper: parse case number safely ---
def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None

# --- Helper: extract *all* quantities (supports multiple pcs/qty mentions) ---
def extract_quantity(title):
    if not title:
        return None
    text = title.lower()

    # Find all patterns like:
    # qty 5, qty:5, 5 pcs, 5pcs, etc.
    matches = re.findall(r'(?:qty\s*[:\-]?\s*(\d+))|(?:(\d+)\s*pcs?)', text)

    total_qty = 0
    for m in matches:
        num = next((int(x) for x in m if x and x.isdigit()), None)
        if num:
            total_qty += num

    return total_qty if total_qty > 0 else None

try:
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')

        # --- Get existing case numbers from Excel ---
        output_df = pd.DataFrame(table_out.range.value[1:], columns=table_out.range.value[0])
        parsed_cases = output_df['Case Number'].apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # --- Pull Salesforce report ---
        report_id = "00OUI00000EsGR72AN"
        report_data = sf.restful(f'analytics/reports/{report_id}', params={'includeDetails': 'true'})
        report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])

        # --- Extract new case numbers ---
        new_case_numbers = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) > 3:
                case_number = cells[3].get('label', '').strip()
                if case_number and re.match(r'^\d+$', case_number):
                    if parse_case_number(case_number) not in existing_cases:
                        new_case_numbers.append(case_number)

        print(f"🔍 Found {len(new_case_numbers)} new cases not in Excel.\n")

        # --- Fetch case titles and extract quantities ---
        for num in new_case_numbers:
            soql = f"SELECT Id, CaseNumber, Subject FROM Case WHERE CaseNumber = '{num}'"
            result = sf.query(soql)
            records = result.get('records', [])
            if records:
                case = records[0]
                title = case['Subject'] or ''
                qty = extract_quantity(title)
                print(f"Case {case['CaseNumber']}: {title} | Qty: {qty if qty is not None else 'N/A'}")
            else:
                print(f"⚠️ Case {num} not found via SOQL.")

        print("\nWaiting 60s before next cycle...\n")
        time.sleep(60)

except KeyboardInterrupt:
    print("\nStopped by user.")
