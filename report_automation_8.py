import time
import re
import os
import pandas as pd
import xlwings as xw
from rapidfuzz import process
from simple_salesforce import Salesforce

# --- Salesforce login ---
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# --- Workbook setup ---
WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"
wb = xw.Book(WB_PATH)

sheet_sup = wb.sheets['Class Site Supplier']
sheet_out = wb.sheets['test']
table_out = sheet_out.tables['test']

# --- Load supplier data ---
table_range = sheet_sup.tables['Table3'].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])
supplier_df = supplier_df.fillna('')
stock_codes_list = supplier_df['Stock Code'].astype(str).tolist()
man_sites_list   = supplier_df['Mfg Plant Name'].astype(str).tolist()
suppliers_list   = supplier_df['Supplier'].astype(str).tolist()
normalized_stock_codes = [re.sub(r"[^a-z0-9\s]", " ", s.lower()).strip() for s in stock_codes_list]

# --- Helpers ---
def clear_console():
    os.system('cls' if os.name == 'nt' else 'clear')

def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None

def normalize(text):
    t = str(text or '').lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def extract_quantity(title):
    if not title:
        return None
    text = title.lower()
    matches = re.findall(r'(?:qty\s*[:\-]?\s*(\d+))|(?:(\d+)\s*pcs?)', text)
    total_qty = 0
    for m in matches:
        num = next((int(x) for x in m if x and x.isdigit()), None)
        if num:
            total_qty += num
    return total_qty if total_qty > 0 else None

def determine_source(description):
    nd = normalize(description)
    if not nd:
        return "N/A"
    match = process.extractOne(nd, normalized_stock_codes)
    if not match:
        return "N/A"
    best_norm_code = match[0]
    try:
        idx = normalized_stock_codes.index(best_norm_code)
        supplier = suppliers_list[idx].strip()
        if supplier:
            return supplier
        else:
            return man_sites_list[idx] if idx < len(man_sites_list) else "N/A"
    except Exception:
        return "N/A"

# --- Main Loop ---
REPORT_ID = "00OUI00000EsGR72AN"
CYCLE_SECONDS = 60
SOQL_BATCH = 100

try:
    while True:
        clear_console()
        print("🔄 Checking for new Salesforce cases...\n")

        # --- Load existing Excel data ---
        table_vals = table_out.range.value
        if table_vals and len(table_vals) > 1:
            output_df = pd.DataFrame(table_vals[1:], columns=table_vals[0])
        else:
            output_df = pd.DataFrame(columns=table_vals[0] if table_vals else [])
        parsed_cases = output_df.get('Case Number', pd.Series([])).apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # --- Pull Salesforce report ---
        try:
            report_data = sf.restful(f'analytics/reports/{REPORT_ID}', params={'includeDetails': 'true'})
            report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])
        except Exception as e:
            print("❌ Error fetching report:", e)
            time.sleep(CYCLE_SECONDS)
            continue

        # --- Collect new case numbers ---
        new_case_numbers = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) > 3:
                cn_raw = str(cells[3].get('label', '')).strip()
                cn_int = parse_case_number(cn_raw)
                if cn_int and cn_int not in existing_cases:
                    new_case_numbers.append(cn_raw)

        print(f"Found {len(new_case_numbers)} new case(s).")

        if not new_case_numbers:
            print(f"No new cases. Sleeping {CYCLE_SECONDS}s...\n")
            time.sleep(CYCLE_SECONDS)
            continue

        # --- Batch SOQL: get titles ---
        case_subject_map = {}
        for batch in [new_case_numbers[i:i+SOQL_BATCH] for i in range(0, len(new_case_numbers), SOQL_BATCH)]:
            quoted = ",".join(f"'{cn}'" for cn in batch)
            soql = f"SELECT CaseNumber, Subject FROM Case WHERE CaseNumber IN ({quoted})"
            try:
                result = sf.query_all(soql)
                for rec in result.get('records', []):
                    case_subject_map[rec['CaseNumber']] = rec.get('Subject', '')
            except Exception as e:
                print("⚠️ SOQL batch failed:", e)

        # --- Build new rows ---
        new_rows_to_add = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) < 10:
                continue
            cn_raw = str(cells[3].get('label', '')).strip()
            cn_int = parse_case_number(cn_raw)
            if not cn_int or cn_int in existing_cases:
                continue
            if cn_raw not in new_case_numbers:
                continue

            description = str(cells[4].get('label', ''))
            subject = case_subject_map.get(cn_raw, '')
            qty = extract_quantity(subject)
            source = determine_source(description)

            new_row = [
                cells[0].get('label', ''),  # Opened Date
                cells[1].get('label', ''),  # Case Reason
                cells[2].get('label', ''),  # Case Owner
                cn_int,                     # Case Number
                description,                # Description
                qty if qty is not None else '',  # Quantity
                cells[5].get('label', ''),  # RMA Value
                cells[6].get('label', ''),  # Case Category
                cells[7].get('label', ''),  # Account Name
                '',                         # Comments
                cells[8].get('label', ''),  # Contact Type
                cells[9].get('label', ''),  # Shipping Whse
                source                      # Source
            ]
            new_rows_to_add.append(new_row)
            existing_cases.add(cn_int)

        # --- Append into the Excel Table properly ---
        if new_rows_to_add:
            # Expand the table range
            table_out.range.end('down').offset(1, 0).value = new_rows_to_add
            print(f"\n✅ Added {len(new_rows_to_add)} new row(s) into the table.\n")
            for nr in new_rows_to_add:
                print(f" → Case {nr[3]} | Qty: {nr[5] or 'N/A'} | Source: {nr[-1]}")
        else:
            print("No valid new rows to add this cycle.")

        print(f"\nSleeping {CYCLE_SECONDS}s before next check...\n")
        time.sleep(CYCLE_SECONDS)

except KeyboardInterrupt:
    print("\nScript stopped by user.")
