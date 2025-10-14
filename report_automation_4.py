import time
import csv
from simple_salesforce import Salesforce
import json
from tabulate import tabulate
import pandas as pd
import xlwings as xw
import re
from rapidfuzz import process

# ----- Salesforce login -----
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# --- Fetch Salesforce report ---
report_id = "00OUI00000EsGR72AN"
report_data = sf.restful(f'analytics/reports/{report_id}', params={'includeDetails': 'true'})

# --- Limit for testing — first 5 rows ---
report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])[:5]

# --- Load supplier table from Excel ---
wb = xw.Book(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx")
sheet_sup = wb.sheets['Class Site Supplier']
table_range = sheet_sup.tables['Table3'].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])

# --- Clean and align lists ---
supplier_df = supplier_df.dropna(subset=['Stock Code'])
stock_codes_list = supplier_df['Stock Code'].astype(str).tolist()
man_sites_list   = supplier_df['Mfg Plant Name'].fillna("").astype(str).tolist()
suppliers_list   = supplier_df['Supplier'].fillna("").astype(str).tolist()

# --- Normalize helper ---
def normalize(text):
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# --- Get reference to output table ---
sheet_out = wb.sheets['test']
table_out = sheet_out.tables['test']
output_df = pd.DataFrame(table_out.range.value[1:], columns=table_out.range.value[0])

# --- Robust case number parser ---
def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None

# --- Track existing case numbers ---
parsed_cases = output_df['Case Number'].apply(parse_case_number).tolist()
existing_cases = set(filter(None, parsed_cases))

# --- Process each Salesforce row ---
new_rows_to_add = []

for row in report_rows:
    cells = row.get('dataCells', [])

    # Convert Salesforce Case Number to integer robustly
    case_number = parse_case_number(cells[3].get('label', ''))
    if case_number is None or case_number in existing_cases:
        continue  # skip if invalid or already exists

    description = str(cells[4].get('label', ''))

    # --- Match to supplier/manufacturer ---
    normalized_description = normalize(description)
    normalized_stock_codes = [normalize(code) for code in stock_codes_list]
    matches = process.extract(normalized_description, normalized_stock_codes, limit=1)

    most_common_source = "N/A"
    if matches:
        best_match = matches[0][0]
        try:
            match_index = normalized_stock_codes.index(best_match)
            supplier = suppliers_list[match_index].strip() if match_index < len(suppliers_list) else ""
            most_common_source = supplier if supplier else man_sites_list[match_index]
        except Exception:
            pass

    # --- Prepare new row ---
    new_row = [
        cells[0].get('label', ''),  # Opened Date
        cells[1].get('label', ''),  # Case Reason
        cells[2].get('label', ''),  # Case Owner
        case_number,                # Case Number (int)
        description,                # Description
        '',                         # Quantity
        cells[5].get('label', ''),  # RMA Value
        cells[6].get('label', ''),  # Case Category
        cells[7].get('label', ''),  # Account Name
        '',                         # Comments
        cells[8].get('label', ''),  # Contact Type
        cells[9].get('label', ''),  # Shipping Whse
        most_common_source          # Source
    ]

    new_rows_to_add.append(new_row)
    existing_cases.add(case_number)

# --- Append all new rows using the working method ---
if new_rows_to_add:
    current_table_rows = table_out.range.rows.count
    current_table_cols = table_out.range.columns.count

    # Write new rows below the table
    start_row = table_out.range.row + current_table_rows
    sheet_out.range((start_row, table_out.range.column)).value = new_rows_to_add

    # Resize table to include the new rows
    table_out.resize(sheet_out.range((table_out.range.row, table_out.range.column),
                                     (table_out.range.row + current_table_rows + len(new_rows_to_add) - 1,
                                      table_out.range.column + current_table_cols - 1)))

# --- Output arrays for inspection ---
rows_added_array = [row[3] for row in new_rows_to_add]  # Only Case Numbers
existing_cases_array = list(existing_cases)

print("New rows to add (rows_added_array):", rows_added_array)
print("All case numbers including existing (existing_cases_array):", existing_cases_array)

