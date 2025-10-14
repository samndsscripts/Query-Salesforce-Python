import time
import csv
from simple_salesforce import Salesforce
import json
from tabulate import tabulate
from openpyxl import load_workbook
import pandas as pd
import xlwings as xw
import re
from collections import Counter
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
output_range = table_out.range
output_df = pd.DataFrame(output_range.value[1:], columns=output_range.value[0])

# --- Track existing case numbers as integers ---
existing_cases = set(
    output_df['Case Number'].dropna().apply(lambda x: int(float(str(x).strip()))).tolist()
)

# --- Process each Salesforce row ---
new_rows_to_add = []

for row in report_rows:
    cells = row.get('dataCells', [])

    # Convert Salesforce Case Number to integer robustly
    case_number_str = str(cells[3].get('label', '')).strip()
    if not case_number_str.isdigit():
        continue
    case_number = int(case_number_str)

    # Skip if already in table
    if case_number in existing_cases:
        continue

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

    # --- Append only new rows ---
    new_rows_to_add.append(new_row)
    existing_cases.add(case_number)

# --- Append all new rows at once to Excel ---
if new_rows_to_add:
    first_empty_row = table_out.range.last_cell.row + 1
    sheet_out.range(f"A{first_empty_row}").value = new_rows_to_add
    table_out.resize(table_out.range.expand('table'))

# --- Output arrays for inspection ---
rows_added_array = [row[3] for row in new_rows_to_add]  # Only Case Numbers
existing_cases_array = list(existing_cases)

print("New rows to add (rows_added_array):", rows_added_array)
print("All case numbers including existing (existing_cases_array):", existing_cases_array)
