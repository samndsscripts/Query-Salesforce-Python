import time
import csv
from simple_salesforce import Salesforce
import json
from tabulate import tabulate
import pandas as pd
import xlwings as xw
import re
from rapidfuzz import process
import os  # for clearing the screen

# ----- Salesforce login -----
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

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

# --- Robust case number parser ---
def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None

# --- Quantity parser ---
def parse_quantity_from_title(title):
    """
    Extracts an integer quantity based on 'qty' or 'pcs' patterns.
    Rules:
      - If 'qty' is found → number comes AFTER it.
      - If 'pcs' is found → number comes BEFORE it.
    Returns '' if unparsable.
    """
    text = str(title).lower()

    # QTY pattern (number after)
    qty_match = re.search(r'qty\s*(\d+)', text)
    if qty_match:
        return int(qty_match.group(1))

    # PCS pattern (number before)
    pcs_match = re.search(r'(\d+)\s*pcs', text)
    if pcs_match:
        return int(pcs_match.group(1))

    return ''  # no quantity found


# --- Main loop (Version 6.1) ---
try:
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')  # clear console each cycle

        # --- Get reference to output table ---
        sheet_out = wb.sheets['test']
        table_out = sheet_out.tables['test']
        output_df = pd.DataFrame(table_out.range.value[1:], columns=table_out.range.value[0])

        # --- Track existing case numbers ---
        parsed_cases = output_df['Case Number'].apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # --- Fetch Salesforce report ---
        report_id = "00OUI00000EsGR72AN"
        report_data = sf.restful(f'analytics/reports/{report_id}', params={'includeDetails': 'true'})
        report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])

        # --- Process each Salesforce row ---
        new_rows_to_add = []
        for row in report_rows:
            cells = row.get('dataCells', [])
            case_number = parse_case_number(cells[3].get('label', ''))
            if case_number is None or case_number in existing_cases:
                continue  # skip duplicates or invalid

            # --- Get Case Title (Subject) via SOQL ---
            soql = f"SELECT Id, CaseNumber, Subject FROM Case WHERE CaseNumber = '{case_number}'"
            result = sf.query(soql)
            records = result.get('records', [])
            case_title = ""
            if records:
                case_title = records[0].get('Subject', '')
            else:
                case_title = cells[4].get('label', '')  # fallback if missing

            # --- Extract quantity from title ---
            quantity_value = parse_quantity_from_title(case_title)

            # --- Match to supplier/manufacturer ---
            normalized_description = normalize(case_title)
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
                case_number,                # Case Number
                case_title,                 # Case Title
                quantity_value,             # Quantity (parsed)
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

        # --- Append new rows using working method ---
        if new_rows_to_add:
            current_table_rows = table_out.range.rows.count
            current_table_cols = table_out.range.columns.count
            start_row = table_out.range.row + current_table_rows
            sheet_out.range((start_row, table_out.range.column)).value = new_rows_to_add
            table_out.resize(sheet_out.range(
                (table_out.range.row, table_out.range.column),
                (table_out.range.row + current_table_rows + len(new_rows_to_add) - 1,
                 table_out.range.column + current_table_cols - 1)
            ))

        # --- Print new rows added this cycle ---
        print(f"Cycle complete. New rows added: {len(new_rows_to_add)}")
        if new_rows_to_add:
            for row in new_rows_to_add:
                print(f"Case {row[3]} | Qty: {row[5]} | {row[4]}")

        # --- Sleep before next cycle ---
        time.sleep(60)

except KeyboardInterrupt:
    print("\nScript stopped by user.")
