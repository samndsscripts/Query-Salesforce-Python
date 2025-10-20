# version6_continuous_fixed.py
import time
import re
import os
from collections import defaultdict

import pandas as pd
import xlwings as xw
from simple_salesforce import Salesforce
from rapidfuzz import process
from colorama import init as colorama_init, Fore, Style

# initialize colorama for colored terminal output
colorama_init(autoreset=True)

# ----- Salesforce login (use your credentials) -----
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# ----- Excel workbook / supplier table setup -----
WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"
SUPPLIER_SHEET = 'Class Site Supplier'
SUPPLIER_TABLE = 'Table3'
OUTPUT_SHEET = 'test'
OUTPUT_TABLE = 'test'

wb = xw.Book(WB_PATH)
sheet_sup = wb.sheets[SUPPLIER_SHEET]
table_range = sheet_sup.tables[SUPPLIER_TABLE].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])

# ensure supplier_df aligned
supplier_df = supplier_df.fillna('')

stock_codes_list = supplier_df['Stock Code'].astype(str).tolist()
man_sites_list = supplier_df['Mfg Plant Name'].astype(str).tolist()
suppliers_list = supplier_df['Supplier'].astype(str).tolist()

# --- Helpers ---
def normalize(text: str) -> str:
    t = str(text or "").lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

# robust parse for case numbers (works with "141331.00" or "00141331")
def parse_case_number(x):
    try:
        xs = str(x).strip()
        if xs == "":
            return None
        digits = re.findall(r"\d+", xs)
        if not digits:
            return None
        return int("".join(digits))
    except Exception:
        return None

# extract and sum all qty/pcs mentions from a title
def extract_total_quantity(title: str):
    if not title:
        return None
    text = str(title).lower()
    qty_after_matches = re.findall(r'qty\s*[:\-]?\s*(\d+)', text)
    pcs_before_matches = re.findall(r'(\d+)\s*pcs?', text)
    total = 0
    for m in qty_after_matches:
        try:
            total += int(m)
        except:
            pass
    for m in pcs_before_matches:
        try:
            total += int(m)
        except:
            pass
    return int(total) if total > 0 else None

# batch a list into chunks of size n
def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

# Find Source via fuzzy match (single best match)
def determine_source(description, normalized_stock_codes):
    nd = normalize(description)
    if not nd or not normalized_stock_codes:
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

# --- Main continuous loop (Version 6) ---
REPORT_ID = "00OUI00000EsGR72AN"
CYCLE_SECONDS = 60
SOQL_BATCH = 100  # number of case numbers per SOQL IN() (safe)

try:
    while True:
        # clear console
        os.system('cls' if os.name == 'nt' else 'clear')

        # get reference to output table and sheet
        sheet_out = wb.sheets[OUTPUT_SHEET]
        table_out = sheet_out.tables[OUTPUT_TABLE]

        # read current table data (handle empty table gracefully)
        try:
            table_vals = table_out.range.value or []
            if isinstance(table_vals, list) and len(table_vals) >= 1:
                output_df = pd.DataFrame(table_vals[1:], columns=table_vals[0])
            else:
                # header-only table or weird shape -> empty df with expected columns if available
                cols = table_vals[0] if isinstance(table_vals, list) and table_vals else []
                output_df = pd.DataFrame(columns=cols)
        except Exception:
            output_df = pd.DataFrame()

        # build set of existing case numbers (integers)
        if 'Case Number' in output_df.columns:
            parsed_cases = output_df['Case Number'].apply(parse_case_number).tolist()
            existing_cases = set(filter(None, parsed_cases))
        else:
            existing_cases = set()

        # fetch report rows
        try:
            report_data = sf.restful(f'analytics/reports/{REPORT_ID}', params={'includeDetails': 'true'})
            report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', []) or []
        except Exception as e:
            print(Fore.RED + "❌ Error fetching report:", str(e))
            time.sleep(CYCLE_SECONDS)
            continue

        # collect unique new case numbers (strings), keeping order
        new_case_numbers = []
        seen_cn = set()
        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) <= 3:
                continue
            raw_cn = str(cells[3].get('label', '')).strip()
            if not raw_cn:
                continue
            cn_int = parse_case_number(raw_cn)
            if cn_int is None or cn_int in existing_cases:
                continue
            if raw_cn not in seen_cn:
                seen_cn.add(raw_cn)
                new_case_numbers.append(raw_cn)

        # quick status
        print(f"Cycle start: Found {len(new_case_numbers)} new case(s) not in Excel table.")

        # if none, show summary and sleep
        if not new_case_numbers:
            print("\nNo new cases this cycle.")
            print(f"Waiting {CYCLE_SECONDS}s before next cycle...")
            time.sleep(CYCLE_SECONDS)
            continue

        # Batch SOQL queries to get Case.Subject for the new cases
        case_subject_map = {}
        for batch in chunks(new_case_numbers, SOQL_BATCH):
            quoted = ",".join(f"'{cn}'" for cn in batch)
            soql = f"SELECT CaseNumber, Subject FROM Case WHERE CaseNumber IN ({quoted})"
            try:
                res = sf.query_all(soql)
                recs = res.get('records', []) if res else []
                for r in recs:
                    case_subject_map[r['CaseNumber']] = r.get('Subject') or ''
            except Exception as e:
                print(Fore.RED + "SOQL error:", str(e))

        # prepare normalized stock codes once for this cycle
        normalized_stock_codes = [normalize(c) for c in stock_codes_list]

        # build new rows to append
        new_rows_to_add = []
        printed_rows = []  # for printing after all prep
        for raw_cn in new_case_numbers:
            cn_int = parse_case_number(raw_cn)
            if cn_int is None or cn_int in existing_cases:
                continue

            # find the matching report row to get other fields (Opened Date etc.)
            report_row = None
            for r in report_rows:
                cells = r.get('dataCells', [])
                if len(cells) > 3 and str(cells[3].get('label', '')).strip() == raw_cn:
                    report_row = r
                    break
            if report_row is None:
                continue

            cells = report_row.get('dataCells', [])
            description = str(cells[4].get('label', '') or "")
            subject = case_subject_map.get(raw_cn, "")

            # parse quantity from subject (supports multiple mentions)
            qty_val = None
            if subject:
                qty_val = extract_total_quantity(subject)

            # determine source by fuzzy match on description
            source = determine_source(description, normalized_stock_codes)

            # prepare new row using same column order you used earlier
            new_row = [
                cells[0].get('label', ''),  # Opened Date
                cells[1].get('label', ''),  # Case Reason
                cells[2].get('label', ''),  # Case Owner
                cn_int,                     # Case Number (int)
                description,                # Description
                '',
                qty_val if qty_val is not None else '',  # Quantity
                cells[5].get('label', ''),  # RMA Value
                cells[6].get('label', ''),  # Case Category
                cells[7].get('label', ''),  # Account Name
                cells[8].get('label', ''),  # Contact Type
                cells[9].get('label', ''),  # Shipping Whse
                source                      # Source
            ]

            new_rows_to_add.append(new_row)
            existing_cases.add(cn_int)
            printed_rows.append((cn_int, qty_val, subject))

        # Append new rows into the table: write below current table and rely on Excel auto-expand
        if new_rows_to_add:
            current_table_rows = table_out.range.rows.count
            current_table_cols = table_out.range.columns.count
            start_row = table_out.range.row + current_table_rows
            # write new rows immediately below table
            sheet_out.range((start_row, table_out.range.column)).value = new_rows_to_add

            # Refresh the table object reference and print new count for confirmation
            try:
                # Re-fetch table object (some Excel states require re-binding)
                table_out = sheet_out.tables[OUTPUT_TABLE]
                new_count = table_out.range.rows.count
                print(Fore.CYAN + f"\nTable now has {new_count} rows (including header).")
            except Exception:
                print(Fore.YELLOW + "Unable to refresh table object; please check Excel.")

        # Print results for this cycle (colored)
        print()
        if printed_rows:
            for cn_int, qty_val, subject in printed_rows:
                if qty_val is not None:
                    print(Fore.GREEN + f"New → Case {cn_int} | Qty: {qty_val} | {subject}")
                else:
                    print(Fore.YELLOW + f"New → Case {cn_int} | Qty: N/A | {subject}")
        else:
            print("No rows were actually prepared to add (race condition or subject fetch missing).")

        print(f"\nCycle complete. New rows added this cycle: {len(new_rows_to_add)}")
        print(f"Waiting {CYCLE_SECONDS}s before next cycle...")
        time.sleep(CYCLE_SECONDS)

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")
