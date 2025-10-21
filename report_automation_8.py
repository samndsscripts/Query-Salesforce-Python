# version8_caseIDs_full.py
import time
import re
import os
from collections import defaultdict
import pandas as pd
import xlwings as xw
from simple_salesforce import Salesforce
from rapidfuzz import process
from colorama import init as colorama_init, Fore, Style

# Initialize colorama for colored terminal output
colorama_init(autoreset=True)

# ----- Salesforce login -----
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)

# ----- Excel workbook setup -----
WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"
SUPPLIER_SHEET = 'Class Site Supplier'
SUPPLIER_TABLE = 'Table3'
OUTPUT_SHEET = 'test'
OUTPUT_TABLE = 'test'

wb = xw.Book(WB_PATH)
sheet_sup = wb.sheets[SUPPLIER_SHEET]
sheet_out = wb.sheets[OUTPUT_SHEET]
table_out = sheet_out.tables[OUTPUT_TABLE]

# --- Load supplier data (precompute normalized lists) ---
table_range = sheet_sup.tables[SUPPLIER_TABLE].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0]).fillna('')

stock_codes_list = supplier_df['Stock Code'].astype(str).tolist()
man_sites_list = supplier_df['Mfg Plant Name'].astype(str).tolist()
suppliers_list = supplier_df['Supplier'].astype(str).tolist()

# Pre-normalize stock codes for fast fuzzy matching
def normalize(text: str) -> str:
    t = str(text or "").lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

normalized_stock_codes = [normalize(s) for s in stock_codes_list]

# --- Helpers ---
def clear_console():
    os.system('cls' if os.name == 'nt' else 'clear')

def parse_case_number(x):
    """
    Safer parsing: pull all digits and convert to int if possible.
    This handles values like 'Case 000123' or '123.0' or '123'.
    """
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

def extract_quantity(title):
    """
    Robust quantity extraction:
    - captures patterns like 'qty: 5', 'qty - 5', '5 pcs', '5 pc'
    - sums multiple matches in the same text
    """
    if not title:
        return None
    text = str(title).lower()
    # find "qty: 5" or "qty - 5" etc.
    qty_after_matches = re.findall(r'qty\s*[:\-]?\s*(\d+)', text)
    # find "5 pcs" or "5 pc" or "5pcs"
    pcs_before_matches = re.findall(r'(\d+)\s*pcs?\b', text)
    # also handle patterns like "5x" (optional)
    x_matches = re.findall(r'(\d+)\s*x\b', text)

    all_nums = qty_after_matches + pcs_before_matches + x_matches
    total = sum(int(x) for x in all_nums if x.isdigit()) if all_nums else 0

    return total if total > 0 else None

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def determine_source(description):
    """
    Use normalized_stock_codes to fuzzy-match against description.
    Return suppliers_list[idx] if present, else man_sites_list[idx], else 'N/A'.
    """
    nd = normalize(description)
    if not nd:
        return "N/A"
    match = process.extractOne(nd, normalized_stock_codes)
    if not match:
        return "N/A"
    best_norm_code = match[0]
    try:
        idx = normalized_stock_codes.index(best_norm_code)
        supplier = suppliers_list[idx].strip() if idx < len(suppliers_list) else ""
        if supplier:
            return supplier
        # fallback to manufacturing site if supplier blank
        return man_sites_list[idx] if idx < len(man_sites_list) and man_sites_list[idx].strip() else "N/A"
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
            report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', []) or []
        except Exception as e:
            print(Fore.RED + "❌ Error fetching report:", str(e))
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

        # ---------------------------------------------------------------------
        # --- Batch SOQL: Get Case IDs, Subjects, and Comments ---
        # ---------------------------------------------------------------------
        case_data_map = {}  # { CaseNumber: { 'Id': ..., 'Subject': ... } }

        for batch in chunks(new_case_numbers, SOQL_BATCH):
            quoted = ",".join(f"'{cn}'" for cn in batch)
            soql = f"SELECT Id, CaseNumber, Subject FROM Case WHERE CaseNumber IN ({quoted})"
            try:
                result = sf.query_all(soql)
                records = result.get('records', [])
                for rec in records:
                    cn = rec['CaseNumber']
                    case_data_map[cn] = {
                        'Id': rec.get('Id', ''),
                        'Subject': rec.get('Subject', '')
                    }
            except Exception as e:
                print(Fore.YELLOW + "⚠️ SOQL batch failed while getting IDs:", str(e))

        # --- Batch SOQL: Get Comments for those Case IDs ---
        case_comments_map = {}  # { CaseNumber: "concatenated comment bodies" }

        case_ids = [data['Id'] for data in case_data_map.values() if data.get('Id')]
        for batch in chunks(case_ids, SOQL_BATCH):
            quoted = ",".join(f"'{cid}'" for cid in batch)
            soql = f"SELECT ParentId, CommentBody FROM CaseComment WHERE ParentId IN ({quoted})"
            try:
                result = sf.query_all(soql)
                records = result.get('records', [])
                for rec in records:
                    parent_id = rec.get('ParentId', '')
                    comment = rec.get('CommentBody', '').strip()
                    if not comment:
                        continue
                    # Find CaseNumber by ID
                    for cn, data in case_data_map.items():
                        if data['Id'] == parent_id:
                            case_comments_map.setdefault(cn, []).append(comment)
                            break
            except Exception as e:
                print(Fore.YELLOW + "⚠️ SOQL batch failed while getting comments:", str(e))

        # --- Combine Subjects and Comments ---
        case_combined_map = {}
        for cn, data in case_data_map.items():
            case_combined_map[cn] = {
                'Subject': data.get('Subject', ''),
                'Comments': "\n---\n".join(case_comments_map.get(cn, []))
            }

        # ---------------------------------------------------------------------
        # --- Build and Append New Rows ---
        # ---------------------------------------------------------------------
        new_rows_to_add = []
        printed_rows = []

        for row in report_rows:
            cells = row.get('dataCells', [])
            if len(cells) < 10:
                continue
            cn_raw = str(cells[3].get('label', '')).strip()
            cn_int = parse_case_number(cn_raw)
            if not cn_int or cn_raw not in case_combined_map or cn_int in existing_cases:
                continue

            description = str(cells[4].get('label', ''))
            subject = case_combined_map[cn_raw].get('Subject', '')
            qty = extract_quantity(subject)
            source = determine_source(description)
            comments = case_combined_map[cn_raw].get('Comments', '')

            new_row = [
                cells[0].get('label', ''),  # Opened Date
                cells[1].get('label', ''),  # Case Reason
                cells[2].get('label', ''),  # Case Owner
                cn_int,                     # Case Number
                description,                # Description
                subject,                    # Subject
                qty if qty else '',         # Quantity
                comments,                   # Comments
                cells[5].get('label', ''),  # RMA Value
                cells[6].get('label', ''),  # Case Category
                cells[7].get('label', ''),  # Account Name
                cells[8].get('label', ''),  # Contact Type
                cells[9].get('label', ''),  # Shipping Whse
                source                      # Source (Supplier or Mfg Plant)
            ]
            new_rows_to_add.append(new_row)
            printed_rows.append((cn_int, qty, subject, bool(comments)))
            existing_cases.add(cn_int)

        # --- Write new rows to Excel ---
        if new_rows_to_add:
            list_obj = table_out.api
            for row_data in new_rows_to_add:
                # Add a new ListRow to the Table/ListObject
                list_obj.ListRows.Add()
                # Write to last row of the visible table range
                new_row_range = table_out.range.rows[-1]
                new_row_range.value = [row_data]

        # --- Print summary ---
        for cn_int, qty, subject, has_comment in printed_rows:
            comment_status = "💬" if has_comment else ""
            print(Fore.GREEN + f"New → Case {cn_int} | Qty: {qty or 'N/A'} | {comment_status} {subject}")

        print(f"\n✅ Added {len(new_rows_to_add)} new row(s). Sleeping {CYCLE_SECONDS}s...\n")
        time.sleep(CYCLE_SECONDS)

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")
