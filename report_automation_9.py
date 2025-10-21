import time
import re
import os
import pandas as pd
import xlwings as xw
from rapidfuzz import process
from simple_salesforce import Salesforce
from datetime import datetime
from colorama import init as colorama_init, Fore, Style
from collections import Counter

# Initialize colorama
colorama_init(autoreset=True)

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

# --- Determine Source with top-N fuzzy matches ---
def determine_source(description, top_n=3, threshold=70):
    nd = normalize(description)
    if not nd:
        return "N/A", []

    # --- Get top N fuzzy matches ---
    matches = process.extract(nd, normalized_stock_codes, limit=top_n)
    
    top_sources_with_scores = []
    for match in matches:
        best_norm_code, score = match[0], match[1]
        if score < threshold:
            continue  # skip low-confidence matches
        try:
            idx = normalized_stock_codes.index(best_norm_code)
            supplier = suppliers_list[idx].strip()
            source = supplier if supplier else (man_sites_list[idx] if idx < len(man_sites_list) else "N/A")
            top_sources_with_scores.append((source, score))
        except Exception:
            top_sources_with_scores.append(("N/A", score))
    
    # --- Sort sources by descending score ---
    top_sources_with_scores.sort(key=lambda x: x[1], reverse=True)
    top_sources = [s for s, sc in top_sources_with_scores]

    # --- Determine final source by most common among top matches ---
    if top_sources:
        source_counter = Counter(top_sources)
        final_source = source_counter.most_common(1)[0][0]
    else:
        final_source = "N/A"

    return final_source, top_sources

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
            print(Fore.RED + "❌ Error fetching report:", e)
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

        # --- Batch SOQL: get titles and comments ---
        case_subject_map = {}
        case_comments_map = {}
        for batch in [new_case_numbers[i:i+SOQL_BATCH] for i in range(0, len(new_case_numbers), SOQL_BATCH)]:
            quoted = ",".join(f"'{cn}'" for cn in batch)
            soql = f"""
                SELECT CaseNumber, Subject, 
                (SELECT CommentBody FROM CaseComments ORDER BY CreatedDate DESC LIMIT 1)
                FROM Case 
                WHERE CaseNumber IN ({quoted})
            """
            try:
                result = sf.query_all(soql)
                for rec in result.get('records', []):
                    case_number = rec.get('CaseNumber', '')
                    case_subject_map[case_number] = rec.get('Subject', '')
                    comments = rec.get('CaseComments', {}).get('records', [])
                    latest_comment = comments[0]['CommentBody'] if comments else ''
                    case_comments_map[case_number] = latest_comment
            except Exception as e:
                print(Fore.YELLOW + "⚠️ SOQL batch failed:", e)

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
            comments = case_comments_map.get(cn_raw, '')
            subject = case_subject_map.get(cn_raw, '')
            qty = extract_quantity(subject)  # use subject
            rma_value = cells[5].get('label', '')
            case_category = cells[6].get('label', '')
            account_name = cells[7].get('label', '')
            contact_type = cells[8].get('label', '')
            shipping_whse = cells[9].get('label', '')
            
            # --- Determine top source matches ---
            source, top_matches = determine_source(description, top_n=3, threshold=70)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            new_row = [
                cells[0].get('label', ''),          # Opened Date
                cells[1].get('label', ''),          # Case Reason
                cells[2].get('label', ''),          # Case Owner
                cn_int,                             # Case Number
                description,                        # Description
                comments,                           # Comments
                qty if qty else '',                 # Quantity
                rma_value,                          # RMA Value
                case_category,                      # Case Category
                account_name,                       # Account Name
                contact_type,                       # Contact Type
                shipping_whse,                      # Shipping Whse
                ", ".join(top_matches),             # Top Source Matches
                source,                             # Source
                timestamp                           # Time Stamp
            ]
            new_rows_to_add.append(new_row)
            existing_cases.add(cn_int)

        # --- Append to Excel inside table ---
        if new_rows_to_add:
            if table_out.data_body_range:
                start_row = table_out.data_body_range.last_cell.row + 1
            else:
                start_row = table_out.range.row + 1

            start_col = table_out.range.column
            num_table_cols = len(table_out.range.value[0]) if table_out.range.value else len(new_rows_to_add[0])

            for r in new_rows_to_add:
                if len(r) < num_table_cols:
                    r.extend([''] * (num_table_cols - len(r)))
                elif len(r) > num_table_cols:
                    r = r[:num_table_cols]

            sheet_out.range((start_row, start_col)).value = new_rows_to_add

            print(Fore.GREEN + f"\n✅ Added {len(new_rows_to_add)} new row(s) to Excel.\n")
            for nr in new_rows_to_add:
                print(f" → Case {nr[3]} | Qty: {nr[6] or 'N/A'} | Source: {nr[13]} | Added: {nr[14]}")
        else:
            print(Fore.YELLOW + "No valid new rows to add this cycle.")

        print(f"\nSleeping {CYCLE_SECONDS}s before next check...\n")
        time.sleep(CYCLE_SECONDS)

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")
