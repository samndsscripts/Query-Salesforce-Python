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
from dotenv import load_dotenv  # <-- NEW import

# Initialize colorama
colorama_init(autoreset=True)

# --- Load environment variables securely ---
load_dotenv(r"C:\Users\emp35107\Documents\.env")

SF_USERNAME = os.getenv("SALESFORCE_USERNAME")
SF_PASSWORD = os.getenv("SALESFORCE_PASSWORD")
SF_TOKEN    = os.getenv("SALESFORCE_TOKEN")
SF_INSTANCE = os.getenv("SALESFORCE_INSTANCE")

if not all([SF_USERNAME, SF_PASSWORD, SF_TOKEN, SF_INSTANCE]):
    raise ValueError("❌ Missing one or more Salesforce environment variables in .env file")

# --- Secure Salesforce login ---
sf = Salesforce(
    username=SF_USERNAME,
    password=SF_PASSWORD,
    security_token=SF_TOKEN,
    instance_url=SF_INSTANCE
)

# --- Workbook setup ---
WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"
wb = xw.Book(WB_PATH)
sheet_sup = wb.sheets['Class Site Supplier']
sheet_out = wb.sheets['RMA Raw']
table_out = sheet_out.tables['RMA_Raw']

# --- Load supplier data ---
table_range = sheet_sup.tables['Table3'].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])
supplier_df = supplier_df.fillna('')

stock_codes_list = supplier_df['Stock Code'].astype(str).tolist()
man_sites_list   = supplier_df['Mfg Plant Name'].astype(str).tolist()
suppliers_list   = supplier_df['Supplier'].astype(str).tolist()

normalized_stock_codes = [
    re.sub(r"[^a-z0-9\s]", " ", s.lower()).strip() for s in stock_codes_list
]

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
    """Extract quantity from case title (subject line)."""
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
def determine_source(description, top_n=3, fuzzy_threshold=70, exact_threshold=90):
    nd = normalize(description)
    if not nd:
        return "N/A", []

    # Exact match first
    exact_matches = []
    for i, code in enumerate(normalized_stock_codes):
        score = process.extractOne(nd, [code])[1]
        if score >= exact_threshold:
            supplier = suppliers_list[i].strip()
            source = supplier if supplier else (man_sites_list[i].strip() if i < len(man_sites_list) else "N/A")
            exact_matches.append((stock_codes_list[i], source, score))
    
    if exact_matches:
        exact_matches.sort(key=lambda x: x[2], reverse=True)
        code, source, _ = exact_matches[0]
        return source, [f"Stock Code: {code} | Source: {source}", '', '']

    # Length-based candidate filter
    desc_len = len(nd)
    candidates = [
        (i, stock_codes_list[i], suppliers_list[i], man_sites_list[i], normalized_stock_codes[i])
        for i in range(len(stock_codes_list))
        if 0.9 * desc_len <= len(normalized_stock_codes[i]) <= 1.1 * desc_len
    ]

    # Fuzzy match
    top_sources_with_scores = []
    for i, code, supplier, mfg, norm_code in candidates:
        score = process.extractOne(nd, [norm_code])[1]
        if score >= fuzzy_threshold:
            source = supplier.strip() if supplier.strip() else (mfg.strip() if mfg.strip() else "N/A")
            top_sources_with_scores.append((code, source, score))

    top_sources_with_scores.sort(key=lambda x: x[2], reverse=True)
    top_matches = [f"Stock Code: {code} | Source: {src}" for code, src, _ in top_sources_with_scores[:top_n]]
    while len(top_matches) < 3:
        top_matches.append('')

    final_source = top_sources_with_scores[0][1] if top_sources_with_scores else "N/A"
    return final_source, top_matches

# --- Determine Source from Comments (Exact Match Only) ---
def determine_source_from_comments(comment_text, stock_code_list, supplier_list, mfg_list, top_n=3, exact_threshold=90):
    nd = normalize(comment_text)
    if not nd:
        return "N/A", []

    exact_matches = []
    for i, code in enumerate(stock_code_list):
        norm_code = normalize(code)
        score = process.extractOne(nd, [norm_code])[1]
        if score >= exact_threshold:
            supplier = supplier_list[i].strip()
            source = supplier if supplier else (mfg_list[i].strip() if mfg_list[i] else "N/A")
            exact_matches.append((code, source, score))

    if not exact_matches:
        return "N/A", []

    exact_matches.sort(key=lambda x: x[2], reverse=True)
    code, source, _ = exact_matches[0]

    top_matches = [f"Stock Code: {code} | Source: {source}"]
    while len(top_matches) < top_n:
        top_matches.append('')

    return source, top_matches

# --- Main Loop ---
REPORT_ID = "00OUI00000EsGR72AN"
CYCLE_SECONDS = 1800  # 30 minutes
SOQL_BATCH = 100

try:
    while True:
        cycle_start_time = time.time()
        clear_console()
        print("🔄 Checking for new Salesforce cases... (Time elapsed: 00:00:00)\n")

        # Load existing Excel data
        table_vals = table_out.range.value
        if table_vals and len(table_vals) > 1:
            output_df = pd.DataFrame(table_vals[1:], columns=table_vals[0])
        else:
            output_df = pd.DataFrame(columns=table_vals[0] if table_vals else [])

        parsed_cases = output_df.get('Case Number', pd.Series([])).apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # Pull Salesforce report
        try:
            report_data = sf.restful(f'analytics/reports/{REPORT_ID}', params={'includeDetails': 'true'})
            report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])
        except Exception as e:
            print(Fore.RED + "❌ Error fetching report:", str(e))
            time.sleep(CYCLE_SECONDS)
            continue

        # Collect new case numbers
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
            print(f"No new cases.\n")

        # Batch SOQL: get titles and comments
        case_subject_map = {}
        case_comments_map = {}
        case_number_to_id = {}

        for batch in [new_case_numbers[i:i+SOQL_BATCH] for i in range(0, len(new_case_numbers), SOQL_BATCH)]:
            quoted = ",".join(f"'{cn}'" for cn in batch)

            soql_titles = f"""
                SELECT Id, CaseNumber, Subject
                FROM Case
                WHERE CaseNumber IN ({quoted})
            """
            try:
                result_titles = sf.query_all(soql_titles)
                if result_titles is None:
                    print(Fore.YELLOW + f"⚠️ Title query returned None for batch: {batch[:5]}...")
                    continue

                case_ids = []
                for rec in result_titles.get('records', []):
                    case_number = rec.get('CaseNumber', '')
                    subject = rec.get('Subject', None)
                    case_subject_map[case_number] = subject or ''
                    case_number_to_id[case_number] = rec.get('Id', '')
                    if rec.get('Id'):
                        case_ids.append(rec['Id'])

                if case_ids:
                    quoted_ids = ",".join(f"'{cid}'" for cid in case_ids)
                    soql_comments = f"""
                        SELECT ParentId, CommentBody
                        FROM CaseComment
                        WHERE ParentId IN ({quoted_ids})
                        ORDER BY CreatedDate DESC
                    """
                    try:
                        result_comments = sf.query_all(soql_comments)
                        if result_comments and 'records' in result_comments:
                            for rec in result_comments['records']:
                                parent_id = rec.get('ParentId', '')
                                comment_body = rec.get('CommentBody', '')
                                if parent_id not in case_comments_map:
                                    case_comments_map[parent_id] = comment_body
                        else:
                            print(Fore.CYAN + f"⚠️ No comments found for batch {batch[:5]}...")
                    except Exception as e:
                        print(Fore.YELLOW + f"⚠️ Comment query failed: {e}")
                else:
                    print(Fore.CYAN + f"⚠️ No Case IDs found for batch {batch[:5]}...")
            except Exception as e:
                print(Fore.YELLOW + f"⚠️ SOQL title query failed: {e}")

        # Build new rows
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
            case_id = case_number_to_id.get(cn_raw, None)
            comments = case_comments_map.get(case_id, '') if case_id else ''
            subject = case_subject_map.get(cn_raw, '')
            qty = extract_quantity(subject)
            rma_value = cells[5].get('label', '')
            case_category = cells[6].get('label', '')
            account_name = cells[7].get('label', '')
            contact_type = cells[8].get('label', '')
            shipping_whse = cells[9].get('label', '')

            # --- Comment-based source exact match first ---
            comment_source, comment_matches = determine_source_from_comments(
                comments,
                stock_codes_list,
                suppliers_list,
                man_sites_list,
                top_n=3
            )

            if comment_source != "N/A":
                source = comment_source
                top_matches = comment_matches
            else:
                source, top_matches = determine_source(description, top_n=3)

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            top1 = top_matches[0] if len(top_matches) > 0 else ''
            top2 = top_matches[1] if len(top_matches) > 1 else ''
            top3 = top_matches[2] if len(top_matches) > 2 else ''

            new_row = [
                cells[0].get('label', ''),  # Opened Date
                cells[1].get('label', ''),  # Case Reason
                cells[2].get('label', ''),  # Case Owner
                cn_int,                     # Case Number
                description,                # Description
                comments,                   # Comments
                qty if qty else '',         # Quantity
                rma_value,                  # RMA Value
                case_category,              # Case Category
                account_name,               # Account Name
                contact_type,               # Contact Type
                shipping_whse,              # Shipping Whse
                top1,                       # Top Match 1
                top2,                       # Top Match 2
                top3,                       # Top Match 3
                source,                     # Source
                timestamp                   # Time Stamp
            ]
            new_rows_to_add.append(new_row)
            existing_cases.add(cn_int)

        # Append to Excel
        if new_rows_to_add:
            if table_out.data_body_range:
                start_row = table_out.data_body_range.last_cell.row + 1
            else:
                start_row = table_out.range.row + 1

            start_col = table_out.range.column
            sheet_out.range((start_row, start_col)).value = new_rows_to_add

            print(Fore.GREEN + f"\n✅ Added {len(new_rows_to_add)} new row(s) to Excel.\n")
            for nr in new_rows_to_add:
                print(f" → Case {nr[3]} | Qty: {nr[6] or 'N/A'} | Source: {nr[15]} | Added: {nr[16]}")
        else:
            print(Fore.YELLOW + "No valid new rows to add this cycle.")

        # --- Adaptive countdown until next cycle ---
        elapsed = time.time() - cycle_start_time
        remaining_seconds = max(0, CYCLE_SECONDS - int(elapsed))

        for sec in range(remaining_seconds, 0, -1):
            hours, remainder = divmod(sec, 3600)
            minutes, seconds = divmod(remainder, 60)
            countdown_str = f"{hours:02}:{minutes:02}:{seconds:02}"
            clear_console()
            print(f"🔄 Checking Salesforce cases... (Last cycle duration: {int(elapsed)} sec)\n")
            print(f"⏳ Time until next check (hibernation countdown): {countdown_str}")
            time.sleep(1)

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")

