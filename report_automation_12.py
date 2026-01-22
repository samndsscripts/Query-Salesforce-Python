import time
import re
import os
import pandas as pd
import xlwings as xw
from rapidfuzz import process
from simple_salesforce import Salesforce
from datetime import datetime
from datetime import date
from datetime import timezone
from colorama import init as colorama_init, Fore, Style
from collections import Counter
from dotenv import load_dotenv  # <-- NEW import

# Initialize colorama
colorama_init(autoreset=True)

# --- Load environment variables securely ---
load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\salesforce.env")


SF_USERNAME = os.getenv("SF_USERNAME")
SF_PASSWORD = os.getenv("SF_PASSWORD")
SF_TOKEN    = os.getenv("SF_TOKEN")
SF_INSTANCE = os.getenv("SF_INSTANCE")

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
category_list = supplier_df['Category'].astype(str).tolist()


normalized_stock_codes = [
    re.sub(r"[^a-z0-9\s]", " ", s.lower()).strip() for s in stock_codes_list
]
# --- Load product prices data ---
sheet_prices = wb.sheets['Product Prices']
table_range = sheet_prices.tables['prices'].range
prices_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0]).fillna('')

stock_codes_prices = prices_df['Stock Code'].astype(str).tolist()
sale_prices = pd.to_numeric(prices_df['Sale Price'], errors='coerce').fillna(0).tolist()


product_categories = prices_df['Part Category'].astype(str).tolist()
descriptions       = prices_df['Description'].astype(str).tolist()

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
        return "N/A", [], None  # added None for bought_out

    # Exact match first
    exact_matches = []
    for i, code in enumerate(normalized_stock_codes):
        score = process.extractOne(nd, [code])[1]
        if score >= exact_threshold:
            supplier = suppliers_list[i].strip()
            mfg_plant = man_sites_list[i].strip() if i < len(man_sites_list) else ""
            if supplier:
                source = supplier
                bought_out = 1
            elif mfg_plant:
                source = mfg_plant
                bought_out = 0
            else:
                source = "N/A"
                bought_out = None
            exact_matches.append((stock_codes_list[i], source, score, bought_out))
    
    if exact_matches:
        exact_matches.sort(key=lambda x: x[2], reverse=True)
        code, source, _, bought_out = exact_matches[0]
        return source, [f"Stock Code: {code} | Source: {source}", '', ''], bought_out

    # Length-based candidate filter
    desc_len = len(nd)
    candidates = [
        (i, stock_codes_list[i], suppliers_list[i], man_sites_list[i], normalized_stock_codes[i])
        for i in range(len(stock_codes_list))
        if 0.9 * desc_len <= len(normalized_stock_codes[i]) <= 1.1 * desc_len
    ]

    # You can continue your fuzzy matching logic here, making sure to set `bought_out` similarly.
    # For candidates that match by fuzzy score:
    for i, code, supplier, mfg_plant, norm_code in candidates:
        score = process.extractOne(nd, [norm_code])[1]
        if score >= fuzzy_threshold:
            if supplier:
                source = supplier
                bought_out = 1
            elif mfg_plant:
                source = mfg_plant
                bought_out = 0
            else:
                source = "N/A"
                bought_out = None
            return source, [f"Stock Code: {code} | Source: {source}", '', ''], bought_out

    return "N/A", [], None

# --- Determine Source from Comments (Exact Match Only) ---
def determine_source_from_comments(comment_text, stock_code_list, supplier_list, mfg_list, top_n=3, exact_threshold=90):
    """
    Determines the source (supplier or manufacturing plant) from a comment text,
    returns source, top matches, and bought_out flag (1 if supplier, 0 if mfg plant).
    """
    nd = normalize(comment_text)
    if not nd:
        return "N/A", [], None  # No source found

    exact_matches = []
    for i, code in enumerate(stock_code_list):
        norm_code = normalize(code)
        score = process.extractOne(nd, [norm_code])[1]
        if score >= exact_threshold:
            supplier = supplier_list[i].strip() if supplier_list[i] else ''
            mfg = mfg_list[i].strip() if mfg_list[i] else ''
            source = supplier if supplier else (mfg if mfg else "N/A")
            exact_matches.append((code, source, score, supplier, mfg))

    if not exact_matches:
        return "N/A", [], None

    # Sort by score descending
    exact_matches.sort(key=lambda x: x[2], reverse=True)
    code, source, _, supplier, mfg = exact_matches[0]

    # Determine bought_out
    bought_out = 1 if supplier else 0

    # Build top matches list
    top_matches = [f"Stock Code: {code} | Source: {source}"]
    while len(top_matches) < top_n:
        top_matches.append('')

    return source, top_matches, bought_out

# --- Load existing RMA_Raw table safely ---
RMA_table = sheet_out.tables['RMA_Raw']

# Pull headers and body safely
rma_headers = RMA_table.header_row_range.value if RMA_table.header_row_range else []
rma_data = RMA_table.data_body_range.value if RMA_table.data_body_range else []

# Build DataFrame
RMA_DF = pd.DataFrame(rma_data, columns=rma_headers) if rma_data else pd.DataFrame(columns=rma_headers)
RMA_DF = RMA_DF.fillna('')

# Build set of existing cases
parsed_cases = [parse_case_number(x) for x in RMA_DF.get('Case Number', pd.Series([])) if x]
existing_cases = set([x for x in parsed_cases if x is not None])

# --- DEBUG ---
print(Fore.MAGENTA + "--- Current RMA_Raw Table ---")
print(RMA_DF.head())
print(Fore.MAGENTA + f"Existing cases count: {len(existing_cases)}\n")


# --- Main Loop ---
CYCLE_SECONDS = 1800  # 30 minutes
SOQL_BATCH = 200

try:
    while True:
        cycle_start_time = time.time()
        clear_console()

        # --- Load existing cases from Excel to prevent duplicates ---
        table_vals = RMA_table.range.value
        if table_vals and len(table_vals) > 1:
            output_df = pd.DataFrame(table_vals[1:], columns=table_vals[0])
        else:
            output_df = pd.DataFrame(columns=table_vals[0] if table_vals else [])

        parsed_cases = output_df.get('Case Number', pd.Series([])).apply(parse_case_number).tolist()
        existing_cases = set(filter(None, parsed_cases))

        # --- Rolling 12 months start date ---
        today = datetime.utcnow()

        # Add back
        one_year_ago = today - pd.DateOffset(months=12)
        start_date = one_year_ago.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        print("🔄 Pulling Salesforce RMA cases from last 12 months with PROD reasons...\n")

        try:
            # --- SOQL query ---
            soql = f"""
                SELECT 
                    Id,
                    CaseNumber,
                    Subject,
                    Description,
                    RMA_Value__c,
                    Type,
                    CreatedDate,
                    Owner.Name,
                    Account.Name,
                    Contact_Type__c,
                    WH__c,
                    Reason,
                    Comments_Exist__c
                FROM Case
                WHERE 
                    CreatedDate >= {start_date}
                    AND RecordType.Name = 'RMA'
                    AND Reason LIKE 'PROD%'
                ORDER BY CreatedDate ASC
            """
            result = sf.query_all(soql)
            records = result.get('records', [])
            print(f"✅ Retrieved {len(records)} RMA case(s) from Salesforce.\n")

            # --- Collect comments ---
            cases_with_comments = [r['Id'] for r in records if r.get('Comments_Exist__c')]
            comments_map = {}
            if cases_with_comments:
                quoted_ids = ",".join(f"'{cid}'" for cid in cases_with_comments)
                soql_comments = f"""
                    SELECT ParentId, CommentBody
                    FROM CaseComment
                    WHERE ParentId IN ({quoted_ids})
                    ORDER BY CreatedDate ASC
                """
                comments_result = sf.query_all(soql_comments)
                for rec in comments_result.get('records', []):
                    pid = rec['ParentId']
                    body = rec.get('CommentBody', '')
                    comments_map[pid] = (comments_map[pid] + "\n---\n" + body
                                         if pid in comments_map else body)

            # --- Build new rows ---
            new_rows_to_add = []
            qty_tally = Counter()

            for rec in records:
                cn_int = parse_case_number(rec.get('CaseNumber', ''))
                if not cn_int or cn_int in existing_cases:
                    continue  # Skip old/duplicate rows

                description = rec.get('Description', '')
                subject = rec.get('Subject', '')
                rma_value = rec.get('RMA_Value__c', 0)
                title_qty = extract_quantity(subject) or 0

                # Fuzzy match source
                source, top_matches, bought_out = determine_source(description, top_n=3)
                matched_stock_code = None
                if top_matches and "Stock Code:" in top_matches[0]:
                    matched_stock_code = top_matches[0].split("|")[0].replace("Stock Code:", "").strip()

                if matched_stock_code:
                    try:
                        idx = stock_codes_list.index(matched_stock_code)
                        category = category_list[idx] or "Unidentified"
                    except ValueError:
                        category = "Unidentified"
                else:
                    category = "Unidentified"


                qty = title_qty
                qty_source = "title_qty"

                if matched_stock_code:
                    normalized_prices_codes = [normalize(c) for c in stock_codes_prices]
                    idx_match = next((i for i, c in enumerate(normalized_prices_codes)
                                      if normalize(matched_stock_code) == c), None)
                    if idx_match is not None and sale_prices[idx_match] > 0:
                        price_qty = round(float(rma_value) / sale_prices[idx_match])
                        if price_qty > 0:
                            qty = price_qty
                            qty_source = "price_qty"

                qty_tally[qty_source] += 1

                dt = datetime.strptime(rec.get('CreatedDate', '').split('T')[0], "%Y-%m-%d")
                opened_date_formatted = f"{dt.month}/{dt.day}/{dt.year}"
                year_int = dt.year
                month_int = dt.month

                case_id = rec["Id"]
                case_number = cn_int

                hyperlink = case_number


                new_row = [
                    opened_date_formatted,
                    year_int,
                    month_int,
                    rec.get('Reason', ''),
                    rec.get('Owner', {}).get('Name', ''),
                    hyperlink,
                    description,
                    comments_map.get(rec['Id'], ''),
                    qty or '',
                    rma_value,
                    rec.get('Type', ''),
                    rec.get('Account', {}).get('Name', ''),
                    rec.get('Contact_Type__c', ''),
                    rec.get('WH__c', ''),
                    top_matches[0] if len(top_matches) > 0 else '',
                    top_matches[1] if len(top_matches) > 1 else '',
                    top_matches[2] if len(top_matches) > 2 else '',
                    source,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    bought_out,
                    category,
                    matched_stock_code,
                    "YES"
                ]
                new_rows_to_add.append({
                "row": new_row,
                "case_id": case_id,
                "case_number": case_number
                })

                existing_cases.add(cn_int)  # Mark as added

            
            # --- Append new rows to the table first ---
            table_start_row = RMA_table.data_body_range.last_cell.row + 1 if RMA_table.data_body_range else RMA_table.range.row + 1
            table_start_col = RMA_table.range.column

            # Write new rows
            sheet_out.range((table_start_row, table_start_col)).value = [item["row"] for item in new_rows_to_add]

            # --- Build formulas for the Case Number column, matching the new rows ---
            rma_headers = RMA_table.header_row_range.value
            case_col_idx = rma_headers.index("Case Number")  # 0-based index

            hyperlink_formulas = [
                f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{item["case_id"]}/view","{item["case_number"]}")'
                for item in new_rows_to_add
            ]

            # Write formulas directly into the Case Number column of the table
            for i, formula in enumerate(hyperlink_formulas):
                # Row in sheet = first appended row + offset
                sheet_out.range((table_start_row + i, table_start_col + case_col_idx)).formula = formula


            # --- Countdown until next cycle ---
            elapsed = time.time() - cycle_start_time
            remaining_seconds = max(0, CYCLE_SECONDS - int(elapsed))
            for sec in range(remaining_seconds, 0, -1):
                hours, remainder = divmod(sec, 3600)
                minutes, seconds = divmod(remainder, 60)
                countdown_str = f"{hours:02}:{minutes:02}:{seconds:02}"
                print(f"🔄 Pulling Salesforce RMA cases... (Last cycle: {int(elapsed)} sec) ⏳ Next pull in: {countdown_str}", end='\r', flush=True)
                time.sleep(1)

        except Exception as e:
            print(Fore.RED + "❌ Error during Salesforce pull:", str(e))

except KeyboardInterrupt:
    print("\n" + Style.BRIGHT + "Script stopped by user.")
