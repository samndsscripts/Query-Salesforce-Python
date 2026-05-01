import os
import io
import re
import sys
import time
import traceback
import pandas as pd
import xlwings as xw
from datetime import datetime, timezone
from dotenv import load_dotenv
import tableauserverclient as TSC
from simple_salesforce import Salesforce
from rapidfuzz import process
from colorama import init as colorama_init, Fore, Style

# Initialize colorama
colorama_init(autoreset=True)

# ============================================================================
# CONFIG
# ============================================================================

load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\salesforce.env")
load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\tableau.env")

WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"

# Salesforce creds
SF_USERNAME = os.getenv("SF_USERNAME")
SF_PASSWORD = os.getenv("SF_PASSWORD")
SF_TOKEN = os.getenv("SF_TOKEN")
SF_INSTANCE = os.getenv("SF_INSTANCE")

# Tableau creds
TABLEAU_SERVER = os.getenv("TABLEAU_SERVER")
TABLEAU_PAT_NAME = os.getenv("TABLEAU_PAT_NAME")
TABLEAU_PAT_SECRET = os.getenv("TABLEAU_PAT_SECRET")
TABLEAU_SITE = os.getenv("TABLEAU_SITE_CONTENTURL", "")
TABLEAU_SHIPMENT_VIEW_ID = "41b7e54e-3da8-4b53-bfaf-047410bb4fd8"
TABLEAU_RMA_VIEW_ID = "128dec6d-e4ab-4843-a337-43ebad53c779"
TABLEAU_PRODUCT_VIEW_ID = "241e5cc1-6d7c-484a-99ad-d5700c3fd281"

CYCLE_INTERVAL_SECONDS = 3600

# ============================================================================
# SHARED HELPERS
# ============================================================================

def clear_terminal():
    os.system("cls")


def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def banner(msg):
    print(f"\n{'=' * 60}")
    print(f"  {msg}")
    print(f"  {timestamp()}")
    print(f"{'=' * 60}\n")


def normalize_number(val):
    """Strip leading zeros and whitespace from a numeric string."""
    s = str(val).strip()
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def normalize_text(text):
    """Lowercase, remove special chars, collapse whitespace."""
    t = str(text or '').lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def truncate_for_fuzzy(text, max_chars=500):
    """
    Truncate text for fuzzy matching to avoid email chain noise.
    Takes the first N characters which is typically the actual complaint,
    not the forwarded email thread below.
    """
    if not text:
        return ""
    truncated = text[:max_chars]
    for marker in ["---", "From:", "-----", "Sent:", "Original Message"]:
        idx = truncated.find(marker)
        if idx > 50:
            return truncated[:idx].strip()
    return truncated.strip()


# ============================================================================
# TABLEAU AUTH (ONCE PER CYCLE)
# ============================================================================

def tableau_sign_in():
    auth = TSC.PersonalAccessTokenAuth(
        TABLEAU_PAT_NAME,
        TABLEAU_PAT_SECRET,
        site_id=TABLEAU_SITE
    )
    server = TSC.Server(TABLEAU_SERVER, use_server_version=True)
    server.auth.sign_in(auth)
    return server


def tableau_sign_out(server):
    try:
        server.auth.sign_out()
    except Exception:
        pass


# ============================================================================
# SHIPMENT HELPERS (LOCKED)
# ============================================================================

def shipment_clean_val(val):
    if pd.isna(val):
        return ''
    val = str(val).replace('\xa0', '').replace('\t', '').strip().upper()
    try:
        f = float(val.replace(',', ''))
        if f.is_integer():
            val = str(int(f))
    except:
        pass
    return val


def shipment_derive_source(row):
    part_category = str(row.get("Part Category", "")).strip().upper()
    if part_category == "M":
        return str(row.get("Mfg Plant Name", "")).strip().upper()
    if part_category == "B":
        return str(row.get("Supplier", "")).strip().upper()
    return ""


def shipment_generate_ident(row):
    return (
        shipment_clean_val(row.get("Mfg Plant Name")) + '|' +
        shipment_clean_val(row.get("Source")) + '|' +
        shipment_clean_val(row.get("Order Qty")) + '|' +
        shipment_clean_val(row.get("Year")) + '|' +
        shipment_clean_val(row.get("Month"))
    )


# ============================================================================
# RMA HELPERS (LOCKED)
# ============================================================================

def rma_parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None


def rma_clean_number(x):
    try:
        return float(str(x).replace(",", ""))
    except:
        return None


def rma_collect_case_comments(sf, case_ids):
    if not case_ids:
        return {}

    quoted_ids = ",".join(f"'{cid}'" for cid in case_ids)
    comments_map = {}

    soql = f"""
        SELECT ParentId, CommentBody
        FROM CaseComment
        WHERE ParentId IN ({quoted_ids})
        ORDER BY CreatedDate ASC
    """

    result = sf.query_all(soql)
    for rec in result.get("records", []):
        pid = rec["ParentId"]
        body = rec.get("CommentBody", "") or ""
        comments_map[pid] = (
            comments_map[pid] + "\n---\n" + body
            if pid in comments_map else body
        )

    return comments_map


def rma_derive_source(row):
    pc = str(row.get("Part Category", "")).upper()
    if pc == "M":
        return str(row.get("Mfg Plant Name", ""))
    if pc == "B":
        return str(row.get("Supplier", ""))
    return ""


# ============================================================================
# SUBJECT PARSING HELPERS
# ============================================================================

def extract_stock_code_from_subject(subject):
    """Extract stock code from subject line patterns like #1720C150 or #DS-226."""
    if not subject:
        return None
    match = re.search(r'#([\w\-/]+)', subject)
    if match:
        return match.group(1).strip()
    return None


def extract_quantity_from_subject(subject):
    """Extract quantity from subject line."""
    if not subject:
        return None
    text = subject.lower()
    matches = re.findall(r'(?:qty\s*[:\-]?\s*(\d+))|(?:(\d+)\s*pcs?)', text)
    total_qty = 0
    for m in matches:
        num = next((int(x) for x in m if x and x.isdigit()), None)
        if num:
            total_qty += num
    return total_qty if total_qty > 0 else None


def extract_stock_codes_from_items(rec):
    """Extract stock codes from Item_1__c through Item_10__c fields."""
    codes = []
    for i in range(1, 11):
        field = f"Item_{i}__c"
        val = rec.get(field)
        if val:
            cleaned = str(val).strip()
            # Skip template text
            if cleaned.startswith("SKU") or not cleaned:
                continue
            # Some Item fields have embedded data like "DS-092_124.0100_1.0000..."
            # Take just the stock code part before any underscore+number pattern
            code_part = re.split(r'_\d', cleaned)[0].strip()
            if code_part:
                codes.append(code_part)
    return codes


# ============================================================================
# PRODUCT TABLE LOOKUP
# ============================================================================

def fetch_product_table(server):
    """Fetch product reference data from Tableau."""
    print(f"  [{timestamp()}] Fetching product table from Tableau...")
    view = server.views.get_by_id(TABLEAU_PRODUCT_VIEW_ID)
    server.views.populate_csv(view)
    df = pd.read_csv(io.BytesIO(b"".join(view.csv)))
    df.columns = df.columns.str.strip()
    print(Fore.GREEN + f"  [{timestamp()}] Product table loaded: {len(df)} rows.")
    return df


def build_product_lookup(product_df):
    """Build lookup structures from product table for direct and fuzzy matching."""
    stock_codes = product_df['Stock Code'].astype(str).str.strip().tolist()
    normalized_codes = [normalize_text(s) for s in stock_codes]
    return product_df, stock_codes, normalized_codes


def lookup_product_by_stock_code(stock_code, product_df, stock_codes, normalized_codes):
    """Direct lookup: find product info by exact stock code match."""
    norm_input = normalize_text(stock_code)
    for i, norm_code in enumerate(normalized_codes):
        if norm_input == norm_code:
            row = product_df.iloc[i]
            pc = str(row.get("Part Category", "")).upper()
            if pc == "M":
                source = str(row.get("Mfg Plant Name", "")).strip()
            elif pc == "B":
                source = str(row.get("Supplier", "")).strip()
            else:
                source = ""
            category = str(row.get("Category", "")).strip()
            part_category = str(row.get("Part Category", "")).strip()
            return {
                "stock_code": stock_codes[i],
                "source": source,
                "category": category,
                "part_category": part_category,
                "match_type": "DIRECT"
            }
    return None


def fuzzy_match_product(text, product_df, stock_codes, normalized_codes, threshold=70):
    """Fuzzy match text against product stock codes."""
    norm_input = normalize_text(text)
    if not norm_input:
        return None

    best = process.extractOne(norm_input, normalized_codes, score_cutoff=threshold)
    if best:
        matched_norm, score, idx = best
        row = product_df.iloc[idx]
        pc = str(row.get("Part Category", "")).upper()
        if pc == "M":
            source = str(row.get("Mfg Plant Name", "")).strip()
        elif pc == "B":
            source = str(row.get("Supplier", "")).strip()
        else:
            source = ""
        category = str(row.get("Category", "")).strip()
        part_category = str(row.get("Part Category", "")).strip()
        return {
            "stock_code": stock_codes[idx],
            "source": source,
            "category": category,
            "part_category": part_category,
            "match_type": f"FUZZY ({score:.0f}%)"
        }
    return None


def resolve_product_info(rec, comments_text, product_df, stock_codes, normalized_codes):
    """
    Multi-pass product resolution:
    1. Item_1__c - Item_10__c (direct lookup)
    2. Subject (regex parse -> direct lookup)
    3. Description (fuzzy match, truncated to skip email chains)
    4. Comments (fuzzy match, truncated to skip email chains)
    """
    # PASS 1: Item fields (direct)
    item_codes = extract_stock_codes_from_items(rec)
    for code in item_codes:
        result = lookup_product_by_stock_code(code, product_df, stock_codes, normalized_codes)
        if result:
            result["match_type"] = "ITEM_FIELD"
            return result

    # PASS 2: Subject (regex -> direct)
    subject = rec.get("Subject", "") or ""
    parsed_code = extract_stock_code_from_subject(subject)
    if parsed_code:
        result = lookup_product_by_stock_code(parsed_code, product_df, stock_codes, normalized_codes)
        if result:
            result["match_type"] = "SUBJECT_PARSE"
            return result

    # PASS 3: Description (fuzzy, truncated to avoid email chain noise)
    description = rec.get("Description", "") or ""
    if description:
        truncated_desc = truncate_for_fuzzy(description)
        result = fuzzy_match_product(truncated_desc, product_df, stock_codes, normalized_codes, threshold=70)
        if result:
            return result

    # PASS 4: Comments (fuzzy, truncated, higher threshold)
    if comments_text:
        truncated_comments = truncate_for_fuzzy(comments_text)
        result = fuzzy_match_product(truncated_comments, product_df, stock_codes, normalized_codes, threshold=90)
        if result:
            return result

    return None


# ============================================================================
# RMA PIPELINE
# ============================================================================

def run_rma_pipeline(server, wb):
    banner("STARTING RMA PIPELINE")

    # STEP 1: Load existing cases from Excel
    sheet = wb.sheets["RMA Raw"]
    table = sheet.tables["RMA_Raw"]

    values = table.range.value
    if len(values) <= 1:
        existing = set()
    else:
        df_existing = pd.DataFrame(values[1:], columns=values[0])
        existing = set(
            df_existing["Case Number"]
            .apply(rma_parse_case_number)
            .dropna()
            .astype(int)
            .tolist()
        )

    # STEP 2: Salesforce query
    sf = Salesforce(
        username=SF_USERNAME,
        password=SF_PASSWORD,
        security_token=SF_TOKEN,
        instance_url=SF_INSTANCE
    )

    # -----------------------------------------------------------------------
    # HISTORICAL PULL - change this line back after initial backfill:
    # start_date = (datetime.now(timezone.utc) - pd.DateOffset(months=2)).strftime("%Y-%m-%dT%H:%M:%SZ")
    # -----------------------------------------------------------------------
    start_date = "2025-01-01T00:00:00Z"

    soql = f"""
        SELECT
            Id,
            CaseNumber,
            RMA_Number__c,
            RMA_Value__c,
            Order_Number__c,
            Owner.Name,
            Description,
            Subject,
            Account.Name,
            Contact_Type__c,
            WH__c,
            Reason,
            CreatedDate,
            Item_1__c,
            Item_2__c,
            Item_3__c,
            Item_4__c,
            Item_5__c,
            Item_6__c,
            Item_7__c,
            Item_8__c,
            Item_9__c,
            Item_10__c,
            Comments_Exist__c
        FROM Case
        WHERE
            CreatedDate >= {start_date}
            AND RecordType.Name = 'RMA'
            AND Reason LIKE 'PROD%'
            AND Reason != 'Product Inquiry'
    """

    records = sf.query_all(soql)["records"]
    print(f"  [{timestamp()}] Salesforce returned {len(records)} cases.")

    new_records = [
        r for r in records
        if rma_parse_case_number(r["CaseNumber"]) not in existing
    ]
    print(f"  [{timestamp()}] {len(new_records)} new cases after dedup.")

    if not new_records:
        print("✅ No new cases.")
        return

    # Collect comments
    comments = rma_collect_case_comments(
        sf, [r["Id"] for r in new_records]
    )

    # STEP 3: Fetch Tableau RMA data (using shared server)
    view = server.views.get_by_id(TABLEAU_RMA_VIEW_ID)
    server.views.populate_csv(view)
    tableau_df = pd.read_csv(io.BytesIO(b"".join(view.csv)))
    tableau_df.columns = tableau_df.columns.str.strip()

    # Normalize Tableau RMA numbers (strip leading zeros)
    tableau_df["_rma_norm"] = tableau_df["Rma Number"].astype(str).str.strip().apply(normalize_number)

    # Normalize Tableau Order numbers (strip leading zeros) if column exists
    has_order_col = "Original Order" in tableau_df.columns
    if has_order_col:
        tableau_df["_order_norm"] = tableau_df["Original Order"].astype(str).str.strip().apply(normalize_number)

    # Drop null Issue Date rows
    tableau_df = tableau_df[pd.notna(tableau_df["Issue Date"])].copy()

    # STEP 3b: Fetch product table for fuzzy fallback
    product_df = fetch_product_table(server)
    product_df_ref, stock_codes, normalized_codes = build_product_lookup(product_df)

    # STEP 4: Build case maps with normalized keys
    rma_case_map = {}
    order_case_map = {}
    orphan_cases = []

    for r in new_records:
        rma_raw = r.get("RMA_Number__c")
        order_raw = r.get("Order_Number__c")

        if rma_raw:
            rma_norm = normalize_number(rma_raw)
            rma_case_map[rma_norm] = r
        elif order_raw:
            order_norm = normalize_number(order_raw)
            order_case_map[order_norm] = r
        else:
            orphan_cases.append(r)

    print(f"  [{timestamp()}] Cases with RMA#: {len(rma_case_map)} | Order# fallback: {len(order_case_map)} | Orphans: {len(orphan_cases)}")

    # STEP 5: Match and build rows
    start_row = table.data_body_range.last_cell.row + 1 if table.data_body_range else table.range.row + 1
    start_col = table.range.column

    rows = []
    matched_case_ids = set()

    # --- PASS A: Match by RMA Number ---
    for _, t in tableau_df.iterrows():
        rma_norm = normalize_number(t["Rma Number"])
        sf_rec = rma_case_map.get(rma_norm)
        if not sf_rec:
            continue

        case_id = sf_rec["Id"]
        if case_id in matched_case_ids:
            continue
        matched_case_ids.add(case_id)

        case_num = rma_parse_case_number(sf_rec["CaseNumber"])
        qty = rma_clean_number(t.get("Ordered Qty"))

        excel_row = [
            t.get("Issue Date"),
            case_num,
            t.get("Problem"),
            sf_rec["Owner"]["Name"] if sf_rec.get("Owner") else "",
            sf_rec.get("Description", ""),
            comments.get(case_id, ""),
            int(qty) if qty else "",
            float(sf_rec.get("RMA_Value__c")) if sf_rec.get("RMA_Value__c") else "",
            sf_rec.get("Account", {}).get("Name", ""),
            sf_rec.get("Contact_Type__c", ""),
            rma_derive_source(t),
            t.get("Stock Code", ""),
            t.get("Part Category", ""),
            t.get("Category", ""),
            t.get("Warehouse Name", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "YES"
        ]

        link = f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{case_id}/view","{case_num}")'
        rows.append((excel_row, link))

    print(f"  [{timestamp()}] RMA# matched: {len(rows)}")

    # --- PASS B: Match by Order Number ---
    order_matched = 0
    if has_order_col and order_case_map:
        for _, t in tableau_df.iterrows():
            order_norm = normalize_number(t.get("Original Order", ""))
            sf_rec = order_case_map.get(order_norm)
            if not sf_rec:
                continue

            case_id = sf_rec["Id"]
            if case_id in matched_case_ids:
                continue
            matched_case_ids.add(case_id)

            case_num = rma_parse_case_number(sf_rec["CaseNumber"])
            qty = rma_clean_number(t.get("Ordered Qty"))

            excel_row = [
                t.get("Issue Date"),
                case_num,
                t.get("Problem"),
                sf_rec["Owner"]["Name"] if sf_rec.get("Owner") else "",
                sf_rec.get("Description", ""),
                comments.get(case_id, ""),
                int(qty) if qty else "",
                float(sf_rec.get("RMA_Value__c")) if sf_rec.get("RMA_Value__c") else "",
                sf_rec.get("Account", {}).get("Name", ""),
                sf_rec.get("Contact_Type__c", ""),
                rma_derive_source(t),
                t.get("Stock Code", ""),
                t.get("Part Category", ""),
                t.get("Category", ""),
                t.get("Warehouse Name", ""),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "YES"
            ]

            link = f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{case_id}/view","{case_num}")'
            rows.append((excel_row, link))
            order_matched += 1

    print(f"  [{timestamp()}] Order# matched: {order_matched}")

    # --- PASS C: Unmatched cases (no Tableau match) ---
    all_unmatched = list(orphan_cases)
    for r in rma_case_map.values():
        if r["Id"] not in matched_case_ids:
            all_unmatched.append(r)
    for r in order_case_map.values():
        if r["Id"] not in matched_case_ids:
            all_unmatched.append(r)

    fuzzy_matched = 0
    fuzzy_failed = 0

    for sf_rec in all_unmatched:
        case_id = sf_rec["Id"]
        if case_id in matched_case_ids:
            continue
        matched_case_ids.add(case_id)

        case_num = rma_parse_case_number(sf_rec["CaseNumber"])
        comment_text = comments.get(case_id, "")

        # Try product resolution (Item fields -> Subject -> Description -> Comments)
        product_info = resolve_product_info(
            sf_rec, comment_text, product_df_ref, stock_codes, normalized_codes
        )

        # Extract qty from subject
        subject = sf_rec.get("Subject", "") or ""
        qty = extract_quantity_from_subject(subject)

        if product_info:
            source = product_info["source"] or ""
            stock_code = product_info["stock_code"] or ""
            category = product_info["category"] or ""
            part_category = product_info["part_category"] or ""
            fuzzy_matched += 1
        else:
            source = "Unidentified"
            stock_code = ""
            category = ""
            part_category = ""
            fuzzy_failed += 1

        # Parse date from CreatedDate
        created = sf_rec.get("CreatedDate", "")
        try:
            dt = datetime.strptime(created.split("T")[0], "%Y-%m-%d")
            issue_date = f"{dt.month}/{dt.day}/{dt.year}"
        except:
            issue_date = ""

        excel_row = [
            issue_date,
            case_num,
            sf_rec.get("Reason", ""),
            sf_rec["Owner"]["Name"] if sf_rec.get("Owner") else "",
            sf_rec.get("Description", ""),
            comment_text,
            int(qty) if qty else "",
            float(sf_rec.get("RMA_Value__c")) if sf_rec.get("RMA_Value__c") else "",
            sf_rec.get("Account", {}).get("Name", ""),
            sf_rec.get("Contact_Type__c", ""),
            source,
            stock_code,
            part_category,
            category,
            sf_rec.get("WH__c", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "YES"
        ]

        link = f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{case_id}/view","{case_num}")'
        rows.append((excel_row, link))

    print(f"  [{timestamp()}] Orphan/unmatched: {len(all_unmatched)} (product matched: {fuzzy_matched}, no match: {fuzzy_failed})")

    # STEP 6: Write to Excel
    if not rows:
        print(Fore.YELLOW + "⚠️ No rows to write.")
        return

    table.resize(table.range.resize(table.range.rows.count + len(rows)))
    sheet.range((start_row, start_col)).value = [r[0] for r in rows]

    case_idx = table.header_row_range.value.index("Case Number")
    for i, (_, link) in enumerate(rows):
        sheet.range((start_row + i, start_col + case_idx)).formula = link

    wb.save()
    print(Fore.GREEN + f"✅ {len(rows)} RMA rows written to Excel.")


# ============================================================================
# SHIPMENT PIPELINE (LOCKED - NO CHANGES)
# ============================================================================

def run_shipment_pipeline(server, wb):
    banner("STARTING SHIPMENT PIPELINE")

    sheet = wb.sheets["Shipment Data"]
    table = sheet.tables["shipment_data"]

    # Fetch from Tableau (using shared server)
    view = server.views.get_by_id(TABLEAU_SHIPMENT_VIEW_ID)
    server.views.populate_csv(view)
    df = pd.read_csv(io.BytesIO(b"".join(view.csv)))

    df.columns = df.columns.str.strip()

    if df.empty:
        print("✅ No shipment data returned from Tableau.")
        return

    required_columns = [
        "Month, Day, Year of Invoice Date",
        "Category",
        "Mfg Plant Name",
        "Supplier",
        "Part Category",
        "Qty"
    ]

    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise RuntimeError(f"Shipment view missing columns: {missing}")

    df = df[required_columns].copy()

    # ---------------- Business Fields ----------------
    df["Bought Out"] = df["Part Category"]
    df["Source"] = df.apply(shipment_derive_source, axis=1)

    # ---------------- Date ----------------
    df["Ship Date"] = pd.to_datetime(
        df["Month, Day, Year of Invoice Date"], errors="coerce"
    )
    df = df[df["Ship Date"].notna()]
    if df.empty:
        print("✅ No rows with valid Ship Date.")
        return

    df["Date Shipped"] = df["Ship Date"].dt.strftime("%m/%d/%Y")
    df["Year"] = df["Ship Date"].dt.year
    df["Month"] = df["Ship Date"].dt.month

    # ---------------- Quantity ----------------
    df["Order Qty"] = (
        df["Qty"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
        .astype(int)
    )

    # ---------------- Ident (AFTER ALL INPUTS EXIST) ----------------
    df["Ident"] = df.apply(shipment_generate_ident, axis=1)

    # ---------------- Excel Dedup ----------------
    table_vals = table.range.value
    if table_vals and len(table_vals) > 1:
        excel_df = pd.DataFrame(
            table_vals[1:], columns=table_vals[0]
        ).fillna("")
        existing_idents = set(excel_df["Ident"].astype(str))
    else:
        existing_idents = set()

    new_rows = df[~df["Ident"].isin(existing_idents)].copy()
    if new_rows.empty:
        print("✅ No new shipment rows.")
        return

    # ---------------- Output ----------------
    new_rows["TimeStamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    out = new_rows[
        [
            "Ident",
            "Date Shipped",
            "Category",
            "Order Qty",
            "Source",
            "Bought Out",
            "TimeStamp"
        ]
    ].fillna("").values.tolist()

    start_row = (
        table.data_body_range.last_cell.row + 1
        if table.data_body_range
        else table.range.row + 1
    )
    start_col = table.range.column

    table.resize(table.range.resize(table.range.rows.count + len(out)))
    sheet.range((start_row, start_col)).value = out
    wb.save()

    print(Fore.GREEN + f"✅ {len(out)} shipment rows appended.")


# ============================================================================
# COUNTDOWN
# ============================================================================

def countdown(seconds):
    while seconds > 0:
        mins, secs = divmod(seconds, 60)
        sys.stdout.write(f"\r  Next run in {mins:02d}:{secs:02d}  ")
        sys.stdout.flush()
        time.sleep(1)
        seconds -= 1
    sys.stdout.write("\r" + " " * 40 + "\r")
    sys.stdout.flush()


# ============================================================================
# MAIN LOOP
# ============================================================================

def main():
    clear_terminal()
    print(Fore.CYAN + Style.BRIGHT + "\n  COMBINED PIPELINE - RMA + SHIPMENTS")
    print(Fore.CYAN + f"  Cycle interval: {CYCLE_INTERVAL_SECONDS // 60} minutes\n")

    while True:
        cycle_start = time.time()
        clear_terminal()

        banner("CYCLE STARTING")

        server = None
        try:
            wb = xw.Book(WB_PATH)

            # Authenticate to Tableau ONCE
            print(f"  [{timestamp()}] Authenticating to Tableau...")
            server = tableau_sign_in()
            print(Fore.GREEN + f"  [{timestamp()}] Tableau authenticated.\n")

            # --- RMA Pipeline ---
            try:
                run_rma_pipeline(server, wb)
            except Exception:
                print(Fore.RED + f"\n  ❌ RMA PIPELINE ERROR:\n")
                traceback.print_exc()

            # --- Shipment Pipeline ---
            try:
                run_shipment_pipeline(server, wb)
            except Exception:
                print(Fore.RED + f"\n  ❌ SHIPMENT PIPELINE ERROR:\n")
                traceback.print_exc()

        except Exception:
            print(Fore.RED + f"\n  ❌ CYCLE ERROR:\n")
            traceback.print_exc()

        finally:
            if server:
                tableau_sign_out(server)
                print(f"\n  [{timestamp()}] Tableau signed out.")

        banner("CYCLE COMPLETE")

        elapsed = time.time() - cycle_start
        remaining = max(0, CYCLE_INTERVAL_SECONDS - int(elapsed))

        if remaining > 0:
            countdown(remaining)


if __name__ == "__main__":
    main()
