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
import truststore

truststore.inject_into_ssl()
colorama_init(autoreset=True)

# ============================================================================
# CONFIG
# ============================================================================
load_dotenv(r"C:\Users\samuel.cooper\OneDrive - Advanced Drainage Systems\Documents\salesforce.env")
load_dotenv(r"C:\Users\samuel.cooper\OneDrive - Advanced Drainage Systems\Documents\tableau.env")

WB_PATH = r"C:\Users\samuel.cooper\OneDrive - Advanced Drainage Systems\Documents\PPMs.xlsx"

SF_USERNAME = os.getenv("SF_USERNAME")
SF_PASSWORD = os.getenv("SF_PASSWORD")
SF_TOKEN = os.getenv("SF_TOKEN")
SF_INSTANCE = os.getenv("SF_INSTANCE")

TABLEAU_SERVER = os.getenv("TABLEAU_SERVER")
TABLEAU_PAT_NAME = os.getenv("TABLEAU_PAT_NAME")
TABLEAU_PAT_SECRET = os.getenv("TABLEAU_PAT_SECRET")
TABLEAU_SITE = os.getenv("TABLEAU_SITE_CONTENTURL", "")

TABLEAU_RMA_VIEW_ID = "30de1fa1-c6bd-4b5a-bfe8-020da690af8d"
TABLEAU_SHIPMENT_VIEW_ID = "41b7e54e-3da8-4b53-bfaf-047410bb4fd8"
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
    try:
        return str(int(float(str(val).strip())))
    except:
        return None

def normalize_text(text):
    t = str(text or '').lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()

def truncate_for_fuzzy(text, max_chars=500):
    if not text:
        return ""
    truncated = text[:max_chars]
    for marker in ["---", "From:", "-----", "Sent:", "Original Message"]:
        idx = truncated.find(marker)
        if idx > 50:
            return truncated[:idx].strip()
    return truncated.strip()

# ============================================================================
# TABLEAU AUTH
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
    except:
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
            if cleaned.startswith("SKU") or not cleaned:
                continue
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
# RMA PIPELINE (AUTHORITATIVE – TIERED MATCHING)
# ============================================================================

def run_rma_pipeline(server, wb):
    banner("STARTING RMA PIPELINE")

    sheet = wb.sheets["RMA Raw"]
    table = sheet.tables["RMA_Raw"]

    # ------------------------------------------------------------------
    # Existing Excel case numbers (dedup)
    # ------------------------------------------------------------------
    existing = set()
    if table.range.value and len(table.range.value) > 1:
        existing = {
            int(float(x)) for x in
            pd.DataFrame(
                table.range.value[1:], columns=table.range.value[0]
            )["Case Number"].dropna()
        }

    # ------------------------------------------------------------------
    # Salesforce query (SOURCE OF TRUTH FOR EXISTENCE)
    # ------------------------------------------------------------------
    sf = Salesforce(
        username=SF_USERNAME,
        password=SF_PASSWORD,
        security_token=SF_TOKEN,
        instance_url=SF_INSTANCE
    )

    start_date = (
        datetime.now(timezone.utc) - pd.DateOffset(months=2)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    soql = f"""
        SELECT
            Id,
            CaseNumber,
            RMA_Number__c,
            RMA_Value__c,
            Owner.Name,
            Description,
            Subject,
            Account.Name,
            Contact_Type__c,
            WH__c,
            Reason,
            CreatedDate,
            Item_1__c, Item_2__c, Item_3__c, Item_4__c, Item_5__c,
            Item_6__c, Item_7__c, Item_8__c, Item_9__c, Item_10__c
        FROM Case
        WHERE
            CreatedDate >= {start_date}
            AND RecordType.Name = 'RMA'
            AND Reason LIKE 'PROD%'
            AND Reason != 'Product Inquiry'
    """

    records = sf.query_all(soql)["records"]
    cases = [
        r for r in records
        if int(float(r["CaseNumber"])) not in existing
    ]

    print(f"  [{timestamp()}] Salesforce returned {len(records)} cases.")
    print(f"  [{timestamp()}] {len(cases)} new cases after dedup.")

    if not cases:
        print("✅ No new RMAs.")
        return

    # ------------------------------------------------------------------
    # Tableau RMA extract (ENRICHMENT ONLY)
    # ------------------------------------------------------------------
    view = server.views.get_by_id(TABLEAU_RMA_VIEW_ID)
    server.views.populate_csv(view)
    tdf = pd.read_csv(io.BytesIO(b"".join(view.csv)))
    tdf.columns = tdf.columns.str.strip()

    if "RMA #" in tdf.columns and "Rma Number" not in tdf.columns:
        tdf.rename(columns={"RMA #": "Rma Number"}, inplace=True)

    tdf["_rma_norm"] = tdf["Rma Number"].astype(str).apply(normalize_number)
    tdf["_issue_date"] = pd.to_datetime(tdf["Issue Date"], errors="coerce")

    # ------------------------------------------------------------------
    # Fetch product table (ONCE)
    # ------------------------------------------------------------------
    product_df, stock_codes, normalized_codes = build_product_lookup(
        fetch_product_table(server)
    )

    # ------------------------------------------------------------------
    # Collect comments (only once)
    # ------------------------------------------------------------------
    comments = rma_collect_case_comments(
        sf, [c["Id"] for c in cases]
    )

    rows_with_links = []
    stats = {"tableau_exact": 0, "tableau_contextual": 0, "product_resolved": 0, "sf_only": 0}

    # ------------------------------------------------------------------
    # CASE-DRIVEN PROCESSING
    # ------------------------------------------------------------------
    for sf_rec in cases:
        case_id = sf_rec["Id"]
        enrichment_source = "SALESFORCE_ONLY"
        tableau_row = None

        sf_rma = normalize_number(sf_rec.get("RMA_Number__c"))
        sf_customer = normalize_text(sf_rec.get("Account", {}).get("Name", ""))
        sf_wh = normalize_text(sf_rec.get("WH__c"))
        sf_date = (pd.to_datetime(sf_rec.get("CreatedDate"), errors="coerce").tz_convert(None))

        # ==========================================================
        # TIER 1 — EXACT RMA MATCH
        # ==========================================================
        if sf_rma:
            exact = tdf[tdf["_rma_norm"] == sf_rma]
            if not exact.empty:
                tableau_row = exact.iloc[0]
                enrichment_source = "TABLEAU_EXACT"

        # ==========================================================
        # TIER 2 — CONTEXTUAL MATCH (IF NO EXACT MATCH)
        # ==========================================================
        if tableau_row is None:
            candidates = []
            for _, t in tdf.iterrows():
                if not pd.notna(t["_issue_date"]) or not pd.notna(sf_date):
                    continue

                if abs((t["_issue_date"] - sf_date).days) > 3:
                    continue

                if normalize_text(t.get("Warehouse Name", "")) != sf_wh:
                    continue

                if normalize_text(t.get("Customer Name", "")) != sf_customer:
                    continue

                candidates.append(t)

            if candidates:
                product = resolve_product_info(
                    sf_rec,
                    comments.get(case_id, ""),
                    product_df,
                    stock_codes,
                    normalized_codes
                )

                if product:
                    for t in candidates:
                        t_stock = normalize_text(t.get("Stock Code", ""))
                        if normalize_text(product["stock_code"]) == t_stock:
                            tableau_row = t
                            enrichment_source = "TABLEAU_CONTEXTUAL"
                            break

        # ==========================================================
        # TIER 3 — SALESFORCE-ONLY: PRODUCT RESOLUTION FROM TABLE
        # (Only runs if no Tableau match was found)
        # ==========================================================
        product_info = None
        if tableau_row is None:
            product_info = resolve_product_info(
                sf_rec,
                comments.get(case_id, ""),
                product_df,
                stock_codes,
                normalized_codes
            )
            if product_info:
                enrichment_source = "PRODUCT_RESOLVED"

        # ==========================================================
        # BUILD EXCEL ROW (ONE PER CASE — ALWAYS)
        # ==========================================================
        qty = (
            rma_clean_number(tableau_row.get("Ordered Qty"))
            if tableau_row is not None else
            extract_quantity_from_subject(sf_rec.get("Subject", ""))
        )

        # Determine source, stock code, category, part category
        if tableau_row is not None:
            # Tableau-enriched: use Tableau data
            row_source = rma_derive_source(tableau_row)
            row_stock_code = tableau_row.get("Stock Code", "")
            row_part_category = tableau_row.get("Part Category", "")
            row_category = tableau_row.get("Category", "")
            row_enriched = "YES"
        elif product_info:
            # No Tableau match, but product table resolved it
            row_source = product_info["source"] or ""
            row_stock_code = product_info["stock_code"] or ""
            row_part_category = product_info["part_category"] or ""
            row_category = product_info["category"] or ""
            row_enriched = "YES"
        else:
            # Truly unresolved: no Tableau, no product match
            row_source = "Salesforce Only"
            row_stock_code = ""
            row_part_category = ""
            row_category = ""
            row_enriched = "NO"

        excel_row = [
            tableau_row.get("Issue Date") if tableau_row is not None
            else sf_rec.get("CreatedDate", "").split("T")[0],

            rma_parse_case_number(sf_rec["CaseNumber"]),
            tableau_row.get("Problem") if tableau_row is not None
            else sf_rec.get("Reason", ""),

            sf_rec["Owner"]["Name"] if sf_rec.get("Owner") else "",
            sf_rec.get("Description", ""),
            comments.get(case_id, ""),

            int(qty) if qty else "",
            float(sf_rec.get("RMA_Value__c")) if sf_rec.get("RMA_Value__c") else "",
            sf_rec.get("Account", {}).get("Name", ""),
            sf_rec.get("Contact_Type__c", ""),

            row_source,
            row_stock_code,
            row_part_category,
            row_category,

            tableau_row.get("Warehouse Name", "") if tableau_row is not None
            else sf_rec.get("WH__c", ""),

            timestamp(),
            row_enriched
        ]

        link = (
            f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{case_id}/view",'
            f'"{excel_row[1]}")'
        )

        rows_with_links.append((excel_row, link))

        # Track stats
        if enrichment_source == "TABLEAU_EXACT":
            stats["tableau_exact"] += 1
        elif enrichment_source == "TABLEAU_CONTEXTUAL":
            stats["tableau_contextual"] += 1
        elif enrichment_source == "PRODUCT_RESOLVED":
            stats["product_resolved"] += 1
        else:
            stats["sf_only"] += 1

    # ------------------------------------------------------------------
    # SUMMARY
    # ------------------------------------------------------------------
    print(f"  [{timestamp()}] Matching summary:")
    print(f"    Tableau exact:      {stats['tableau_exact']}")
    print(f"    Tableau contextual: {stats['tableau_contextual']}")
    print(f"    Product resolved:   {stats['product_resolved']}")
    print(f"    Salesforce only:    {stats['sf_only']}")

    # ------------------------------------------------------------------
    # WRITE TO EXCEL
    # ------------------------------------------------------------------
    if not rows_with_links:
        print(Fore.YELLOW + "⚠️ No rows to write.")
        return

    start_row = (
        table.data_body_range.last_cell.row + 1
        if table.data_body_range else table.range.row + 1
    )
    start_col = table.range.column

    table.resize(
        table.range.resize(table.range.rows.count + len(rows_with_links))
    )

    sheet.range((start_row, start_col)).value = [
        r[0] for r in rows_with_links
    ]

    case_idx = table.header_row_range.value.index("Case Number")
    for i, (_, link) in enumerate(rows_with_links):
        sheet.range((start_row + i, start_col + case_idx)).formula = link

    wb.save()
    print(Fore.GREEN + f"✅ {len(rows_with_links)} RMA rows written to Excel.")


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
# LOGGING + EXCEL LIFECYCLE (for unattended / Task Scheduler runs)
# ============================================================================

ANSI_RE = re.compile(r'\x1b\[[0-9;]*m')

class _Tee:
    """Write to multiple streams. Optionally strip ANSI color codes per stream."""

    def __init__(self, targets):
        self.targets = targets  # list of (stream, strip_ansi_bool)

    def write(self, data):
        for stream, strip in self.targets:
            try:
                stream.write(ANSI_RE.sub('', data) if strip else data)
            except Exception:
                pass

    def flush(self):
        for stream, _ in self.targets:
            try:
                stream.flush()
            except Exception:
                pass


def setup_logging():
    """Tee stdout/stderr to a dated log file next to the script.
    Console keeps color (via colorama); file gets clean, stripped text.
    Safe under pythonw.exe where there is no console.
    """
    log_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "logs"
    )
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(
        log_dir,
        f"pipeline_{datetime.now():%Y-%m-%d}.log"
    )

    logfile = open(log_path, "a", encoding="utf-8")

    logfile.write(
        f"\n{'=' * 70}\nRUN START {timestamp()}\n{'=' * 70}\n"
    )
    logfile.flush()

    console_out = sys.stdout
    console_err = sys.stderr

    out_targets = (
        ([(console_out, False)] if console_out else [])
        + [(logfile, True)]
    )

    err_targets = (
        ([(console_err, False)] if console_err else [])
        + [(logfile, True)]
    )

    sys.stdout = _Tee(out_targets)
    sys.stderr = _Tee(err_targets)

    return logfile, log_path


def open_workbook(path):
    """Attach to workbook if already open, otherwise open hidden."""
    target_full = os.path.normcase(os.path.abspath(path))
    target_name = os.path.basename(path).lower()

    for app in xw.apps:
        for bk in app.books:
            try:
                fn = bk.fullname
            except Exception:
                continue

            if (
                os.path.normcase(os.path.abspath(fn)) == target_full
                or os.path.basename(fn).lower() == target_name
            ):
                return bk, app, False

    app = xw.App(visible=False, add_book=False)

    try:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(path)
    except Exception:
        try:
            app.quit()
        except Exception:
            pass
        raise

    return wb, app, True

# ============================================================================
# Main
# ============================================================================
def main():
    logfile, log_path = setup_logging()

    clear_terminal()
    print(
        Fore.CYAN + Style.BRIGHT +
        "\n  COMBINED PIPELINE - RMA + SHIPMENTS (SINGLE RUN)"
    )
    print(f"  Log file: {log_path}\n")

    banner("CYCLE STARTING")

    server = None
    wb = None
    app = None
    we_started_app = False

    try:
        wb, app, we_started_app = open_workbook(WB_PATH)

        print(
            f"  [{timestamp()}] Workbook opened "
            f"({'new hidden Excel instance' if we_started_app else 'attached to already-open instance'})."
        )

        # Authenticate to Tableau ONCE
        print(f"  [{timestamp()}] Authenticating to Tableau...")
        server = tableau_sign_in()
        print(Fore.GREEN + f"  [{timestamp()}] Tableau authenticated.\n")

        # --- RMA Pipeline ---
        try:
            run_rma_pipeline(server, wb)
        except Exception:
            print(Fore.RED + "\n  ❌ RMA PIPELINE ERROR:\n")
            traceback.print_exc()

        # --- Shipment Pipeline ---
        try:
            run_shipment_pipeline(server, wb)
        except Exception:
            print(Fore.RED + "\n  ❌ SHIPMENT PIPELINE ERROR:\n")
            traceback.print_exc()

    except Exception:
        print(Fore.RED + "\n  ❌ CYCLE ERROR:\n")
        traceback.print_exc()

    finally:
        if server:
            tableau_sign_out(server)
            print(f"\n  [{timestamp()}] Tableau signed out.")

        # Only tear down the Excel instance WE created.
        if we_started_app and app is not None:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                traceback.print_exc()

            try:
                app.quit()
                print(f"  [{timestamp()}] Hidden Excel instance closed.")
            except Exception:
                traceback.print_exc()

        banner("CYCLE COMPLETE")
        print(f"  [{timestamp()}] Run finished.")

        try:
            logfile.flush()
            logfile.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()