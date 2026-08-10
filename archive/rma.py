import os
from dotenv import load_dotenv
import tableauserverclient as TSC
import io
import pandas as pd
import xlwings as xw
from simple_salesforce import Salesforce
from datetime import datetime, timezone
from colorama import init as colorama_init, Fore, Style

# Initialize colorama
colorama_init(autoreset=True)

# ============================================================================
# RMA AUTOMATION - FINAL VERSION (SINGLE ROW PER CASE)
# Architecture:
# Excel → Salesforce (bridge + narrative + value) → Tableau (business) → Excel
# ============================================================================

# --- Load environment variables ---
load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\salesforce.env")
load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\tableau.env")

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
TABLEAU_VIEW_ID = "128dec6d-e4ab-4843-a337-43ebad53c779"


WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"

# ============================================================================
# HELPERS
# ============================================================================

def parse_case_number(x):
    try:
        return int(float(str(x).strip()))
    except:
        return None


def clean_number(x):
    try:
        return float(str(x).replace(",", ""))
    except:
        return None


# ============================================================================
# SALESFORCE COMMENTS (REUSABLE FUNCTION)
# ============================================================================

def collect_case_comments(sf, case_ids):
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


# ============================================================================
# STEP 1: LOAD EXISTING EXCEL CASES
# ============================================================================

def load_existing_cases(wb):
    sheet = wb.sheets["RMA Raw"]
    table = sheet.tables["RMA_Raw"]

    values = table.range.value
    if len(values) <= 1:
        return set(), table, sheet

    df = pd.DataFrame(values[1:], columns=values[0])
    existing = (
        df["Case Number"]
        .apply(parse_case_number)
        .dropna()
        .astype(int)
        .tolist()
    )

    return set(existing), table, sheet


# ============================================================================
# STEP 2: SALESFORCE QUERY (BRIDGE + VALUE + NARRATIVE)
# ============================================================================

def query_salesforce(existing_cases):
    sf = Salesforce(
        username=SF_USERNAME,
        password=SF_PASSWORD,
        security_token=SF_TOKEN,
        instance_url=SF_INSTANCE
    )

    start_date = (datetime.now(timezone.utc) - pd.DateOffset(months=2)).strftime("%Y-%m-%dT%H:%M:%SZ")
   

    soql = f"""
        SELECT
            Id,
            CaseNumber,
            RMA_Number__c,
            RMA_Value__c,
            Owner.Name,
            Description,
            Account.Name,
            Contact_Type__c
        FROM Case
        WHERE
            CreatedDate >= {start_date}
            AND RecordType.Name = 'RMA'
            AND Reason LIKE 'PROD%'
    """

    records = sf.query_all(soql)["records"]

    new_records = [
        r for r in records
        if parse_case_number(r["CaseNumber"]) not in existing_cases
    ]

    case_map = {
        r["RMA_Number__c"]: r
        for r in new_records
        if r.get("RMA_Number__c")
    }

    comments = collect_case_comments(
        sf, [r["Id"] for r in new_records]
    )

    return new_records, case_map, comments


# ============================================================================
# STEP 3: TABLEAU DATA (FILTER NULL ISSUE DATE)
# ============================================================================

def fetch_tableau_data(rma_numbers):
    auth = TSC.PersonalAccessTokenAuth(
        TABLEAU_PAT_NAME, TABLEAU_PAT_SECRET, site_id=TABLEAU_SITE
    )
    server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

    with server.auth.sign_in(auth):
        view = server.views.get_by_id(TABLEAU_VIEW_ID)
        server.views.populate_csv(view)
        df = pd.read_csv(io.BytesIO(b"".join(view.csv)))

    # Normalize RMA
    df["_rma"] = df["Rma Number"].astype(str).str.strip()

    # ✅ Keep only RMAs from Salesforce
    df = df[df["_rma"].isin(set(map(str, rma_numbers)))]

    # ✅ DROP NULL Issue Date rows
    df = df[pd.notna(df["Issue Date"])].copy()

    return df
# ============================================================================
# SOURCE DERIVATION
# ============================================================================

def derive_source(row):
    pc = str(row.get("Part Category", "")).upper()
    if pc == "M":
        return str(row.get("Mfg Plant Name", ""))
    if pc == "B":
        return str(row.get("Supplier", ""))
    return ""


# ============================================================================
# STEP 4: MERGE + WRITE
# ============================================================================

def write_to_excel(case_map, comments_map, tableau_df, table, sheet, wb):
    start_row = table.data_body_range.last_cell.row + 1
    start_col = table.range.column

    rows = []

    for _, t in tableau_df.iterrows():
        rma = str(t["Rma Number"]).strip()
        sf = case_map.get(rma)
        if not sf:
            continue

        case_id = sf["Id"]
        case_num = parse_case_number(sf["CaseNumber"])

        qty = clean_number(t.get("Ordered Qty"))

        excel_row = [
            t.get("Issue Date"),
            case_num,
            t.get("Problem"),
            sf["Owner"]["Name"] if sf.get("Owner") else "",
            sf.get("Description", ""),
            comments_map.get(case_id, ""),
            int(qty) if qty else "",
            float(sf.get("RMA_Value__c")) if sf.get("RMA_Value__c") else "",
            sf.get("Account", {}).get("Name", ""),
            sf.get("Contact_Type__c", ""),
            derive_source(t),
            t.get("Stock Code", ""),
            t.get("Part Category", ""),
            t.get("Category", ""),
            t.get("Warehouse Name", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "YES"
        ]

        link = f'=HYPERLINK("{SF_INSTANCE}/lightning/r/Case/{case_id}/view","{case_num}")'
        rows.append((excel_row, link))

    if not rows:
        print(Fore.YELLOW + "⚠️ No rows to write.")
        return

    table.resize(table.range.resize(table.range.rows.count + len(rows)))
    sheet.range((start_row, start_col)).value = [r[0] for r in rows]

    case_idx = table.header_row_range.value.index("Case Number")
    for i, (_, link) in enumerate(rows):
        sheet.range((start_row + i, start_col + case_idx)).formula = link

    wb.save()


# ============================================================================
# MAIN
# ============================================================================

def main():
    print(Fore.GREEN + Style.BRIGHT + "\nRMA AUTOMATION - FINAL\n")

    wb = xw.Book(WB_PATH)

    existing, table, sheet = load_existing_cases(wb)

    sf_records, case_map, comments = query_salesforce(existing)
    if not sf_records:
        print("✅ No new cases.")
        return

    rmas = list(case_map.keys())
    tableau_df = fetch_tableau_data(rmas)

    write_to_excel(case_map, comments, tableau_df, table, sheet, wb)

    print(Fore.GREEN + "✅ SUCCESS - Rows written to Excel.")


if __name__ == "__main__":
    main()
