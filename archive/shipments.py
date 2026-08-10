import os
import io
import pandas as pd
import xlwings as xw
from datetime import datetime
from dotenv import load_dotenv
import tableauserverclient as TSC

# ============================================================================
# CONFIG
# ============================================================================

load_dotenv(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\tableau.env")

WB_PATH = r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx"

TABLEAU_SERVER = os.getenv("TABLEAU_SERVER")
TABLEAU_PAT_NAME = os.getenv("TABLEAU_PAT_NAME")
TABLEAU_PAT_SECRET = os.getenv("TABLEAU_PAT_SECRET")
TABLEAU_SITE = os.getenv("TABLEAU_SITE_CONTENTURL", "")
TABLEAU_SHIPMENT_VIEW_ID = "41b7e54e-3da8-4b53-bfaf-047410bb4fd8"

# ============================================================================
# EXCEL
# ============================================================================

wb = xw.Book(WB_PATH)
sheet = wb.sheets["Shipment Data"]
table = sheet.tables["shipment_data"]

# ============================================================================
# HELPERS (LOCKED)
# ============================================================================

def clean_val(val):
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


def derive_source(row):
    part_category = str(row.get("Part Category", "")).strip().upper()
    if part_category == "M":
        return str(row.get("Mfg Plant Name", "")).strip().upper()
    if part_category == "B":
        return str(row.get("Supplier", "")).strip().upper()
    return ""


def generate_ident(row):
    return (
        clean_val(row.get("Mfg Plant Name")) + '|' +
        clean_val(row.get("Source")) + '|' +
        clean_val(row.get("Order Qty")) + '|' +
        clean_val(row.get("Year")) + '|' +
        clean_val(row.get("Month"))
    )

# ============================================================================
# TABLEAU INGEST
# ============================================================================

def fetch_tableau_shipment_data():
    auth = TSC.PersonalAccessTokenAuth(
        TABLEAU_PAT_NAME,
        TABLEAU_PAT_SECRET,
        site_id=TABLEAU_SITE
    )
    server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

    with server.auth.sign_in(auth):
        view = server.views.get_by_id(TABLEAU_SHIPMENT_VIEW_ID)
        server.views.populate_csv(view)
        df = pd.read_csv(io.BytesIO(b"".join(view.csv)))

    df.columns = df.columns.str.strip()
    return df

# ============================================================================
# MAIN PIPELINE
# ============================================================================

def process_tableau_shipment_data():
    df = fetch_tableau_shipment_data()
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
    df["Source"] = df.apply(derive_source, axis=1)

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
    df["Ident"] = df.apply(generate_ident, axis=1)

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

    print(f"✅ {len(out)} shipment rows appended.")

# ============================================================================
# IMPORT‑SAFE ENTRY POINT
# ============================================================================

def run():
    process_tableau_shipment_data()


if __name__ == "__main__":
    run()
