import time
import csv
from simple_salesforce import Salesforce
import json
from tabulate import tabulate
from openpyxl import load_workbook
import pandas as pd
import xlwings as xw
import re
from collections import Counter
from rapidfuzz import process

# ----- Salesforce login -----
sf = Salesforce(
    username='samuelcooper@ndspro.com',
    password='Summer@NDS2025',
    security_token='zjU2IJAfQmx6zDxgOj3aLkyPQ',
    instance_url='https://nds.my.salesforce.com'
)
#--- Salesforce report fetch ---
report_id = "00OUI00000EsGR72AN"  # or whichever one you want

try:
    report_data = sf.restful(f'analytics/reports/{report_id}', params={'includeDetails': 'true'})
    print("✅ Successfully fetched report data!")
    print("Report name:", report_data.get('reportMetadata', {}).get('name', 'Unknown'))
    print("Number of rows:", len(report_data.get('factMap', {}).get('T!T', {}).get('rows', [])))
except Exception as e:
    print("❌ Error fetching report:", e)

report_rows = report_data.get('factMap', {}).get('T!T', {}).get('rows', [])

# --- Open workbook and read supplier table ---
wb = xw.Book(r"C:\Users\emp35107\OneDrive - NORMA Group\Documents\PPMs.xlsx")
sheet = wb.sheets['Class Site Supplier']
table_range = sheet.tables['Table3'].range
supplier_df = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])

# Extract lists for matching
stock_codes_list = supplier_df['Stock Code'].dropna().astype(str).tolist()
man_sites_list   = supplier_df['Mfg Plant Name'].dropna().astype(str).tolist()
suppliers_list   = supplier_df['Supplier'].dropna().astype(str).tolist()

# --- Define RMA table columns ---
columns = [
    "Opened Date",
    "Case Reason",
    "Case Owner",
    "Case Number",
    "Description",
    "Quantity",
    "RMA Value",
    "Case Category",
    "Account Name",
    "Comments",
    "Contact Type",
    "Shipping Whse",
    "Source"
]

rma_df = pd.DataFrame(columns=columns)

# --- Helper function to normalize text ---
def normalize(text):
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# --- Loop through Salesforce report rows ---
for row in report_rows:
    cells = row.get('dataCells', [])
    description = str(cells[4].get('label', ''))
    
    # --- RapidFuzz matching to determine Source ---
    normalized_description = normalize(description)
    normalized_stock_codes = [normalize(code) for code in stock_codes_list]
    
    matches = process.extract(normalized_description, normalized_stock_codes, limit=3)
    index_matches = [normalized_stock_codes.index(m[0]) for m in matches]
    
    matched_sources = []
    for i in index_matches:
        if i < len(suppliers_list) and suppliers_list[i].strip():
            matched_sources.append(suppliers_list[i])
        else:
            matched_sources.append(man_sites_list[i] if i < len(man_sites_list) else "N/A")
    
    # Most likely source among top 3 matches
    source_counter = Counter(matched_sources)
    most_common_source = source_counter.most_common(1)[0][0]
    
    # --- Map Salesforce report columns to RMA table ---
    new_row = {
        "Opened Date": cells[0].get('label', ''),
        "Case Reason": cells[1].get('label', ''),
        "Case Owner": cells[2].get('label', ''),
        "Case Number": cells[3].get('label', ''),
        "Description": description,
        "Quantity": '',  # Not in report
        "RMA Value": cells[5].get('label', ''),
        "Case Category": cells[6].get('label', ''),
        "Account Name": cells[7].get('label', ''),
        "Comments": '',  # Not in report
        "Contact Type": cells[8].get('label', ''),
        "Shipping Whse": cells[9].get('label', ''),
        "Source": most_common_source  # <- Filled automatically
    }
    
    rma_df = pd.concat([rma_df, pd.DataFrame([new_row])], ignore_index=True)

# Preview first 10 rows
print(rma_df.head(10))