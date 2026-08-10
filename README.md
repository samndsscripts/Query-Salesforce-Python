# Salesforce & Tableau Data Synchronization Pipeline

## Overview

This Python automation combines Salesforce, Tableau, and Excel into a single reporting workflow. The script continuously monitors new RMA (Return Material Authorization) cases and shipment records, enriches data through multiple matching methods, prevents duplicates, and automatically updates a centralized Excel workbook used for business reporting.

The solution is designed to reduce manual data entry, improve data quality, and provide a consistent reporting source across multiple systems.

---

## Features

### RMA Automation

The RMA pipeline:

- Retrieves newly created RMA cases from Salesforce
- Prevents duplicate processing by comparing against existing Excel records
- Pulls supporting RMA data from Tableau
- Collects Salesforce case comments for additional context
- Enriches records using product reference data
- Creates direct hyperlinks back to Salesforce cases
- Writes processed records into an Excel reporting table

### Product Resolution Logic

To maximize data accuracy, the script uses a tiered matching strategy to identify products associated with RMA cases.

Matching attempts occur in the following order:

1. Direct lookup from Salesforce item fields
2. Stock code extraction from case subjects
3. Fuzzy matching against case descriptions
4. Fuzzy matching against case comments

This approach allows product information to be identified even when structured data is incomplete.

### Shipment Automation

The shipment pipeline:

- Retrieves shipment data from Tableau
- Cleans and standardizes source data
- Derives supplier or manufacturing source information
- Generates unique shipment identifiers
- Removes duplicate records
- Appends only new shipment transactions to Excel

---

## Data Sources

### Salesforce

Used for:

- RMA case records
- Case comments
- Customer information
- Product-related fields
- Financial values

### Tableau

Used for:

- RMA enrichment data
- Shipment transaction data
- Product reference lookup tables

### Excel

Used as the centralized reporting repository where processed data is stored for business analysis.

---

## Data Enrichment Workflow

### RMA Matching

The script attempts to enrich each RMA using the following hierarchy:

#### Tier 1: Exact Match

Matches Salesforce RMA numbers directly against Tableau records.

#### Tier 2: Contextual Match

When an exact match cannot be found, the script compares:

- Customer
- Warehouse
- Date proximity
- Product information

to identify likely Tableau records.

#### Tier 3: Product Resolution

If no Tableau match exists, product information is resolved directly from the product reference table using exact and fuzzy matching techniques.

#### Tier 4: Salesforce Only

If no enrichment data can be found, the case is still recorded and flagged for review.

---

## Duplicate Prevention

### RMA Records

Duplicates are prevented using Salesforce case numbers already stored in Excel.

### Shipment Records

Unique shipment identifiers are generated using:

- Manufacturing location
- Source
- Quantity
- Year
- Month

Only records that do not already exist are added.

---

## Automation Cycle

The application operates continuously on a scheduled interval.

Each cycle performs the following tasks:

1. Connect to Tableau
2. Open the reporting workbook
3. Process new RMAs
4. Process new shipments
5. Save Excel updates
6. Disconnect from Tableau
7. Wait for the next execution cycle

---

## Technologies Used

- Python
- Pandas
- xlwings
- Tableau Server Client (TSC)
- simple-salesforce
- RapidFuzz
- python-dotenv
- Colorama

---

## Project Structure

```text
project/
│
├── main.py
├── README.md
├── requirements.txt
├── .gitignore
└── .env
```

---

## Business Benefits

- Eliminates repetitive manual reporting tasks
- Reduces data entry errors
- Consolidates information from multiple systems
- Improves product identification accuracy
- Maintains a deduplicated reporting dataset
- Provides traceability back to Salesforce records
- Delivers near real-time reporting updates

---

## Configuration

This project uses environment variables to store sensitive credentials and connection details.

Configuration items include:

- Salesforce authentication
- Tableau authentication
- Workbook location
- Scheduled execution interval

Credentials and company-specific configuration files should not be committed to source control.

---

## Disclaimer

This repository contains automation logic only. All credentials, environment files, workbook paths, Salesforce data, Tableau resources, and company-specific information have been excluded for security and privacy purposes.
