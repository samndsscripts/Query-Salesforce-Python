# Salesforce & Tableau Data Synchronization Pipeline

## Overview

This Python automation combines Salesforce, Tableau, and Excel into a single reporting workflow. The project retrieves RMA (Return Material Authorization) and shipment data, enriches records using multiple matching strategies, prevents duplicates, and updates centralized reporting workbooks.

The automation is designed to reduce manual reporting effort while improving data quality and consistency across business systems.

---

## Features

### RMA Processing

- Retrieves new RMA cases from Salesforce.
- Prevents duplicate processing using existing Excel records.
- Pulls RMA enrichment data from Tableau.
- Collects Salesforce case comments.
- Resolves product information using direct and fuzzy matching techniques.
- Creates hyperlinks back to Salesforce records.
- Appends processed data into reporting tables.

### Shipment Processing

- Retrieves shipment records from Tableau.
- Cleans and standardizes source data.
- Derives supplier and manufacturing source information.
- Generates unique shipment identifiers.
- Prevents duplicate entries.
- Appends new shipment activity to Excel reporting tables.

### Product Resolution

The automation uses a tiered matching strategy:

1. Item field lookup
2. Subject line stock code extraction
3. Fuzzy matching against case descriptions
4. Fuzzy matching against case comments

This approach improves product identification when structured data is incomplete.

---

## Data Sources

### Salesforce

Used for:

- RMA Cases
- Case Comments
- Customer Information
- Product Data
- Financial Information

### Tableau

Used for:

- RMA Enrichment Data
- Shipment Data
- Product Reference Tables

### Excel

Used as the final reporting and tracking repository.

---

## Repository Structure

```text
Query-Salesforce-Python/
│
├── README.md
│
├── src/
│   ├── combined_pipeline.py
│   └── combined_pipeline_server_v1.py
│
├── archive/
│   ├── rma.py
│   ├── shipments.py
│   ├── report_automation_9.py
│   ├── report_automation_10.py
│   ├── report_automation_11.py
│   └── report_automation_12.py
│
└── miscellaneous utility scripts
```

---

## Active Scripts

### `/src/combined_pipeline.py`

Primary production pipeline that:

- Connects to Salesforce
- Connects to Tableau
- Processes new RMAs
- Processes shipment data
- Updates reporting workbooks
- Runs on a scheduled cycle

### `/src/combined_pipeline_server_v1.py`

Server-oriented implementation of the combined pipeline designed for automated execution environments.

---

## Archive

The `/archive` folder contains previous versions and legacy components retained for reference purposes:

- Earlier RMA workflows
- Earlier shipment workflows
- Historical report automation versions
- Development and testing scripts

These files are not actively maintained.

---

## Automation Workflow

```text
Salesforce
     │
     ▼
Case Retrieval
     │
     ▼
Product Resolution
     │
     ├── Tableau Product Data
     ├── Tableau RMA Data
     └── Tableau Shipment Data
     │
     ▼
Data Enrichment
     │
     ▼
Excel Reporting Workbook
```

---

## Technologies Used

- Python
- Pandas
- xlwings
- Tableau Server Client (TSC)
- Simple Salesforce
- RapidFuzz
- python-dotenv
- Colorama

---

## Version Control Workflow

This repository is maintained across multiple development machines.

Before starting work:

```bash
git pull
```

After making changes:

```bash
git add .
git commit -m "Description of changes"
git push
```

Check repository status:

```bash
git status
```

---

## Business Benefits

- Reduces manual reporting effort
- Improves data consistency
- Consolidates multiple enterprise systems
- Automates data enrichment
- Maintains a deduplicated dataset
- Provides traceability back to Salesforce records
- Delivers near real-time reporting updates

---

## Configuration

Sensitive information such as:

- Salesforce credentials
- Tableau credentials
- Workbook paths
- Environment variables

is stored outside source control using `.env` files.

---

## Disclaimer

This repository contains automation logic only. Credentials, company data, workbook files, and environment-specific configuration are intentionally excluded from source control for security and privacy purposes.