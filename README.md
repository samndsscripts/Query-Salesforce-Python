# RMA Sync — Salesforce & Tableau Reporting Pipeline

Reconciling RMA and shipment activity across Salesforce, Tableau, and a shared Excel workbook used to mean manually cross-referencing three systems by hand. This pipeline pulls from all three, resolves product identity across inconsistent fields using a tiered matching strategy, deduplicates against what's already reported, and appends clean records straight into the reporting workbook.

## The Problem

RMA (Return Material Authorization) and shipment data live in Salesforce and Tableau, but neither system is the reporting source of truth — Excel is. Building that report by hand means opening cases one at a time, cross-referencing product codes that are sometimes structured and sometimes buried in a subject line or case comment, and manually checking whether a record has already been logged. It's slow and error-prone, and product identification quietly degrades whenever the structured `Item` field is missing.

## What It Does

- **RMA processing** — pulls new RMA cases from Salesforce, resolves product identity, collects case comments, and appends deduplicated records to the reporting workbook
- **Shipment processing** — pulls shipment records from Tableau, cleans and standardizes them, derives supplier/manufacturing source, and generates unique shipment identifiers
- **Tiered product resolution** — falls back through four matching strategies so a missing structured field doesn't mean a missing product match
- **Duplicate prevention** — checks incoming records against what's already in the workbook before appending anything
- **Traceability** — every appended RMA record includes a hyperlink back to its source Salesforce case

## How It Works

1. **Case retrieval** — new RMA cases are pulled from Salesforce (cases, comments, customer/product/financial fields)
2. **Product resolution** — each case runs through a tiered fallback: direct `Item` field lookup → subject-line stock code extraction → fuzzy match against the case description → fuzzy match against case comments
3. **Enrichment** — resolved cases are joined against Tableau's RMA and shipment reference data
4. **Deduplication** — candidate records are checked against existing workbook rows before being written
5. **Write-back** — new RMA and shipment rows are appended to the reporting workbook, with RMA rows hyperlinked back to their Salesforce case

```
Salesforce
     │
     ▼
Case Retrieval
     │
     ▼
Product Resolution (tiered fallback)
     │
     ├── Tableau Product Data
     ├── Tableau RMA Data
     └── Tableau Shipment Data
     │
     ▼
Deduplication Check
     │
     ▼
Excel Reporting Workbook
```

## Stack

| Layer | Tech | Why |
|---|---|---|
| Salesforce access | simple-salesforce | Straightforward REST API wrapper, no heavier SDK needed for case/comment retrieval |
| Tableau access | Tableau Server Client (TSC) | Official client for pulling published data sources without hitting the raw REST API directly |
| Data handling | Pandas | Cleaning, standardizing, and joining data across three sources |
| Product matching | RapidFuzz | Fast fuzzy matching for the fallback tiers where structured fields are missing |
| Excel I/O | xlwings | Writes directly into a live workbook (formatting, hyperlinks) rather than just dumping a flat file |
| Config | python-dotenv | Keeps Salesforce/Tableau credentials and workbook paths out of source control |
| Logging | Colorama | Readable console output for a script run on a schedule, not just interactively |

## Key Technical Decisions

**Why a tiered matching strategy instead of relying on the `Item` field?** The structured `Item` field is frequently blank or inconsistent in practice. Falling back through subject-line extraction and then fuzzy matching against description/comment text means a case still gets identified even when the clean path fails — at the cost of more matching logic to maintain.

**Why xlwings instead of writing a flat CSV/openpyxl file?** The reporting workbook is a live, shared artifact with existing formatting and formulas others depend on. xlwings can append to it in place — including writing real hyperlinks back to Salesforce — without regenerating the whole file and losing everything downstream that depends on it.

**Why check for duplicates before appending instead of relying on Salesforce/Tableau as the source of truth for "already processed"?** The workbook, not either source system, is what people actually read from. A case or shipment could be re-pulled on a later run; checking against existing workbook rows is what actually prevents double-counting in the report itself.

## Repository Structure

```
Query-Salesforce-Python/
├── README.md
├── src/
│   ├── combined_pipeline.py            # primary pipeline: Salesforce + Tableau → workbook
│   └── combined_pipeline_server_v1.py  # server-oriented variant for scheduled execution
└── archive/                            # earlier RMA/shipment workflows, kept for reference
```

## Local Setup

```
# 1. Clone and configure
cp .env.example .env
# Fill in: Salesforce credentials, Tableau credentials, workbook path

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the pipeline
python src/combined_pipeline.py
```

## What I'd Do Next

- **Replace the manual/scheduled run with an event trigger** — Salesforce supports outbound messages/platform events on new RMA cases; the pipeline is already idempotent (dedup-checked) so it could safely run per-event instead of on a fixed cycle
- **Move matching confidence into the output** — the tiered fallback currently resolves silently; surfacing which tier matched each record (direct vs. fuzzy) would let reviewers spot low-confidence matches instead of trusting every row equally
- **Replace xlwings with a database-backed report** — a live workbook works at current scale, but a proper table (even SQLite) would remove the single-writer bottleneck and make historical querying possible without opening Excel

## Configuration

Salesforce credentials, Tableau credentials, workbook paths, and other environment-specific settings are stored outside source control using `.env` files — see `.env.example` for the required keys.

## Disclaimer

This repository contains automation logic only. Credentials, company data, workbook files, and environment-specific configuration are intentionally excluded from source control for security and privacy purposes.
