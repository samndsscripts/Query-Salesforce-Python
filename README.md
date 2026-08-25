# RMA Sync — Salesforce & Tableau Reporting Pipeline

RMA and shipment data live in two systems that each tell half the story — Salesforce is where the sales team logs the human side of a return (comments, photos, files), Tableau is connected directly to the database and holds the actual shipment/return quantities. This pipeline pulls from both, resolves product identity across inconsistent fields using a tiered matching strategy, and appends clean records into a shared workbook that quality managers actively review and correct — without stepping on the corrections they've already made.

## The Problem

Building the RMA/shipment report by hand means opening cases one at a time in Salesforce, cross-referencing product codes and quantities against Tableau (which is sometimes a clean match, sometimes not — Salesforce cases can exist before or without a matching database record), and manually checking whether a record has already been logged. On top of that, the resulting report isn't just read passively — it feeds a Power BI dashboard and is where quality managers actually work: filtering to their cases, clicking through to the source Salesforce case, correcting quantities when the pipeline had to fall back to logic instead of a direct Tableau match, and disapproving cases entirely so they don't count toward KPIs. Any automation here has to respect that this workbook has humans actively editing it, not just append blindly.

## What It Does

- **RMA processing** — pulls new RMA cases from Salesforce, resolves product identity, collects case comments, and appends new records to the shared workbook
- **Shipment processing** — pulls shipment records from Tableau (database-backed), cleans and standardizes them, derives supplier/manufacturing source, and generates unique shipment identifiers
- **Tiered product resolution** — falls back through four matching strategies so a case without a clean database-side match still gets identified
- **Duplicate-safe appending** — checks the workbook for existing records before appending, so re-runs don't create duplicate rows
- **Preserves manager review** — quantity corrections and case disapprovals made directly in the workbook are not touched or overwritten by later pipeline runs
- **Traceability** — every appended RMA record includes a hyperlink back to its source Salesforce case, so managers can review without leaving the sheet

## How It Works

1. **Case retrieval** — new RMA cases are pulled from Salesforce (cases, comments, customer/product/financial fields)
2. **Product resolution** — each case runs through a tiered fallback: direct `Item` field lookup → subject-line stock code extraction → fuzzy match against the case description → fuzzy match against case comments
3. **Enrichment** — resolved cases are joined against Tableau's database-backed RMA and shipment data
4. **Existing-record check** — candidate records are checked against the workbook so already-logged cases aren't re-appended, and rows a manager has already corrected or disapproved are left alone
5. **Write-back** — new RMA and shipment rows are appended, with RMA rows hyperlinked back to their Salesforce case, ready for quality manager review
6. **Downstream** — the workbook feeds a Power BI dashboard; managers filter to their own cases, verify or correct quantities, and can disapprove a case to exclude it from KPI reporting entirely

```
Salesforce (human side: cases, comments, photos)   Tableau (database-backed: quantities, shipments)
                    │                                              │
                    ▼                                              │
         Product Resolution (tiered fallback) ◄────────────────────┘
                    │
                    ▼
       Existing-Record Check (skip already-logged / manager-edited rows)
                    │
                    ▼
         Excel Workbook  ──────►  Power BI Dashboard
                    │
                    ▼
   Quality Manager Review (correct quantities, disapprove cases from KPIs)
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

**Why check the workbook itself for existing records, rather than tracking "already processed" separately?** The workbook isn't a passive output — quality managers actively edit it (correcting quantities, disapproving cases), and those edits have to survive the next pipeline run untouched. Checking against the workbook directly, rather than a separate log, is what lets the pipeline tell "already here, leave it" apart from "new, append it" without a second system to keep in sync.

**Why does it matter that the pipeline never overwrites a manager's quantity correction or disapproval?** Those edits exist specifically because the automated match wasn't reliable enough on its own — a quantity derived from fallback logic instead of a direct Tableau match, or a case a manager determined shouldn't count toward KPIs. If a later run silently overwrote that correction, the KPI numbers downstream in Power BI would quietly become wrong again, defeating the entire reason the review step exists.

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
