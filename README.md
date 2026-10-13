# Payment and CRM sync pipeline

A staged ETL pipeline that compares payment transactions with CRM records, prepares missing opportunities and validates updates before loading.

**Stack:** Python, pandas, BigQuery, Salesforce.

## What this demonstrates

Reusable implementation patterns for data integration, validation and maintainable automation. This is a curated, generalized portfolio edition. It contains no production exports, deployment history, internal access inventory or credentials. API names and source labels have been generalized where needed; configure the matching schema in your own sandbox before use.

## Getting started

Install Python dependencies with `python -m pip install -r requirements.txt`. The scripts are sequential notebook-style stages. Run from the project root with your own input workbooks under ignored `data/input/`. Configure `SF_ALIAS`, `GCP_PROJECT_ID` and `SF_SERVICE_OWNER_ID` for your sandbox. Loading scripts default to `DRY_RUN = True`; inspect the preflight before enabling writes. 

## Source map

- `README.md`
- `requirements.txt`
- `scripts/build_opps_to_create.py`
- `scripts/compare_paymenthub_salesforce.py`
- `scripts/cross_paymenthub_excel.py`
- `scripts/query_paymenthub_gcp.py`
- `scripts/query_salesforce.py`
- `scripts/update_dr_status_from_paymenthub.py`
- `scripts/update_loaded_opps_owner.py`
- `scripts/upload_to_salesforce.py`

## Validation

Python sources and JSON/XML configuration are checked locally. Apex test classes are retained where present; executing them requires a compatible sandbox. No Salesforce query, deployment, backfill or record mutation was performed during preparation. See [SECURITY.md](SECURITY.md) for data handling.
