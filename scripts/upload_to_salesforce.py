# %%
import json
import os
import re
import subprocess
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from simple_salesforce import Salesforce

# %%
# =========================
# 1. CONFIGURATION
# =========================

DRY_RUN = True

# Optional safety limit.
# Use None to process all rows.
# Example: MAX_RECORDS_TO_PROCESS = 5 for a small first test.
MAX_RECORDS_TO_PROCESS = None

BATCH_SIZE = 100

# %%
# =========================
# 2. PROJECT PATHS
# =========================

BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

source_path = OUTPUT_DIR / "06_opportunities_to_create_review.xlsx"

source_path

# %%
# =========================
# 3. LOAD ENVIRONMENT
# =========================

load_dotenv(BASE_DIR / ".env", override=True)

SF_ALIAS = os.getenv("SF_ALIAS", "DEMO_SANDBOX")
SF_INSTANCE_URL = os.getenv("SF_INSTANCE_URL")
SF_CLI_PATH = os.getenv("SF_CLI_PATH", "sf")

SF_ALIAS, SF_INSTANCE_URL, SF_CLI_PATH

# %%
# =========================
# 4. SALESFORCE CLI CONNECTION HELPERS
# =========================

ANSI_ESCAPE_RE = re.compile(
    r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])"
)


def strip_ansi_codes(text: str) -> str:
    return ANSI_ESCAPE_RE.sub("", text)


def parse_cli_output_to_dict(stdout: str) -> dict:
    stdout = strip_ansi_codes(stdout).strip()

    if not stdout:
        raise ValueError("Salesforce CLI returned empty STDOUT.")

    json_start = stdout.find("{")
    json_end = stdout.rfind("}")

    if json_start == -1 or json_end == -1:
        raise ValueError(
            "Could not find JSON object in Salesforce CLI output. "
            f"First 300 chars: {stdout[:300]!r}"
        )

    candidate = stdout[json_start:json_end + 1]

    return json.loads(candidate)


def get_salesforce_connection_from_cli(alias: str, sf_cli_path: str = "sf") -> Salesforce:
    command_env = os.environ.copy()
    command_env["NO_COLOR"] = "1"
    command_env["FORCE_COLOR"] = "0"
    command_env["TERM"] = "dumb"

    result = subprocess.run(
        [
            sf_cli_path,
            "org",
            "display",
            "--target-org",
            alias,
            "--json",
            "--verbose",
        ],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
        errors="replace",
        env=command_env,
    )

    if result.returncode != 0:
        raise RuntimeError(
            "Salesforce CLI failed.\n\n"
            f"Return code: {result.returncode}\n\n"
            f"STDOUT:\n{result.stdout}\n\n"
            f"STDERR:\n{result.stderr}"
        )

    payload = parse_cli_output_to_dict(result.stdout)
    org_info = payload["result"]

    access_token = org_info.get("accessToken")
    instance_url = org_info.get("instanceUrl")

    if not access_token:
        raise ValueError("No accessToken returned by Salesforce CLI.")

    if not instance_url:
        raise ValueError("No instanceUrl returned by Salesforce CLI.")

    return Salesforce(
        instance_url=instance_url,
        session_id=access_token,
    )


# %%
# =========================
# 5. CONNECT TO SALESFORCE
# =========================

sf = get_salesforce_connection_from_cli(SF_ALIAS, SF_CLI_PATH)

test_result = sf.query("""
SELECT Id, Name
FROM Organization
LIMIT 1
""")

test_result["totalSize"]

# %%
# =========================
# 6. LOAD INSERT DRAFT
# =========================

opps_to_insert_raw = pd.read_excel(
    source_path,
    sheet_name="opps_to_insert_final_draft",
    dtype=str
)

opps_to_insert_raw.shape

# %%
# =========================
# 7. CLEAN INSERT DRAFT
# =========================

opps_to_insert = opps_to_insert_raw.copy()

# Normalize empty-like values.
opps_to_insert = opps_to_insert.replace({
    "nan": pd.NA,
    "None": pd.NA,
    "NaT": pd.NA,
    "": pd.NA,
})

# Strip string columns.
for col in opps_to_insert.columns:
    if opps_to_insert[col].dtype == "object":
        opps_to_insert[col] = opps_to_insert[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )

# Convert dates and numbers.
opps_to_insert["CloseDate"] = pd.to_datetime(
    opps_to_insert["CloseDate"],
    errors="coerce"
).dt.date

opps_to_insert["Amount"] = pd.to_numeric(
    opps_to_insert["Amount"],
    errors="coerce"
)

opps_to_insert["npsp__Recurring_Donation_Installment_Number__c"] = pd.to_numeric(
    opps_to_insert["npsp__Recurring_Donation_Installment_Number__c"],
    errors="coerce"
)

# Optional limit for test batches.
if MAX_RECORDS_TO_PROCESS is not None:
    opps_to_insert = opps_to_insert.head(MAX_RECORDS_TO_PROCESS).copy()

opps_to_insert.shape

# %%
# =========================
# 8. VALIDATE CRITICAL FIELDS
# =========================

critical_fields = [
    "Name",
    "CloseDate",
    "StageName",
    "RecordTypeId",
    "Amount",
    "CurrencyIsoCode",
    "Auth__c",
    "npe03__Recurring_Donation__c",
    "npsp__Recurring_Donation_Installment_Number__c",
    "OwnerId",
]

critical_missing = opps_to_insert[
    opps_to_insert[critical_fields].isna().any(axis=1)
].copy()

critical_missing.shape

# %%
# =========================
# 9. VALIDATE STAGENAME VALUES
# =========================

allowed_stage_names = {
    "Closed Won",
    "Reembolso",
}

invalid_stage_names = opps_to_insert[
    ~opps_to_insert["StageName"].isin(allowed_stage_names)
].copy()

opps_to_insert["StageName"].value_counts(dropna=False), invalid_stage_names.shape

# %%
# =========================
# 10. VALIDATE AUTH UNIQUENESS INSIDE FILE
# =========================

auth_duplicates_in_file = (
    opps_to_insert
    .groupby("Auth__c", dropna=False)
    .size()
    .reset_index(name="records")
    .query("records > 1")
)

auth_duplicates_in_file.shape

# %%
# =========================
# 11. QUERY SALESFORCE FOR EXISTING AUTH__C
# =========================

def soql_escape(value: str) -> str:
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def chunk_list(values: list, chunk_size: int = 100) -> list[list]:
    return [
        values[i:i + chunk_size]
        for i in range(0, len(values), chunk_size)
    ]


auth_values = (
    opps_to_insert["Auth__c"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

existing_auth_records = []

for chunk in chunk_list(auth_values, 100):
    auth_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    existing_query = f"""
    SELECT
        Id,
        Name,
        Auth__c,
        npe03__Recurring_Donation__c,
        CloseDate,
        StageName,
        Amount
    FROM Opportunity
    WHERE Auth__c IN ({auth_list_soql})
    """

    result = sf.query_all(existing_query)
    existing_auth_records.extend(result["records"])

existing_auth_df = pd.DataFrame(existing_auth_records)

if not existing_auth_df.empty and "attributes" in existing_auth_df.columns:
    existing_auth_df = existing_auth_df.drop(columns=["attributes"])

existing_auth_df.shape

# %%
# =========================
# 12. EXCLUDE RECORDS THAT ALREADY EXIST IN SALESFORCE
# =========================

existing_auth_set = set()

if not existing_auth_df.empty:
    existing_auth_set = set(
        existing_auth_df["Auth__c"]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

opps_to_insert["already_exists_in_salesforce_by_auth"] = (
    opps_to_insert["Auth__c"]
    .astype(str)
    .str.strip()
    .isin(existing_auth_set)
)

opps_blocked_existing_auth = opps_to_insert[
    opps_to_insert["already_exists_in_salesforce_by_auth"]
].copy()

opps_ready_for_insert = opps_to_insert[
    ~opps_to_insert["already_exists_in_salesforce_by_auth"]
].copy()

opps_ready_for_insert.shape, opps_blocked_existing_auth.shape

# %%
# =========================
# 13. STOP IF VALIDATION FAILS
# =========================

validation_errors = []

if not critical_missing.empty:
    validation_errors.append("critical_missing is not empty")

if not invalid_stage_names.empty:
    validation_errors.append("invalid_stage_names is not empty")

if not auth_duplicates_in_file.empty:
    validation_errors.append("auth_duplicates_in_file is not empty")

if not opps_blocked_existing_auth.empty:
    print(
        f"Warning: {len(opps_blocked_existing_auth)} records already exist in Salesforce by Auth__c. "
        "They will be excluded from insert."
    )

validation_errors

# %%
# =========================
# 14. PREPARE SALESFORCE RECORDS
# =========================

def clean_value_for_salesforce(value):
    if pd.isna(value):
        return None

    if hasattr(value, "isoformat"):
        return value.isoformat()

    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def row_to_salesforce_record(row: pd.Series, columns: list[str]) -> dict:
    record = {}

    for col in columns:
        value = clean_value_for_salesforce(row[col])

        # Do not send null optional fields.
        # Keep this conservative: omit None values.
        if value is not None:
            record[col] = value

    return record


insert_columns = [
    "Name",
    "npsp__Primary_Contact__c",
    "CloseDate",
    "StageName",
    "RecordTypeId",
    "Amount",
    "CurrencyIsoCode",
    "CampaignId",
    "Campa_a_de_origen__c",
    "Canal__c",
    "Fecha_de_captacion__c",
    "Campana_web__c",
    "Fuente__c",
    "Medio__c",
    "Auth__c",
    "Medio_de_Pago__c",
    "N_mero_de_Tarjeta__c",
    "Fecha_de_vencimiento_TC__c",
    "npe03__Recurring_Donation__c",
    "npsp__Recurring_Donation_Installment_Number__c",
    "OwnerId",
]

records_to_insert = [
    row_to_salesforce_record(row, insert_columns)
    for _, row in opps_ready_for_insert.iterrows()
]

len(records_to_insert)

# %%
# =========================
# 15. EXPORT PREFLIGHT REVIEW
# =========================

preflight_output_path = OUTPUT_DIR / "07_salesforce_insert_preflight.xlsx"

with pd.ExcelWriter(preflight_output_path, engine="openpyxl") as writer:
    opps_to_insert_raw.to_excel(writer, sheet_name="source_insert_draft", index=False)
    opps_to_insert.to_excel(writer, sheet_name="cleaned_insert_draft", index=False)
    opps_ready_for_insert.to_excel(writer, sheet_name="ready_for_insert", index=False)
    opps_blocked_existing_auth.to_excel(writer, sheet_name="blocked_existing_auth", index=False)
    existing_auth_df.to_excel(writer, sheet_name="existing_auth_in_sf", index=False)
    critical_missing.to_excel(writer, sheet_name="critical_missing", index=False)
    invalid_stage_names.to_excel(writer, sheet_name="invalid_stage_names", index=False)
    auth_duplicates_in_file.to_excel(writer, sheet_name="auth_duplicates_in_file", index=False)
    pd.DataFrame({"validation_errors": validation_errors}).to_excel(writer, sheet_name="validation_errors", index=False)

print(f"Preflight file created: {preflight_output_path}")

# %%
# =========================
# 16. DRY RUN SAFETY CHECK
# =========================

print("DRY_RUN:", DRY_RUN)
print("Rows in source:", len(opps_to_insert))
print("Ready for insert:", len(opps_ready_for_insert))
print("Blocked existing Auth__c:", len(opps_blocked_existing_auth))
print("Validation errors:", validation_errors)

# %%
# =========================
# 17. INSERT INTO SALESFORCE
# =========================

insert_results = []

if DRY_RUN:
    print("DRY_RUN is True. No records were inserted.")
else:
    if validation_errors:
        raise RuntimeError(
            "Insert blocked due to validation errors: "
            + "; ".join(validation_errors)
        )

    if not records_to_insert:
        raise RuntimeError("No records to insert.")

    insert_results = sf.bulk.Opportunity.insert(
        records_to_insert,
        batch_size=BATCH_SIZE,
        use_serial=True,
    )

    print(f"Inserted attempt count: {len(insert_results)}")

# %%
# =========================
# 18. EXPORT INSERT RESULTS
# =========================

insert_results_df = pd.DataFrame(insert_results)

if DRY_RUN:
    insert_results_df = pd.DataFrame({
        "dry_run": [True],
        "message": ["No records inserted."],
        "records_ready_for_insert": [len(records_to_insert)],
    })

insert_results_output_path = OUTPUT_DIR / "08_salesforce_insert_results.xlsx"

with pd.ExcelWriter(insert_results_output_path, engine="openpyxl") as writer:
    opps_ready_for_insert.to_excel(writer, sheet_name="attempted_records", index=False)
    insert_results_df.to_excel(writer, sheet_name="insert_results", index=False)

print(f"Insert results file created: {insert_results_output_path}")

# %%
# =========================
# 19. SUMMARIZE INSERT RESULTS
# =========================

if not DRY_RUN and not insert_results_df.empty:
    insert_results_df["success"].value_counts(dropna=False)
else:
    insert_results_df

# %%
# =========================
# OPTIONAL. VERIFY INSERTED RECORDS BY ID
# =========================
# Use this block only after an insert.
# Requires insert_results_df to contain Salesforce inserted IDs.

# inserted_ids = insert_results_df["id"].dropna().tolist()
#
# verify_query = f"""
# SELECT
#     Id,
#     Name,
#     CloseDate,
#     StageName,
#     Amount,
#     CurrencyIsoCode,
#     Auth__c,
#     npe03__Recurring_Donation__c,
#     npsp__Recurring_Donation_Installment_Number__c
# FROM Opportunity
# WHERE Id IN ({", ".join([f"'{x}'" for x in inserted_ids])})
# ORDER BY CloseDate, Name
# """
#
# verify_result = sf.query_all(verify_query)
#
# verify_inserted_df = pd.DataFrame(verify_result["records"])
#
# if "attributes" in verify_inserted_df.columns:
#     verify_inserted_df = verify_inserted_df.drop(columns=["attributes"])
#
# verify_inserted_df

# %%
# =========================
# 20. POST-INSERT VERIFICATION BY AUTH__C
# =========================

auth_values_to_verify = (
    opps_to_insert["Auth__c"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

verified_records = []

for chunk in chunk_list(auth_values_to_verify, 100):
    auth_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    verify_query = f"""
    SELECT
        Id,
        Name,
        CloseDate,
        StageName,
        Amount,
        CurrencyIsoCode,
        Auth__c,
        npe03__Recurring_Donation__c,
        npsp__Recurring_Donation_Installment_Number__c,
        CreatedDate
    FROM Opportunity
    WHERE Auth__c IN ({auth_list_soql})
    """

    result = sf.query_all(verify_query)
    verified_records.extend(result["records"])

verified_inserted_by_auth_df = pd.DataFrame(verified_records)

if not verified_inserted_by_auth_df.empty and "attributes" in verified_inserted_by_auth_df.columns:
    verified_inserted_by_auth_df = verified_inserted_by_auth_df.drop(columns=["attributes"])

verified_inserted_by_auth_df.shape

# %%
# =========================
# SUMMARY OF LOADED CLOSED WON OPPORTUNITIES BY YEAR
# =========================

loaded_closed_won_summary_by_year = verified_inserted_by_auth_df.copy()

loaded_closed_won_summary_by_year["CloseDate"] = pd.to_datetime(
    loaded_closed_won_summary_by_year["CloseDate"],
    errors="coerce"
)

loaded_closed_won_summary_by_year["Amount"] = pd.to_numeric(
    loaded_closed_won_summary_by_year["Amount"],
    errors="coerce"
)

loaded_closed_won_summary_by_year = loaded_closed_won_summary_by_year[
    loaded_closed_won_summary_by_year["StageName"] == "Closed Won"
].copy()

loaded_closed_won_summary_by_year["year"] = (
    loaded_closed_won_summary_by_year["CloseDate"].dt.year
)

loaded_closed_won_summary_by_year = (
    loaded_closed_won_summary_by_year
    .groupby("year", dropna=False)
    .agg(
        opportunities=("Id", "count"),
        total_amount=("Amount", "sum"),
    )
    .reset_index()
    .sort_values("year")
)

loaded_closed_won_summary_by_year

# %%
# =========================
# ADD GRAND TOTAL
# =========================

grand_total = pd.DataFrame({
    "year": ["TOTAL"],
    "opportunities": [loaded_closed_won_summary_by_year["opportunities"].sum()],
    "total_amount": [loaded_closed_won_summary_by_year["total_amount"].sum()],
})

loaded_closed_won_summary_by_year_with_total = pd.concat(
    [loaded_closed_won_summary_by_year, grand_total],
    ignore_index=True
)

loaded_closed_won_summary_by_year_with_total

# %%
# =========================
# EXPORT CLOSED WON SUMMARY
# =========================

closed_won_summary_path = OUTPUT_DIR / "09_loaded_closed_won_summary_by_year.xlsx"

with pd.ExcelWriter(closed_won_summary_path, engine="openpyxl") as writer:
    loaded_closed_won_summary_by_year_with_total.to_excel(
        writer,
        sheet_name="closed_won_by_year",
        index=False
    )

print(f"Archivo creado: {closed_won_summary_path}")