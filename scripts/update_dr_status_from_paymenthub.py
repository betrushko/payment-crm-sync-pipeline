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

# Safety rule:
# The FH subscriptions CSV must be recent before updating Salesforce.
# If the file is older than this threshold, the script will block real updates.
MAX_SOURCE_FILE_AGE_HOURS = 24

FH_TO_SF_STATUS_MAP = {
    "Active": "Operational",
    "Suspended": "Suspended",
    "Paused": "Paused",
    "Closed": "Closed",
    "Canceled": "Closed",
    "Cancelled": "Closed",
}

OPERATIONAL_STATUSES = {
    "Operational",
    "Open",
}

NON_OPERATIONAL_STATUSES = {
    "Suspended",
    "Paused",
    "Closed",
}

CLOSE_REASON_VALUE = "Regla de negocios"

# %%
# =========================
# 2. PROJECT PATHS
# =========================

BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"
OVERRIDES_DIR = BASE_DIR / "data" / "overrides"

fh_subscriptions_path = INPUT_DIR / "fh_subscriptions_status.csv"
fh_status_overrides_path = OVERRIDES_DIR / "fh_subscription_status_overrides.xlsx"

fh_subscriptions_path, fh_status_overrides_path

# %%
# =========================
# 3. LOAD ENVIRONMENT
# =========================

load_dotenv(BASE_DIR / ".env", override=True)

SF_ALIAS = os.getenv("SF_ALIAS", "DEMO_SANDBOX")
SF_CLI_PATH = os.getenv("SF_CLI_PATH", "sf")

SF_ALIAS, SF_CLI_PATH

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

    return json.loads(stdout[json_start:json_end + 1])


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


def soql_escape(value: str) -> str:
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def chunk_list(values: list, chunk_size: int = 100) -> list[list]:
    return [
        values[i:i + chunk_size]
        for i in range(0, len(values), chunk_size)
    ]


def is_blank(value) -> bool:
    if pd.isna(value):
        return True

    return str(value).strip() == ""


def is_not_blank(value) -> bool:
    return not is_blank(value)


# %%
# =========================
# 5. VALIDATE PaymentHub SUBSCRIPTIONS SOURCE FILE FRESHNESS
# =========================

if not fh_subscriptions_path.exists():
    raise FileNotFoundError(
        f"PaymentHub subscriptions file not found: {fh_subscriptions_path}"
    )

source_file_modified_at = pd.Timestamp.fromtimestamp(
    fh_subscriptions_path.stat().st_mtime
)

source_file_age_hours = (
    pd.Timestamp.now()
    - source_file_modified_at
).total_seconds() / 3600

source_file_is_stale = source_file_age_hours > MAX_SOURCE_FILE_AGE_HOURS

source_file_freshness_summary = pd.DataFrame({
    "metric": [
        "source_file",
        "source_file_modified_at",
        "source_file_age_hours",
        "max_source_file_age_hours",
        "source_file_is_stale",
    ],
    "value": [
        str(fh_subscriptions_path),
        source_file_modified_at,
        round(source_file_age_hours, 2),
        MAX_SOURCE_FILE_AGE_HOURS,
        source_file_is_stale,
    ],
})

source_file_freshness_summary

# %%
# =========================
# 6. CONNECT TO SALESFORCE
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
# 7. LOAD PaymentHub SUBSCRIPTIONS
# =========================

fh_subscriptions_raw = pd.read_csv(
    fh_subscriptions_path,
    dtype=str
)

fh_subscriptions = fh_subscriptions_raw.copy()

fh_subscriptions["agreement_id"] = (
    fh_subscriptions["agreement_id"]
    .astype(str)
    .str.strip()
)

# Keep original status from BigQuery for audit.
fh_subscriptions["agreement_status_bq"] = (
    fh_subscriptions["agreement_status"]
    .astype(str)
    .str.strip()
)

fh_subscriptions["agreement_status_final"] = fh_subscriptions["agreement_status_bq"]

fh_subscriptions["date_updated_parsed"] = pd.to_datetime(
    fh_subscriptions.get("date_updated"),
    errors="coerce"
)

fh_subscriptions.head()

# %%
# =========================
# 8. LOAD AND APPLY MANUAL STATUS OVERRIDES
# =========================

if fh_status_overrides_path.exists():
    fh_status_overrides = pd.read_excel(
        fh_status_overrides_path,
        sheet_name="status_overrides",
        dtype=str
    )

    fh_status_overrides["agreement_id"] = (
        fh_status_overrides["agreement_id"]
        .astype(str)
        .str.strip()
    )

    fh_status_overrides["override_agreement_status"] = (
        fh_status_overrides["override_agreement_status"]
        .astype(str)
        .str.strip()
    )

    # Optional columns. Create them if missing.
    for col in ["override_reason", "override_source"]:
        if col not in fh_status_overrides.columns:
            fh_status_overrides[col] = pd.NA

    fh_status_overrides = fh_status_overrides[[
        "agreement_id",
        "override_agreement_status",
        "override_reason",
        "override_source",
    ]].copy()

    # Validate duplicate overrides.
    duplicate_status_overrides = (
        fh_status_overrides
        .groupby("agreement_id", dropna=False)
        .size()
        .reset_index(name="records")
        .query("records > 1")
    )

    if not duplicate_status_overrides.empty:
        raise ValueError(
            "Duplicate agreement_id values found in fh_subscription_status_overrides.xlsx. "
            "Each agreement_id must appear only once."
        )

    fh_subscriptions = fh_subscriptions.merge(
        fh_status_overrides,
        on="agreement_id",
        how="left"
    )

    fh_subscriptions["has_status_override"] = (
        fh_subscriptions["override_agreement_status"].notna()
        & (fh_subscriptions["override_agreement_status"].astype(str).str.strip() != "")
    )

    fh_subscriptions.loc[
        fh_subscriptions["has_status_override"],
        "agreement_status_final"
    ] = fh_subscriptions.loc[
        fh_subscriptions["has_status_override"],
        "override_agreement_status"
    ]

else:
    fh_status_overrides = pd.DataFrame(columns=[
        "agreement_id",
        "override_agreement_status",
        "override_reason",
        "override_source",
    ])

    duplicate_status_overrides = pd.DataFrame(columns=[
        "agreement_id",
        "records",
    ])

    fh_subscriptions["override_agreement_status"] = pd.NA
    fh_subscriptions["override_reason"] = pd.NA
    fh_subscriptions["override_source"] = pd.NA
    fh_subscriptions["has_status_override"] = False

# Use final status for Salesforce target mapping.
fh_subscriptions["agreement_status"] = fh_subscriptions["agreement_status_final"]

fh_subscriptions["target_sf_status"] = (
    fh_subscriptions["agreement_status"]
    .map(FH_TO_SF_STATUS_MAP)
)

fh_subscriptions[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
]].head(50)

# %%
# =========================
# 9. REVIEW PaymentHub SUBSCRIPTION STATUS VALUES
# =========================

fh_subscription_status_review = (
    fh_subscriptions
    .groupby([
        "agreement_status_bq",
        "override_agreement_status",
        "agreement_status_final",
        "target_sf_status",
        "has_status_override",
    ], dropna=False)
    .agg(
        agreements=("agreement_id", "nunique"),
        earliest_date_updated=("date_updated_parsed", "min"),
        latest_date_updated=("date_updated_parsed", "max"),
    )
    .reset_index()
    .sort_values([
        "agreement_status_final",
        "target_sf_status",
        "has_status_override",
    ])
)

fh_subscription_status_review

# %%
# =========================
# 10. VALIDATE PaymentHub STATUS VALUES
# =========================

unknown_fh_status = fh_subscriptions[
    fh_subscriptions["target_sf_status"].isna()
].copy()

unknown_fh_status[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
]]

# %%
# =========================
# 11. REVIEW FH4973995 STATUS SOURCE
# =========================

fh4973995_source_review = fh_subscriptions[
    fh_subscriptions["agreement_id"] == "FH4973995"
].copy()

fh4973995_source_review[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "date_updated",
    "override_reason",
    "override_source",
]]

# %%
# =========================
# 12. QUERY SALESFORCE RECURRING DONATIONS BY AGREEMENT ID
# =========================

agreement_ids = (
    fh_subscriptions["agreement_id"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

sf_dr_records = []

for chunk in chunk_list(agreement_ids, 100):
    agreement_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    dr_query = f"""
    SELECT
        Id,
        Name,
        Transaction_ID_Unique__c,
        npe03__Open_Ended_Status__c,
        Fecha_de_cierre_de_DDRR__c,
        Motivo_de_cierre_de_DDRR__c
    FROM npe03__Recurring_Donation__c
    WHERE Transaction_ID_Unique__c IN ({agreement_list_soql})
    """

    result = sf.query_all(dr_query)
    sf_dr_records.extend(result["records"])

sf_drs = pd.DataFrame(sf_dr_records)

if not sf_drs.empty and "attributes" in sf_drs.columns:
    sf_drs = sf_drs.drop(columns=["attributes"])

sf_drs.shape

# %%
# =========================
# 13. CLEAN SALESFORCE RECURRING DONATIONS
# =========================

sf_drs_clean = sf_drs.copy()

sf_drs_clean = sf_drs_clean.rename(columns={
    "Id": "sf_recurring_donation_id",
    "Name": "sf_dr_name",
    "Transaction_ID_Unique__c": "agreement_id",
    "npe03__Open_Ended_Status__c": "current_sf_status",
    "Fecha_de_cierre_de_DDRR__c": "current_sf_close_date",
    "Motivo_de_cierre_de_DDRR__c": "current_sf_close_reason",
})

sf_drs_clean["agreement_id"] = (
    sf_drs_clean["agreement_id"]
    .astype(str)
    .str.strip()
)

sf_drs_clean["current_sf_status"] = (
    sf_drs_clean["current_sf_status"]
    .astype(str)
    .str.strip()
)

sf_drs_clean["current_sf_close_date"] = pd.to_datetime(
    sf_drs_clean["current_sf_close_date"],
    errors="coerce"
)

sf_drs_clean["current_sf_close_reason"] = sf_drs_clean[
    "current_sf_close_reason"
].where(
    sf_drs_clean["current_sf_close_reason"].notna(),
    pd.NA
)

sf_drs_clean.head()

# %%
# =========================
# 14. QUERY LATEST SUCCESSFUL OPPORTUNITY BY DR
# =========================

dr_ids = (
    sf_drs_clean["sf_recurring_donation_id"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

sf_opp_records = []

for chunk in chunk_list(dr_ids, 100):
    dr_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    opp_query = f"""
    SELECT
        Id,
        Name,
        CloseDate,
        StageName,
        Amount,
        Auth__c,
        npe03__Recurring_Donation__c,
        npsp__Recurring_Donation_Installment_Number__c
    FROM Opportunity
    WHERE npe03__Recurring_Donation__c IN ({dr_list_soql})
    AND StageName = 'Closed Won'
    ORDER BY npe03__Recurring_Donation__c, CloseDate DESC
    """

    result = sf.query_all(opp_query)
    sf_opp_records.extend(result["records"])

sf_closed_won_opps = pd.DataFrame(sf_opp_records)

if not sf_closed_won_opps.empty and "attributes" in sf_closed_won_opps.columns:
    sf_closed_won_opps = sf_closed_won_opps.drop(columns=["attributes"])

sf_closed_won_opps.shape

# %%
# =========================
# 15. CALCULATE LATEST SUCCESSFUL CLOSE DATE BY DR
# =========================

sf_closed_won_opps_clean = sf_closed_won_opps.copy()

sf_closed_won_opps_clean = sf_closed_won_opps_clean.rename(columns={
    "Id": "sf_opportunity_id",
    "Name": "sf_opportunity_name",
    "CloseDate": "sf_opp_close_date",
    "StageName": "sf_opp_stage_name",
    "Amount": "sf_opp_amount",
    "Auth__c": "sf_opp_auth",
    "npe03__Recurring_Donation__c": "sf_recurring_donation_id",
    "npsp__Recurring_Donation_Installment_Number__c": "sf_installment_number",
})

sf_closed_won_opps_clean["sf_opp_close_date"] = pd.to_datetime(
    sf_closed_won_opps_clean["sf_opp_close_date"],
    errors="coerce"
)

latest_closed_won_by_dr = (
    sf_closed_won_opps_clean
    .sort_values([
        "sf_recurring_donation_id",
        "sf_opp_close_date",
        "sf_opportunity_id",
    ])
    .groupby("sf_recurring_donation_id", dropna=False)
    .tail(1)
    .copy()
)

latest_closed_won_by_dr = latest_closed_won_by_dr[[
    "sf_recurring_donation_id",
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_opp_close_date",
    "sf_opp_amount",
    "sf_opp_auth",
    "sf_installment_number",
]].copy()

latest_closed_won_by_dr.head()

# %%
# =========================
# 16. BUILD DR STATUS UPDATE REVIEW TABLE
# =========================

dr_status_work = fh_subscriptions.merge(
    sf_drs_clean,
    on="agreement_id",
    how="left"
)

dr_status_work = dr_status_work.merge(
    latest_closed_won_by_dr,
    on="sf_recurring_donation_id",
    how="left"
)

dr_status_work["calculated_close_date"] = (
    dr_status_work["sf_opp_close_date"]
    + pd.offsets.MonthBegin(1)
)

dr_status_work["status_needs_update"] = (
    dr_status_work["current_sf_status"]
    != dr_status_work["target_sf_status"]
)

dr_status_work["is_target_operational"] = (
    dr_status_work["target_sf_status"]
    .isin(OPERATIONAL_STATUSES)
)

dr_status_work["is_target_non_operational"] = (
    dr_status_work["target_sf_status"]
    .isin(NON_OPERATIONAL_STATUSES)
)

# Operational/Open rule:
# close date and close reason must be blank.
dr_status_work["close_date_needs_clear"] = (
    dr_status_work["is_target_operational"]
    & dr_status_work["current_sf_close_date"].notna()
)

dr_status_work["close_reason_needs_clear"] = (
    dr_status_work["is_target_operational"]
    & dr_status_work["current_sf_close_reason"].apply(is_not_blank)
)

# Non-operational rule:
# close date is first day of the next month after latest Closed Won Opportunity.
dr_status_work["close_date_needs_update"] = (
    dr_status_work["is_target_non_operational"]
    & dr_status_work["sf_opp_close_date"].notna()
    & (
        dr_status_work["current_sf_close_date"].isna()
        | (
            dr_status_work["current_sf_close_date"]
            < dr_status_work["sf_opp_close_date"]
        )
    )
)

# Non-operational rule:
# close reason is filled only if blank.
dr_status_work["close_reason_needs_update"] = (
    dr_status_work["is_target_non_operational"]
    & dr_status_work["current_sf_close_reason"].apply(is_blank)
)

dr_status_work["missing_sf_dr"] = (
    dr_status_work["sf_recurring_donation_id"].isna()
)

dr_status_work["missing_latest_closed_won_opp"] = (
    dr_status_work["is_target_non_operational"]
    & dr_status_work["sf_opp_close_date"].isna()
)

dr_status_work["blocked_reason"] = ""

dr_status_work.loc[
    dr_status_work["target_sf_status"].isna(),
    "blocked_reason"
] = "Unknown PaymentHub agreement_status mapping."

dr_status_work.loc[
    dr_status_work["missing_sf_dr"],
    "blocked_reason"
] = "No Salesforce Recurring Donation found for agreement_id."

dr_status_work.loc[
    dr_status_work["missing_latest_closed_won_opp"],
    "blocked_reason"
] = "No Closed Won Opportunity found to calculate close date."

dr_status_work["has_blocking_issue"] = (
    dr_status_work["blocked_reason"].astype(str).str.strip() != ""
)

dr_status_work["needs_any_update"] = (
    dr_status_work["status_needs_update"]
    | dr_status_work["close_date_needs_update"]
    | dr_status_work["close_reason_needs_update"]
    | dr_status_work["close_date_needs_clear"]
    | dr_status_work["close_reason_needs_clear"]
)

dr_status_work["ready_for_update"] = (
    dr_status_work["needs_any_update"]
    & ~dr_status_work["has_blocking_issue"]
)

dr_status_work[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
    "date_updated",
    "sf_recurring_donation_id",
    "sf_dr_name",
    "current_sf_status",
    "current_sf_close_date",
    "current_sf_close_reason",
    "sf_opp_close_date",
    "calculated_close_date",
    "status_needs_update",
    "close_date_needs_update",
    "close_reason_needs_update",
    "close_date_needs_clear",
    "close_reason_needs_clear",
    "ready_for_update",
    "blocked_reason",
]].head(50)

# %%
# =========================
# 17. SPECIFIC REVIEW FOR FH4973995
# =========================

fh4973995_review = dr_status_work[
    dr_status_work["agreement_id"] == "FH4973995"
].copy()

fh4973995_review[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
    "date_updated",
    "sf_recurring_donation_id",
    "sf_dr_name",
    "current_sf_status",
    "current_sf_close_date",
    "current_sf_close_reason",
    "sf_opp_close_date",
    "calculated_close_date",
    "status_needs_update",
    "close_date_needs_update",
    "close_reason_needs_update",
    "close_date_needs_clear",
    "close_reason_needs_clear",
    "ready_for_update",
    "blocked_reason",
]]

# %%
# =========================
# 18. BUILD SALESFORCE UPDATE PAYLOAD
# =========================

records_to_update = []

for _, row in dr_status_work[dr_status_work["ready_for_update"]].iterrows():
    record = {
        "Id": row["sf_recurring_donation_id"],
    }

    if row["status_needs_update"]:
        record["npe03__Open_Ended_Status__c"] = row["target_sf_status"]

    # Operational/Open rule:
    # Close date and close reason must be cleared.
    if row["close_date_needs_clear"]:
        record["Fecha_de_cierre_de_DDRR__c"] = None

    if row["close_reason_needs_clear"]:
        record["Motivo_de_cierre_de_DDRR__c"] = None

    # Non-operational rule:
    # Close date and close reason should be populated only when needed.
    if row["close_date_needs_update"]:
        record["Fecha_de_cierre_de_DDRR__c"] = (
            row["calculated_close_date"]
            .date()
            .isoformat()
        )

    if row["close_reason_needs_update"]:
        record["Motivo_de_cierre_de_DDRR__c"] = CLOSE_REASON_VALUE

    records_to_update.append(record)

len(records_to_update)

# %%
# =========================
# 19. REVIEW UPDATE PAYLOAD
# =========================

update_payload_df = pd.DataFrame(records_to_update)

update_payload_df.head(50)

# %%
# =========================
# 20. REVIEW OPERATIONAL CLEARING CASES
# =========================

operational_clear_review = dr_status_work[
    dr_status_work["is_target_operational"]
    & (
        dr_status_work["status_needs_update"]
        | dr_status_work["close_date_needs_clear"]
        | dr_status_work["close_reason_needs_clear"]
    )
].copy()

operational_clear_review[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
    "date_updated",
    "sf_recurring_donation_id",
    "sf_dr_name",
    "current_sf_status",
    "current_sf_close_date",
    "current_sf_close_reason",
    "status_needs_update",
    "close_date_needs_clear",
    "close_reason_needs_clear",
    "ready_for_update",
]]

# %%
# =========================
# 21. REVIEW NON-OPERATIONAL CLOSE UPDATE CASES
# =========================

non_operational_close_review = dr_status_work[
    dr_status_work["is_target_non_operational"]
    & (
        dr_status_work["status_needs_update"]
        | dr_status_work["close_date_needs_update"]
        | dr_status_work["close_reason_needs_update"]
    )
].copy()

non_operational_close_review[[
    "agreement_id",
    "agreement_status_bq",
    "override_agreement_status",
    "agreement_status_final",
    "target_sf_status",
    "has_status_override",
    "override_reason",
    "override_source",
    "date_updated",
    "sf_recurring_donation_id",
    "sf_dr_name",
    "current_sf_status",
    "current_sf_close_date",
    "current_sf_close_reason",
    "sf_opp_close_date",
    "calculated_close_date",
    "status_needs_update",
    "close_date_needs_update",
    "close_reason_needs_update",
    "ready_for_update",
]]

# %%
# =========================
# 22. SUMMARY
# =========================

status_update_summary = pd.DataFrame({
    "metric": [
        "fh_subscriptions",
        "sf_drs_found",
        "source_file_modified_at",
        "source_file_age_hours",
        "max_source_file_age_hours",
        "source_file_is_stale",
        "status_overrides_loaded",
        "status_overrides_applied",
        "unknown_fh_status",
        "missing_sf_dr",
        "missing_latest_closed_won_opp",
        "status_needs_update",
        "close_date_needs_update",
        "close_reason_needs_update",
        "close_date_needs_clear",
        "close_reason_needs_clear",
        "operational_clear_review",
        "non_operational_close_review",
        "ready_for_update",
        "records_to_update",
    ],
    "value": [
        len(fh_subscriptions),
        dr_status_work["sf_recurring_donation_id"].notna().sum(),
        source_file_modified_at,
        round(source_file_age_hours, 2),
        MAX_SOURCE_FILE_AGE_HOURS,
        source_file_is_stale,
        len(fh_status_overrides),
        dr_status_work["has_status_override"].sum(),
        dr_status_work["target_sf_status"].isna().sum(),
        dr_status_work["missing_sf_dr"].sum(),
        dr_status_work["missing_latest_closed_won_opp"].sum(),
        dr_status_work["status_needs_update"].sum(),
        dr_status_work["close_date_needs_update"].sum(),
        dr_status_work["close_reason_needs_update"].sum(),
        dr_status_work["close_date_needs_clear"].sum(),
        dr_status_work["close_reason_needs_clear"].sum(),
        len(operational_clear_review),
        len(non_operational_close_review),
        dr_status_work["ready_for_update"].sum(),
        len(records_to_update),
    ],
})

status_update_summary

# %%
# =========================
# 23. EXPORT PREFLIGHT REVIEW
# =========================

preflight_path = OUTPUT_DIR / "10_dr_status_update_preflight.xlsx"

with pd.ExcelWriter(preflight_path, engine="openpyxl") as writer:
    source_file_freshness_summary.to_excel(writer, sheet_name="source_file_freshness", index=False)
    fh_status_overrides.to_excel(writer, sheet_name="status_overrides", index=False)
    duplicate_status_overrides.to_excel(writer, sheet_name="duplicate_overrides", index=False)
    fh_subscriptions_raw.to_excel(writer, sheet_name="source_fh_subscriptions", index=False)
    fh_subscription_status_review.to_excel(writer, sheet_name="fh_status_review", index=False)
    fh_subscriptions.to_excel(writer, sheet_name="fh_subscriptions_final", index=False)
    sf_drs_clean.to_excel(writer, sheet_name="sf_drs", index=False)
    latest_closed_won_by_dr.to_excel(writer, sheet_name="latest_closed_won_by_dr", index=False)
    dr_status_work.to_excel(writer, sheet_name="dr_status_work", index=False)
    update_payload_df.to_excel(writer, sheet_name="update_payload", index=False)
    operational_clear_review.to_excel(writer, sheet_name="operational_clear", index=False)
    non_operational_close_review.to_excel(writer, sheet_name="non_operational_close", index=False)
    fh4973995_source_review.to_excel(writer, sheet_name="FH4973995_source", index=False)
    fh4973995_review.to_excel(writer, sheet_name="FH4973995_review", index=False)
    unknown_fh_status.to_excel(writer, sheet_name="unknown_fh_status", index=False)
    status_update_summary.to_excel(writer, sheet_name="summary", index=False)

print(f"Preflight file created: {preflight_path}")

# %%
# =========================
# 24. UPDATE SALESFORCE RECURRING DONATIONS
# =========================

update_results = []

print("DRY_RUN:", DRY_RUN)
print("Source file modified at:", source_file_modified_at)
print("Source file age hours:", round(source_file_age_hours, 2))
print("Source file is stale:", source_file_is_stale)
print("Status overrides loaded:", len(fh_status_overrides))
print("Status overrides applied:", dr_status_work["has_status_override"].sum())
print("Records ready for update:", len(records_to_update))

if DRY_RUN:
    print("DRY_RUN is True. No Recurring Donations were updated.")
else:
    if source_file_is_stale:
        raise RuntimeError(
            "Update blocked: fh_subscriptions_status.csv is stale. "
            "Refresh PaymentHub subscriptions first by rerunning query_paymenthub_gcp.py."
        )

    if unknown_fh_status.shape[0] > 0:
        raise RuntimeError("Update blocked: unknown FH agreement_status values found.")

    if not records_to_update:
        print("No records to update.")
    else:
        rd_object = getattr(sf, "npe03__Recurring_Donation__c")

        for record in records_to_update:
            record_id = record["Id"]
            payload = {
                key: value
                for key, value in record.items()
                if key != "Id"
            }

            try:
                rd_object.update(record_id, payload)

                update_results.append({
                    "success": True,
                    "id": record_id,
                    "errors": [],
                    "payload": json.dumps(payload, ensure_ascii=False),
                })

            except Exception as exc:
                update_results.append({
                    "success": False,
                    "id": record_id,
                    "errors": [str(exc)],
                    "payload": json.dumps(payload, ensure_ascii=False),
                })

        print(f"Update attempt count: {len(update_results)}")

# %%
# =========================
# 25. EXPORT UPDATE RESULTS
# =========================

update_results_df = pd.DataFrame(update_results)

if DRY_RUN:
    update_results_df = pd.DataFrame({
        "dry_run": [True],
        "message": ["No Recurring Donations updated."],
        "records_ready_for_update": [len(records_to_update)],
        "source_file_is_stale": [source_file_is_stale],
        "source_file_age_hours": [round(source_file_age_hours, 2)],
        "status_overrides_loaded": [len(fh_status_overrides)],
        "status_overrides_applied": [dr_status_work["has_status_override"].sum()],
    })

results_path = OUTPUT_DIR / "11_dr_status_update_results.xlsx"

with pd.ExcelWriter(results_path, engine="openpyxl") as writer:
    source_file_freshness_summary.to_excel(writer, sheet_name="source_file_freshness", index=False)
    fh_status_overrides.to_excel(writer, sheet_name="status_overrides", index=False)
    update_payload_df.to_excel(writer, sheet_name="attempted_updates", index=False)
    update_results_df.to_excel(writer, sheet_name="update_results", index=False)
    dr_status_work.to_excel(writer, sheet_name="dr_status_work", index=False)
    operational_clear_review.to_excel(writer, sheet_name="operational_clear", index=False)
    non_operational_close_review.to_excel(writer, sheet_name="non_operational_close", index=False)
    fh4973995_review.to_excel(writer, sheet_name="FH4973995_review", index=False)
    status_update_summary.to_excel(writer, sheet_name="summary", index=False)

print(f"Update results file created: {results_path}")

# %%
# =========================
# 26. SUMMARIZE UPDATE RESULTS
# =========================

if not DRY_RUN and not update_results_df.empty:
    update_results_df["success"].value_counts(dropna=False)
else:
    update_results_df

# %%
# =========================
# 27. POST-UPDATE VERIFICATION
# =========================

updated_dr_ids = update_results_df["id"].dropna().astype(str).tolist()

updated_dr_ids_soql = ", ".join([f"'{soql_escape(x)}'" for x in updated_dr_ids])

verify_dr_query = f"""
SELECT
    Id,
    Name,
    Transaction_ID_Unique__c,
    npe03__Open_Ended_Status__c,
    Fecha_de_cierre_de_DDRR__c,
    Motivo_de_cierre_de_DDRR__c
FROM npe03__Recurring_Donation__c
WHERE Id IN ({updated_dr_ids_soql})
ORDER BY Transaction_ID_Unique__c
"""

verify_dr_result = sf.query_all(verify_dr_query)

verify_updated_drs_df = pd.DataFrame(verify_dr_result["records"])

if not verify_updated_drs_df.empty and "attributes" in verify_updated_drs_df.columns:
    verify_updated_drs_df = verify_updated_drs_df.drop(columns=["attributes"])

verify_updated_drs_df