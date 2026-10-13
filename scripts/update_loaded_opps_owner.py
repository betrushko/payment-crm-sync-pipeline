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

PaymentHub_SERVICE_OWNER_ID = os.environ["SF_SERVICE_OWNER_ID"]

BATCH_SIZE = 100

# %%
# =========================
# 2. PROJECT PATHS
# =========================

BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

OUTPUT_DIR = BASE_DIR / "data" / "output"

source_path = OUTPUT_DIR / "06_opportunities_to_create_review.xlsx"

source_path

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
# 6. LOAD SOURCE INSERT DRAFT
# =========================

opps_to_insert_draft = pd.read_excel(
    source_path,
    sheet_name="opps_to_insert_final_draft",
    dtype=str
)

opps_to_insert_draft.shape

# %%
# =========================
# 7. CLEAN SOURCE AUTH VALUES
# =========================

source_auths = (
    opps_to_insert_draft["Auth__c"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

len(source_auths)

# %%
# Validate duplicate Auth__c in source file

source_auth_duplicates = (
    opps_to_insert_draft
    .groupby("Auth__c", dropna=False)
    .size()
    .reset_index(name="records")
    .query("records > 1")
)

source_auth_duplicates

# %%
# =========================
# 8. QUERY SALESFORCE OPPORTUNITIES BY AUTH__C
# =========================

sf_opp_records = []

for chunk in chunk_list(source_auths, 100):
    auth_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    opp_query = f"""
    SELECT
        Id,
        Name,
        CloseDate,
        StageName,
        Amount,
        CurrencyIsoCode,
        Auth__c,
        OwnerId,
        Owner.Name,
        npe03__Recurring_Donation__c,
        npsp__Recurring_Donation_Installment_Number__c,
        CreatedDate,
        LastModifiedDate
    FROM Opportunity
    WHERE Auth__c IN ({auth_list_soql})
    """

    result = sf.query_all(opp_query)
    sf_opp_records.extend(result["records"])

sf_loaded_opps = pd.DataFrame(sf_opp_records)

if not sf_loaded_opps.empty and "attributes" in sf_loaded_opps.columns:
    sf_loaded_opps = sf_loaded_opps.drop(columns=["attributes"])

sf_loaded_opps.shape

# %%
# =========================
# 9. FLATTEN OWNER FIELD
# =========================

sf_loaded_opps_clean = sf_loaded_opps.copy()

if "Owner" in sf_loaded_opps_clean.columns:
    sf_loaded_opps_clean["OwnerName"] = sf_loaded_opps_clean["Owner"].apply(
        lambda x: x.get("Name") if isinstance(x, dict) else pd.NA
    )
    sf_loaded_opps_clean = sf_loaded_opps_clean.drop(columns=["Owner"])

sf_loaded_opps_clean.head()

# %%
# =========================
# 10. VALIDATE SALESFORCE MATCHES
# =========================

sf_loaded_opps_clean["Auth__c"] = (
    sf_loaded_opps_clean["Auth__c"]
    .astype(str)
    .str.strip()
)

found_auths = set(sf_loaded_opps_clean["Auth__c"].dropna().tolist())
expected_auths = set(source_auths)

missing_in_salesforce_auths = sorted(expected_auths - found_auths)
unexpected_salesforce_auths = sorted(found_auths - expected_auths)

missing_in_salesforce_df = pd.DataFrame({
    "Auth__c": missing_in_salesforce_auths
})

unexpected_salesforce_auths_df = pd.DataFrame({
    "Auth__c": unexpected_salesforce_auths
})

len(missing_in_salesforce_auths), len(unexpected_salesforce_auths)

# %%
# Validate duplicate Auth__c in Salesforce result

sf_auth_duplicates = (
    sf_loaded_opps_clean
    .groupby("Auth__c", dropna=False)
    .size()
    .reset_index(name="records")
    .query("records > 1")
)

sf_auth_duplicates

# %%
# =========================
# 11. IDENTIFY OPPORTUNITIES NEEDING OWNER UPDATE
# =========================

sf_loaded_opps_clean["target_owner_id"] = PaymentHub_SERVICE_OWNER_ID

sf_loaded_opps_clean["owner_needs_update"] = (
    sf_loaded_opps_clean["OwnerId"] != PaymentHub_SERVICE_OWNER_ID
)

opps_owner_update_needed = sf_loaded_opps_clean[
    sf_loaded_opps_clean["owner_needs_update"]
].copy()

opps_owner_already_correct = sf_loaded_opps_clean[
    ~sf_loaded_opps_clean["owner_needs_update"]
].copy()

opps_owner_update_needed.shape, opps_owner_already_correct.shape

# %%
# =========================
# 12. BUILD UPDATE PAYLOAD
# =========================

records_to_update = []

for _, row in opps_owner_update_needed.iterrows():
    records_to_update.append({
        "Id": row["Id"],
        "OwnerId": PaymentHub_SERVICE_OWNER_ID,
    })

update_payload_df = pd.DataFrame(records_to_update)

update_payload_df.head()

# %%
# =========================
# 13. VALIDATION SUMMARY
# =========================

validation_errors = []

if not source_auth_duplicates.empty:
    validation_errors.append("Source file has duplicate Auth__c values.")

if missing_in_salesforce_auths:
    validation_errors.append("Some source Auth__c values were not found in Salesforce.")

if unexpected_salesforce_auths:
    validation_errors.append("Salesforce query returned unexpected Auth__c values.")

if not sf_auth_duplicates.empty:
    validation_errors.append("Salesforce has duplicate Opportunities for the same Auth__c.")

owner_update_summary = pd.DataFrame({
    "metric": [
        "source_rows",
        "source_unique_auths",
        "salesforce_records_found",
        "salesforce_unique_auths_found",
        "missing_in_salesforce_auths",
        "unexpected_salesforce_auths",
        "source_auth_duplicates",
        "salesforce_auth_duplicates",
        "target_owner_id",
        "opps_already_correct_owner",
        "opps_needing_owner_update",
        "records_to_update",
        "validation_errors",
    ],
    "value": [
        len(opps_to_insert_draft),
        len(source_auths),
        len(sf_loaded_opps_clean),
        sf_loaded_opps_clean["Auth__c"].nunique(),
        len(missing_in_salesforce_auths),
        len(unexpected_salesforce_auths),
        len(source_auth_duplicates),
        len(sf_auth_duplicates),
        PaymentHub_SERVICE_OWNER_ID,
        len(opps_owner_already_correct),
        len(opps_owner_update_needed),
        len(records_to_update),
        "; ".join(validation_errors),
    ],
})

owner_update_summary

# %%
# =========================
# 14. EXPORT PREFLIGHT REVIEW
# =========================

preflight_path = OUTPUT_DIR / "12_loaded_opps_owner_update_preflight.xlsx"

with pd.ExcelWriter(preflight_path, engine="openpyxl") as writer:
    opps_to_insert_draft.to_excel(writer, sheet_name="source_insert_draft", index=False)
    sf_loaded_opps_clean.to_excel(writer, sheet_name="sf_loaded_opps", index=False)
    opps_owner_update_needed.to_excel(writer, sheet_name="owner_update_needed", index=False)
    opps_owner_already_correct.to_excel(writer, sheet_name="owner_already_correct", index=False)
    update_payload_df.to_excel(writer, sheet_name="update_payload", index=False)
    source_auth_duplicates.to_excel(writer, sheet_name="source_auth_duplicates", index=False)
    sf_auth_duplicates.to_excel(writer, sheet_name="sf_auth_duplicates", index=False)
    missing_in_salesforce_df.to_excel(writer, sheet_name="missing_in_salesforce", index=False)
    unexpected_salesforce_auths_df.to_excel(writer, sheet_name="unexpected_auths", index=False)
    owner_update_summary.to_excel(writer, sheet_name="summary", index=False)

print(f"Preflight file created: {preflight_path}")

# %%
# =========================
# 15. UPDATE SALESFORCE OPPORTUNITY OWNERID
# =========================

update_results = []

print("DRY_RUN:", DRY_RUN)
print("Target OwnerId:", PaymentHub_SERVICE_OWNER_ID)
print("Records to update:", len(records_to_update))
print("Validation errors:", validation_errors)

if DRY_RUN:
    print("DRY_RUN is True. No Opportunities were updated.")
else:
    if validation_errors:
        raise RuntimeError(
            "Owner update blocked due to validation errors: "
            + "; ".join(validation_errors)
        )

    if not records_to_update:
        print("No records to update.")
    else:
        opp_object = getattr(sf, "Opportunity")

        for record in records_to_update:
            record_id = record["Id"]
            payload = {
                key: value
                for key, value in record.items()
                if key != "Id"
            }

            try:
                opp_object.update(record_id, payload)

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
# 16. EXPORT UPDATE RESULTS
# =========================

update_results_df = pd.DataFrame(update_results)

if DRY_RUN:
    update_results_df = pd.DataFrame({
        "dry_run": [True],
        "message": ["No Opportunities updated."],
        "records_ready_for_update": [len(records_to_update)],
        "target_owner_id": [PaymentHub_SERVICE_OWNER_ID],
        "validation_errors": ["; ".join(validation_errors)],
    })

results_path = OUTPUT_DIR / "13_loaded_opps_owner_update_results.xlsx"

with pd.ExcelWriter(results_path, engine="openpyxl") as writer:
    update_payload_df.to_excel(writer, sheet_name="attempted_updates", index=False)
    update_results_df.to_excel(writer, sheet_name="update_results", index=False)
    owner_update_summary.to_excel(writer, sheet_name="summary", index=False)
    opps_owner_update_needed.to_excel(writer, sheet_name="owner_update_needed", index=False)
    opps_owner_already_correct.to_excel(writer, sheet_name="owner_already_correct", index=False)

print(f"Update results file created: {results_path}")

# %%
# =========================
# 17. POST-UPDATE VERIFICATION
# =========================

verify_records = []

for chunk in chunk_list(source_auths, 100):
    auth_list_soql = ", ".join([f"'{soql_escape(x)}'" for x in chunk])

    verify_query = f"""
    SELECT
        Id,
        Name,
        Auth__c,
        OwnerId,
        Owner.Name,
        CloseDate,
        StageName,
        Amount,
        npe03__Recurring_Donation__c,
        LastModifiedDate
    FROM Opportunity
    WHERE Auth__c IN ({auth_list_soql})
    """

    result = sf.query_all(verify_query)
    verify_records.extend(result["records"])

verify_owner_df = pd.DataFrame(verify_records)

if not verify_owner_df.empty and "attributes" in verify_owner_df.columns:
    verify_owner_df = verify_owner_df.drop(columns=["attributes"])

if not verify_owner_df.empty and "Owner" in verify_owner_df.columns:
    verify_owner_df["OwnerName"] = verify_owner_df["Owner"].apply(
        lambda x: x.get("Name") if isinstance(x, dict) else pd.NA
    )
    verify_owner_df = verify_owner_df.drop(columns=["Owner"])

verify_owner_df.shape

# %%
# =========================
# 18. POST-UPDATE OWNER VALIDATION
# =========================

post_update_owner_validation = (
    verify_owner_df
    .groupby("OwnerId", dropna=False)
    .size()
    .reset_index(name="records")
    .sort_values("records", ascending=False)
)

post_update_owner_validation

# %%
post_update_unexpected_owner = verify_owner_df[
    verify_owner_df["OwnerId"] != PaymentHub_SERVICE_OWNER_ID
].copy()

post_update_unexpected_owner.shape

# %%
# =========================
# 19. EXPORT POST-UPDATE VERIFICATION
# =========================

post_update_verification_path = OUTPUT_DIR / "14_loaded_opps_owner_post_update_verification.xlsx"

with pd.ExcelWriter(post_update_verification_path, engine="openpyxl") as writer:
    verify_owner_df.to_excel(writer, sheet_name="verify_owner", index=False)
    post_update_owner_validation.to_excel(writer, sheet_name="owner_validation", index=False)
    post_update_unexpected_owner.to_excel(writer, sheet_name="unexpected_owner", index=False)

print(f"Post-update verification file created: {post_update_verification_path}")