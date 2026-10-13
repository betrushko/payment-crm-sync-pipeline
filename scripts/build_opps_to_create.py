import os
# %%
import re
import pandas as pd
from pathlib import Path

# %%
# =========================
# 1. CONFIGURATION
# =========================

# OwnerId used by the PaymentHub integration/service user.
# This should be used for the generated Opportunities instead of inheriting
# OwnerId from the latest existing Salesforce Opportunity.
PaymentHub_SERVICE_OWNER_ID = os.environ["SF_SERVICE_OWNER_ID"]

# %%
# =========================
# 2. PROJECT PATHS
# =========================

BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

comparison_path = OUTPUT_DIR / "05_paymenthub_salesforce_after_last_sf_review.xlsx"
sf_opps_path = OUTPUT_DIR / "04_salesforce_existing_opportunities.xlsx"

comparison_path, sf_opps_path

# %%
# =========================
# 3. LOAD INPUT FILES
# =========================

opportunity_candidates_raw = pd.read_excel(
    comparison_path,
    sheet_name="opportunity_candidates",
    dtype=str
)

candidate_summary_raw = pd.read_excel(
    comparison_path,
    sheet_name="candidate_summary",
    dtype=str
)

last_opp_by_dr_raw = pd.read_excel(
    sf_opps_path,
    sheet_name="last_opp_by_dr",
    dtype=str
)

opportunity_candidates_raw.shape, candidate_summary_raw.shape, last_opp_by_dr_raw.shape

# %%
# =========================
# 4. CLEAN PaymentHub CANDIDATES
# =========================

candidates = opportunity_candidates_raw.copy()

candidates["agreement_id"] = candidates["agreement_id"].astype(str).str.strip()
candidates["sf_recurring_donation_id"] = candidates["sf_recurring_donation_id"].astype(str).str.strip()
candidates["transaction_id"] = candidates["transaction_id"].astype(str).str.strip()

candidates["completed_date"] = pd.to_datetime(
    candidates["completed_date"],
    errors="coerce"
)

candidates["completed_at"] = pd.to_datetime(
    candidates["completed_at"],
    errors="coerce"
)

candidates["donation_amount"] = pd.to_numeric(
    candidates["donation_amount"],
    errors="coerce"
)

candidates["last_sf_installment_number"] = pd.to_numeric(
    candidates["last_sf_installment_number"],
    errors="coerce"
)

candidates = candidates.sort_values([
    "sf_recurring_donation_id",
    "completed_date",
    "completed_time",
    "transaction_id",
]).copy()

candidates.head()

# %%
# =========================
# 5. CLEAN LAST OPPORTUNITY TEMPLATE
# =========================

last_opp_by_dr = last_opp_by_dr_raw.copy()

last_opp_by_dr["sf_recurring_donation_id"] = (
    last_opp_by_dr["sf_recurring_donation_id"]
    .astype(str)
    .str.strip()
)

last_opp_by_dr["sf_close_date"] = pd.to_datetime(
    last_opp_by_dr["sf_close_date"],
    errors="coerce"
)

last_opp_by_dr["sf_amount"] = pd.to_numeric(
    last_opp_by_dr["sf_amount"],
    errors="coerce"
)

last_opp_by_dr["sf_installment_number"] = pd.to_numeric(
    last_opp_by_dr["sf_installment_number"],
    errors="coerce"
)

last_opp_by_dr.head()

# %%
# =========================
# 6. JOIN CANDIDATES WITH LAST OPPORTUNITY TEMPLATE
# =========================

template_cols = [
    "sf_recurring_donation_id",
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_primary_contact_id",
    "sf_stage_name",
    "sf_record_type_id",
    "sf_campaign_id",
    "sf_campana_origen",
    "sf_canal",
    "sf_fecha_captacion",
    "sf_campana_web",
    "sf_fuente",
    "sf_medio",
    "sf_medio_pago",
    "sf_numero_tarjeta",
    "sf_fecha_vencimiento_tc",
    "sf_owner_id",
    "sf_currency",
    "sf_installment_number",
]

available_template_cols = [
    col for col in template_cols
    if col in last_opp_by_dr.columns
]

last_opp_template = last_opp_by_dr[available_template_cols].copy()

opps_work = candidates.merge(
    last_opp_template,
    on="sf_recurring_donation_id",
    how="left",
    suffixes=("", "_template")
)

opps_work.shape

# %%
# Validate that all candidates have a template Opportunity

opps_work["sf_opportunity_id"].isna().sum()

# %%
# =========================
# 7. CALCULATE NEW INSTALLMENT NUMBERS BY BILLING PERIOD
# =========================

# Important:
# Installment numbers should increase by billing period/month, not by transaction.
# If two PaymentHub transactions exist in the same month for the same DR due to duplicate Agreement issues,
# they should share the same installment number.

opps_work["billing_period"] = opps_work["completed_date"].dt.to_period("M").astype(str)

billing_period_sequence = (
    opps_work[[
        "sf_recurring_donation_id",
        "billing_period",
        "completed_date",
    ]]
    .drop_duplicates(subset=["sf_recurring_donation_id", "billing_period"])
    .sort_values([
        "sf_recurring_donation_id",
        "billing_period",
    ])
    .copy()
)

billing_period_sequence["billing_period_sequence_after_last_sf"] = (
    billing_period_sequence
    .groupby("sf_recurring_donation_id")
    .cumcount()
    + 1
)

opps_work = opps_work.merge(
    billing_period_sequence[[
        "sf_recurring_donation_id",
        "billing_period",
        "billing_period_sequence_after_last_sf",
    ]],
    on=[
        "sf_recurring_donation_id",
        "billing_period",
    ],
    how="left"
)

opps_work["new_installment_number"] = (
    pd.to_numeric(opps_work["sf_installment_number"], errors="coerce")
    + opps_work["billing_period_sequence_after_last_sf"]
)

opps_work[[
    "sf_recurring_donation_id",
    "dr_name",
    "transaction_id",
    "completed_date",
    "billing_period",
    "sf_installment_number",
    "billing_period_sequence_after_last_sf",
    "new_installment_number",
]].sort_values([
    "sf_recurring_donation_id",
    "completed_date",
    "transaction_id",
]).head(50)

# %%
# =========================
# 8. GENERATE NEW OPPORTUNITY NAMES
# =========================

DATE_PATTERN = r"(?:\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2})"


def extract_donor_name_from_dr_name(dr_name: str) -> str:
    """
    Extracts donor name from DR name.

    Expected DR name example:
        Ameyali Jaramillo Méndez 23/10/2024

    Output:
        Ameyali Jaramillo Méndez
    """
    if pd.isna(dr_name):
        return ""

    value = str(dr_name).strip()

    # Remove trailing date if present.
    value = re.sub(rf"\s+{DATE_PATTERN}\s*$", "", value).strip()

    return value


def format_close_date_for_name(value) -> str:
    """
    Formats CloseDate for Opportunity Name as dd/mm/yyyy.
    """
    if pd.isna(value):
        return ""

    return pd.to_datetime(value).strftime("%d/%m/%Y")


def remove_final_installment_from_name(opportunity_name: str) -> str:
    """
    Removes the final numeric installment suffix from an Opportunity name.
    """
    if pd.isna(opportunity_name):
        return ""

    value = str(opportunity_name).strip()

    return re.sub(r"\s+\d+\s*$", "", value).strip()


def extract_campaign_name_from_last_opp(last_opportunity_name: str, donor_name: str) -> str:
    """
    Extracts campaign name from latest Salesforce Opportunity name.

    Expected Opportunity name structure:
        Donor Name + CloseDate + Campaign Name + Installment #

    Example:
        Ameyali Jaramillo Méndez 23/10/2024 Campaign Name 10

    Output:
        Campaign Name
    """
    if pd.isna(last_opportunity_name):
        return ""

    name_without_installment = remove_final_installment_from_name(last_opportunity_name)

    # Preferred path:
    # Remove donor name + date from the beginning.
    pattern = rf"^{re.escape(str(donor_name).strip())}\s+{DATE_PATTERN}\s+"
    campaign = re.sub(pattern, "", name_without_installment).strip()

    if campaign and campaign != name_without_installment:
        return campaign

    # Fallback:
    # Split on the first date found and take what comes after it.
    parts = re.split(DATE_PATTERN, name_without_installment, maxsplit=1)

    if len(parts) > 1:
        return parts[1].strip()

    # Last fallback:
    # Return the name without final installment.
    return name_without_installment.strip()


def build_new_opportunity_name(row) -> str:
    """
    Builds the new Opportunity name using:

        donor name from DR
        + new CloseDate
        + campaign name extracted from latest SF Opportunity
        + new installment number
    """
    donor_name = extract_donor_name_from_dr_name(row["dr_name"])
    close_date_text = format_close_date_for_name(row["completed_date"])
    campaign_name = extract_campaign_name_from_last_opp(
        row["sf_opportunity_name"],
        donor_name,
    )
    installment_text = str(int(row["new_installment_number"]))

    parts = [
        donor_name,
        close_date_text,
        campaign_name,
        installment_text,
    ]

    return " ".join([part for part in parts if str(part).strip()])


opps_work["review_donor_name_for_opp_name"] = opps_work["dr_name"].apply(
    extract_donor_name_from_dr_name
)

opps_work["review_campaign_name_extracted"] = opps_work.apply(
    lambda row: extract_campaign_name_from_last_opp(
        row["sf_opportunity_name"],
        row["review_donor_name_for_opp_name"],
    ),
    axis=1,
)

opps_work["new_opportunity_name"] = opps_work.apply(
    build_new_opportunity_name,
    axis=1,
)

opps_work[[
    "dr_name",
    "sf_opportunity_name",
    "completed_date",
    "billing_period",
    "review_donor_name_for_opp_name",
    "review_campaign_name_extracted",
    "new_installment_number",
    "new_opportunity_name",
]].head(50)

# %%
# =========================
# 9. BUILD OPPORTUNITIES TO CREATE REVIEW TABLE
# =========================

# Important:
# StageName must NOT be inherited from the last Salesforce Opportunity.
# It must be derived from the actual PaymentHub payment state.
#
# Current mapping:
# PaymentHub Payment_State = Completed -> Salesforce StageName = Closed Won
# PaymentHub Payment_State = Refunded  -> Salesforce StageName = Reembolso
#
# Important:
# OwnerId must NOT be inherited from the last Salesforce Opportunity.
# It must be the PaymentHub service/integration user.

opps_to_create_review = pd.DataFrame({
    # Salesforce insert fields
    "Name": opps_work["new_opportunity_name"],
    "npsp__Primary_Contact__c": opps_work.get("sf_primary_contact_id"),
    "CloseDate": opps_work["completed_date"].dt.date,
    "StageName": opps_work["payment_state"].map({
        "Completed": "Closed Won",
        "Refunded": "Reembolso",
    }),
    "RecordTypeId": opps_work.get("sf_record_type_id"),
    "Amount": opps_work["donation_amount"],
    "CurrencyIsoCode": opps_work["currency"],
    "CampaignId": opps_work.get("sf_campaign_id"),
    "Campa_a_de_origen__c": opps_work.get("sf_campana_origen"),
    "Canal__c": opps_work.get("sf_canal"),
    "Fecha_de_captacion__c": opps_work.get("sf_fecha_captacion"),
    "Campana_web__c": opps_work.get("sf_campana_web"),
    "Fuente__c": opps_work.get("sf_fuente"),
    "Medio__c": opps_work.get("sf_medio"),
    "Auth__c": opps_work["transaction_id"],
    "Medio_de_Pago__c": opps_work.get("sf_medio_pago"),
    "N_mero_de_Tarjeta__c": opps_work.get("sf_numero_tarjeta"),
    "Fecha_de_vencimiento_TC__c": opps_work.get("sf_fecha_vencimiento_tc"),
    "npe03__Recurring_Donation__c": opps_work["sf_recurring_donation_id"],
    "npsp__Recurring_Donation_Installment_Number__c": opps_work["new_installment_number"],
    "OwnerId": PaymentHub_SERVICE_OWNER_ID,

    # Review / traceability fields - not for insert
    "review_agreement_id": opps_work["agreement_id"],
    "review_dr_name": opps_work["dr_name"],
    "review_paymenthub_transaction_id": opps_work["transaction_id"],
    "review_paymenthub_payment_state": opps_work["payment_state"],
    "review_paymenthub_completed_time": opps_work["completed_time"],
    "review_previous_sf_opportunity_id": opps_work["sf_opportunity_id"],
    "review_previous_sf_opportunity_name": opps_work["sf_opportunity_name"],
    "review_previous_installment_number": opps_work["sf_installment_number"],
    "review_billing_period": opps_work["billing_period"],
    "review_billing_period_sequence_after_last_sf": opps_work["billing_period_sequence_after_last_sf"],
    "review_donor_name_for_opp_name": opps_work["review_donor_name_for_opp_name"],
    "review_campaign_name_extracted": opps_work["review_campaign_name_extracted"],
    "review_previous_sf_owner_id": opps_work.get("sf_owner_id"),
    "review_target_owner_id": PaymentHub_SERVICE_OWNER_ID,
})

opps_to_create_review.head()

# %%
# =========================
# 10. BASIC VALIDATIONS
# =========================

opps_to_create_review.shape

# %%
opps_to_create_review["Auth__c"].nunique(), opps_to_create_review.shape[0]

# %%
required_fields = [
    "Name",
    "CloseDate",
    "StageName",
    "Amount",
    "CurrencyIsoCode",
    "Auth__c",
    "npe03__Recurring_Donation__c",
    "npsp__Recurring_Donation_Installment_Number__c",
    "OwnerId",
]

missing_required_fields = opps_to_create_review[
    opps_to_create_review[required_fields].isna().any(axis=1)
].copy()

missing_required_fields

# %%
# Validate StageName mapping from PaymentHub payment state

stage_mapping_validation = (
    opps_to_create_review
    .groupby(
        ["review_paymenthub_payment_state", "StageName"],
        dropna=False
    )
    .size()
    .reset_index(name="records")
)

stage_mapping_validation

# %%
# Validate OwnerId assignment

owner_id_validation = (
    opps_to_create_review
    .groupby("OwnerId", dropna=False)
    .size()
    .reset_index(name="records")
    .sort_values("records", ascending=False)
)

owner_id_validation

# %%
unexpected_owner_ids = opps_to_create_review[
    opps_to_create_review["OwnerId"] != PaymentHub_SERVICE_OWNER_ID
].copy()

unexpected_owner_ids[[
    "review_agreement_id",
    "review_dr_name",
    "Name",
    "Auth__c",
    "OwnerId",
    "review_previous_sf_owner_id",
    "review_target_owner_id",
]]

# %%
# =========================
# 11. REVIEW SUMMARY
# =========================

review_summary = (
    opps_to_create_review
    .groupby([
        "review_agreement_id",
        "npe03__Recurring_Donation__c",
        "review_dr_name",
    ], dropna=False)
    .agg(
        opportunities_to_create=("Auth__c", "nunique"),
        billing_periods_to_create=("review_billing_period", "nunique"),
        first_close_date=("CloseDate", "min"),
        last_close_date=("CloseDate", "max"),
        total_amount=("Amount", "sum"),
        first_installment=("npsp__Recurring_Donation_Installment_Number__c", "min"),
        last_installment=("npsp__Recurring_Donation_Installment_Number__c", "max"),
        target_owner_id=("OwnerId", "first"),
    )
    .reset_index()
    .sort_values([
        "review_agreement_id",
        "first_close_date",
    ])
)

review_summary

# %%
# =========================
# 12. SPLIT INSERT FIELDS AND REVIEW FIELDS
# =========================

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

opps_to_insert_draft = opps_to_create_review[insert_columns].copy()

opps_to_insert_draft.head()

# %%
# =========================
# 13. VALIDATE INSTALLMENT SEQUENCES
# =========================

installment_validation = opps_to_create_review[[
    "review_agreement_id",
    "npe03__Recurring_Donation__c",
    "review_dr_name",
    "Name",
    "CloseDate",
    "Amount",
    "Auth__c",
    "review_previous_installment_number",
    "review_billing_period",
    "review_billing_period_sequence_after_last_sf",
    "npsp__Recurring_Donation_Installment_Number__c",
]].copy()

installment_validation.sort_values([
    "review_agreement_id",
    "CloseDate",
    "review_billing_period",
    "Auth__c",
]).head(100)

# %%
# =========================
# 14. CHECK INSTALLMENT CONTINUITY BY DR
# =========================

installment_continuity = (
    opps_to_create_review
    .groupby([
        "review_agreement_id",
        "npe03__Recurring_Donation__c",
        "review_dr_name",
    ], dropna=False)
    .agg(
        previous_installment=("review_previous_installment_number", "first"),
        first_new_installment=("npsp__Recurring_Donation_Installment_Number__c", "min"),
        last_new_installment=("npsp__Recurring_Donation_Installment_Number__c", "max"),
        opportunities_to_create=("Auth__c", "nunique"),
        billing_periods_to_create=("review_billing_period", "nunique"),
    )
    .reset_index()
)

installment_continuity["previous_installment"] = pd.to_numeric(
    installment_continuity["previous_installment"],
    errors="coerce"
)

installment_continuity["expected_first_new_installment"] = (
    installment_continuity["previous_installment"] + 1
)

installment_continuity["expected_last_new_installment"] = (
    installment_continuity["expected_first_new_installment"]
    + installment_continuity["billing_periods_to_create"]
    - 1
)

installment_continuity["first_installment_ok"] = (
    installment_continuity["first_new_installment"]
    == installment_continuity["expected_first_new_installment"]
)

installment_continuity["last_installment_ok"] = (
    installment_continuity["last_new_installment"]
    == installment_continuity["expected_last_new_installment"]
)

installment_continuity

# %%
# =========================
# 15. CHECK INSTALLMENT ISSUES
# =========================

installment_issues = installment_continuity[
    (installment_continuity["first_installment_ok"] == False)
    | (installment_continuity["last_installment_ok"] == False)
].copy()

installment_issues

# %%
# =========================
# 16. CHECK CRITICAL FIELD COMPLETENESS
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

critical_missing = opps_to_create_review[
    opps_to_create_review[critical_fields].isna().any(axis=1)
].copy()

critical_missing[critical_fields + [
    "review_agreement_id",
    "review_dr_name",
    "review_paymenthub_transaction_id",
]]

# %%
# =========================
# 17. CHECK REFUNDED CANDIDATES
# =========================

refunded_candidates = opps_to_create_review[
    opps_to_create_review["review_paymenthub_payment_state"] == "Refunded"
].copy()

refunded_candidates[[
    "review_agreement_id",
    "review_dr_name",
    "Name",
    "CloseDate",
    "Amount",
    "Auth__c",
    "StageName",
    "review_paymenthub_payment_state",
    "npsp__Recurring_Donation_Installment_Number__c",
    "OwnerId",
]]

# %%
# =========================
# 18. CHECK AMOUNTS BY PAYMENT STATE
# =========================

amount_summary_by_state = (
    opps_to_create_review
    .groupby("review_paymenthub_payment_state", dropna=False)
    .agg(
        records=("Auth__c", "nunique"),
        total_amount=("Amount", "sum"),
        min_amount=("Amount", "min"),
        max_amount=("Amount", "max"),
    )
    .reset_index()
)

amount_summary_by_state

# %%
# =========================
# 19. FLAG AGREEMENTS WITH KNOWN PaymentHub DUPLICATION ISSUE
# =========================

known_fh_duplicate_issue_agreements = [
    "FH4420592",
    "FH4424438",
    "FH4446128",
    "FH4466960",
    "FH4481318",
]

opps_to_create_review["review_known_fh_duplicate_issue"] = (
    opps_to_create_review["review_agreement_id"]
    .isin(known_fh_duplicate_issue_agreements)
)

opps_to_create_review["review_special_case_note"] = ""

opps_to_create_review.loc[
    opps_to_create_review["review_known_fh_duplicate_issue"],
    "review_special_case_note"
] = "Known PaymentHub duplicate Agreement issue. Review before insert."

opps_to_create_review.loc[
    opps_to_create_review["review_agreement_id"] == "FH4446128",
    "review_special_case_note"
] = "Known PaymentHub duplicate Agreement issue. Exception noted: appears to be charging."

opps_to_create_review.loc[
    opps_to_create_review["review_agreement_id"] == "FH4420592",
    "review_special_case_note"
] = (
    "Known PaymentHub duplicate Agreement issue. Same-month duplicated transactions "
    "should share the same installment number. Review before insert."
)

# %%
# =========================
# 20. ADD REVIEW FLAGS AND INSERT RECOMMENDATIONS
# =========================

opps_to_create_review["review_is_refunded"] = (
    opps_to_create_review["review_paymenthub_payment_state"] == "Refunded"
)

opps_to_create_review["review_insert_recommendation"] = "include"

opps_to_create_review.loc[
    opps_to_create_review["review_is_refunded"],
    "review_insert_recommendation"
] = "include_as_reembolso"

opps_to_create_review.loc[
    opps_to_create_review["review_known_fh_duplicate_issue"],
    "review_insert_recommendation"
] = "manual_review_known_fh_duplicate_issue"

opps_to_create_review.loc[
    opps_to_create_review["review_is_refunded"]
    & opps_to_create_review["review_known_fh_duplicate_issue"],
    "review_insert_recommendation"
] = "manual_review_refunded_and_duplicate_issue"

opps_to_create_review[[
    "review_agreement_id",
    "review_dr_name",
    "Name",
    "CloseDate",
    "Amount",
    "Auth__c",
    "StageName",
    "OwnerId",
    "review_paymenthub_payment_state",
    "review_is_refunded",
    "review_known_fh_duplicate_issue",
    "review_insert_recommendation",
    "review_special_case_note",
]].head(30)

# %%
# =========================
# 21. VALIDATE BUSINESS RULES
# =========================

stage_mapping_validation = (
    opps_to_create_review
    .groupby(
        ["review_paymenthub_payment_state", "StageName"],
        dropna=False
    )
    .size()
    .reset_index(name="records")
)

stage_mapping_validation

# %%
opps_to_create_review["StageName"].value_counts(dropna=False)

# %%
owner_id_validation = (
    opps_to_create_review
    .groupby("OwnerId", dropna=False)
    .size()
    .reset_index(name="records")
    .sort_values("records", ascending=False)
)

owner_id_validation

# %%
unexpected_owner_ids = opps_to_create_review[
    opps_to_create_review["OwnerId"] != PaymentHub_SERVICE_OWNER_ID
].copy()

unexpected_owner_ids.shape

# %%
# =========================
# 22. REVIEW KNOWN PaymentHub DUPLICATE ISSUE AGREEMENTS
# =========================

known_duplicate_issue_candidates = opps_to_create_review[
    opps_to_create_review["review_known_fh_duplicate_issue"]
].copy()

known_duplicate_issue_candidates[[
    "review_agreement_id",
    "review_dr_name",
    "review_billing_period",
    "Name",
    "CloseDate",
    "Amount",
    "Auth__c",
    "StageName",
    "OwnerId",
    "review_paymenthub_payment_state",
    "npsp__Recurring_Donation_Installment_Number__c",
    "review_insert_recommendation",
    "review_special_case_note",
]].sort_values([
    "review_agreement_id",
    "CloseDate",
    "Auth__c",
])

# %%
# =========================
# 23. VALIDATE SAME-MONTH INSTALLMENT LOGIC
# =========================

same_month_installment_validation = (
    opps_to_create_review
    .groupby([
        "review_agreement_id",
        "npe03__Recurring_Donation__c",
        "review_billing_period",
    ], dropna=False)
    .agg(
        records=("Auth__c", "nunique"),
        distinct_installments=("npsp__Recurring_Donation_Installment_Number__c", "nunique"),
        installment_values=("npsp__Recurring_Donation_Installment_Number__c", lambda s: sorted(set(s))),
        payment_states=("review_paymenthub_payment_state", lambda s: sorted(set(s))),
    )
    .reset_index()
)

same_month_installment_issues = same_month_installment_validation[
    same_month_installment_validation["distinct_installments"] > 1
].copy()

same_month_installment_issues

# %%
# =========================
# 24. REVIEW FH4420592 SPECIFICALLY
# =========================

fh4420592_review = opps_to_create_review[
    opps_to_create_review["review_agreement_id"] == "FH4420592"
].copy()

fh4420592_review[[
    "review_agreement_id",
    "review_dr_name",
    "review_billing_period",
    "Name",
    "CloseDate",
    "Amount",
    "Auth__c",
    "StageName",
    "OwnerId",
    "review_paymenthub_payment_state",
    "npsp__Recurring_Donation_Installment_Number__c",
    "review_special_case_note",
]].sort_values([
    "CloseDate",
    "Auth__c",
])

# %%
# =========================
# 25. CHECK DUPLICATE OPPORTUNITY NAMES
# =========================

duplicate_opportunity_names = (
    opps_to_create_review
    .groupby("Name", dropna=False)
    .size()
    .reset_index(name="records")
    .query("records > 1")
    .sort_values("records", ascending=False)
)

duplicate_opportunity_names

# %%
# =========================
# 26. FINAL INSERT DRAFT
# =========================

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

opps_to_insert_final_draft = opps_to_create_review[insert_columns].copy()

opps_to_insert_final_draft.shape

# %%
# Validate final draft OwnerId

opps_to_insert_final_draft["OwnerId"].value_counts(dropna=False)

# %%
# =========================
# 27. FINAL REVIEW EXPORT
# =========================

output_path = OUTPUT_DIR / "06_opportunities_to_create_review.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    opportunity_candidates_raw.to_excel(writer, sheet_name="source_candidates", index=False)
    candidate_summary_raw.to_excel(writer, sheet_name="source_candidate_summary", index=False)
    last_opp_by_dr_raw.to_excel(writer, sheet_name="last_opp_by_dr", index=False)
    opps_work.to_excel(writer, sheet_name="opps_work", index=False)
    opps_to_create_review.to_excel(writer, sheet_name="opps_to_create_review", index=False)
    opps_to_insert_final_draft.to_excel(writer, sheet_name="opps_to_insert_final_draft", index=False)
    review_summary.to_excel(writer, sheet_name="review_summary", index=False)
    installment_validation.to_excel(writer, sheet_name="installment_validation", index=False)
    installment_continuity.to_excel(writer, sheet_name="installment_continuity", index=False)
    installment_issues.to_excel(writer, sheet_name="installment_issues", index=False)
    critical_missing.to_excel(writer, sheet_name="critical_missing", index=False)
    refunded_candidates.to_excel(writer, sheet_name="refunded_candidates", index=False)
    amount_summary_by_state.to_excel(writer, sheet_name="amount_summary_by_state", index=False)
    known_duplicate_issue_candidates.to_excel(writer, sheet_name="known_duplicate_issue", index=False)
    same_month_installment_validation.to_excel(writer, sheet_name="same_month_installments", index=False)
    same_month_installment_issues.to_excel(writer, sheet_name="same_month_inst_issues", index=False)
    fh4420592_review.to_excel(writer, sheet_name="FH4420592_review", index=False)
    duplicate_opportunity_names.to_excel(writer, sheet_name="duplicate_opp_names", index=False)
    stage_mapping_validation.to_excel(writer, sheet_name="stage_mapping_validation", index=False)
    owner_id_validation.to_excel(writer, sheet_name="owner_id_validation", index=False)
    unexpected_owner_ids.to_excel(writer, sheet_name="unexpected_owner_ids", index=False)

print(f"Archivo final actualizado: {output_path}")