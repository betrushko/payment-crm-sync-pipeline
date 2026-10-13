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
# Base del proyecto
# Funciona si estás parado en fh_sf_sync.
# Si por accidente estás parado en scripts/, sube un nivel.
BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_DIR, INPUT_DIR, OUTPUT_DIR

# %%
# Cargar variables de entorno
load_dotenv(BASE_DIR / ".env", override=True)

SF_ALIAS = os.getenv("SF_ALIAS", "DEMO_SANDBOX")
SF_INSTANCE_URL = os.getenv("SF_INSTANCE_URL")
SF_CLI_PATH = os.getenv("SF_CLI_PATH", "sf")

SF_ALIAS, SF_INSTANCE_URL, SF_CLI_PATH

# %%
# Regex para remover códigos ANSI/color del output de Salesforce CLI.
# Esto evita errores de json.loads cuando la CLI devuelve JSON coloreado.
ANSI_ESCAPE_RE = re.compile(
    r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])"
)


def strip_ansi_codes(text: str) -> str:
    """
    Removes terminal color/control codes from CLI output.
    Salesforce CLI may return colorized JSON in some terminals.
    """
    return ANSI_ESCAPE_RE.sub("", text)


def parse_cli_output_to_dict(stdout: str) -> dict:
    """
    Parses Salesforce CLI output into a Python dict.
    Removes ANSI color codes before parsing JSON.
    """
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
    """
    Creates a simple-salesforce connection using an authenticated Salesforce CLI org.

    Required previous login in terminal:

        sf org login web --alias DEMO_SANDBOX --instance-url https://example.my.salesforce.com

    Important:
    - This uses Salesforce CLI to retrieve the access token.
    - Do not print or share the full verbose output because it contains sensitive data.
    """

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
# Crear conexión a Salesforce
sf = get_salesforce_connection_from_cli(SF_ALIAS, SF_CLI_PATH)

sf

# %%
# Query de prueba
test_result = sf.query("""
SELECT Id, Name
FROM Organization
LIMIT 1
""")

test_result["totalSize"]

# %%
# Leer Excel y preparar lista de Recurring Donations / DRs
excel_path = INPUT_DIR / "FH-SF Sync failures.xlsx"

yellow = pd.read_excel(excel_path, sheet_name="Yellow")

yellow_clean = yellow.rename(columns={
    "Id": "sf_recurring_donation_id",
    "Name": "dr_name",
    "Transaction_ID_Unique__c": "agreement_id",
})

yellow_clean = yellow_clean[
    yellow_clean["sf_recurring_donation_id"].notna()
].copy()

yellow_clean["sf_recurring_donation_id"] = (
    yellow_clean["sf_recurring_donation_id"]
    .astype(str)
    .str.strip()
)

yellow_clean["agreement_id"] = (
    yellow_clean["agreement_id"]
    .astype(str)
    .str.strip()
)

dr_ids = (
    yellow_clean["sf_recurring_donation_id"]
    .dropna()
    .unique()
    .tolist()
)

len(dr_ids), dr_ids[:5]

# %%
# Crear lista SOQL
dr_ids_soql = ", ".join([f"'{x}'" for x in dr_ids])

print(dr_ids_soql)

# %%
# Query de Opportunities existentes en Salesforce
opps_query = f"""
SELECT
    Id,
    Name,
    npsp__Primary_Contact__c,
    CloseDate,
    StageName,
    RecordTypeId,
    Amount,
    CurrencyIsoCode,
    CampaignId,
    Campa_a_de_origen__c,
    Canal__c,
    Fecha_de_captacion__c,
    Campana_web__c,
    Fuente__c,
    Medio__c,
    Auth__c,
    Medio_de_Pago__c,
    N_mero_de_Tarjeta__c,
    Fecha_de_vencimiento_TC__c,
    npe03__Recurring_Donation__c,
    npsp__Recurring_Donation_Installment_Number__c,
    OwnerId,
    CreatedDate
FROM Opportunity
WHERE npe03__Recurring_Donation__c IN ({dr_ids_soql})
ORDER BY npe03__Recurring_Donation__c, CloseDate, CreatedDate
"""

sf_opps_result = sf.query_all(opps_query)

len(sf_opps_result["records"])

# %%
# Convertir resultado a DataFrame
sf_opps = pd.DataFrame(sf_opps_result["records"])

if "attributes" in sf_opps.columns:
    sf_opps = sf_opps.drop(columns=["attributes"])

sf_opps.head()

# %%
# Normalizar columnas
sf_opps = sf_opps.rename(columns={
    "Id": "sf_opportunity_id",
    "Name": "sf_opportunity_name",
    "npsp__Primary_Contact__c": "sf_primary_contact_id",
    "CloseDate": "sf_close_date",
    "StageName": "sf_stage_name",
    "RecordTypeId": "sf_record_type_id",
    "Amount": "sf_amount",
    "CurrencyIsoCode": "sf_currency",
    "CampaignId": "sf_campaign_id",
    "Campa_a_de_origen__c": "sf_campana_origen",
    "Canal__c": "sf_canal",
    "Fecha_de_captacion__c": "sf_fecha_captacion",
    "Campana_web__c": "sf_campana_web",
    "Fuente__c": "sf_fuente",
    "Medio__c": "sf_medio",
    "Auth__c": "sf_auth",
    "Medio_de_Pago__c": "sf_medio_pago",
    "N_mero_de_Tarjeta__c": "sf_numero_tarjeta",
    "Fecha_de_vencimiento_TC__c": "sf_fecha_vencimiento_tc",
    "npe03__Recurring_Donation__c": "sf_recurring_donation_id",
    "npsp__Recurring_Donation_Installment_Number__c": "sf_installment_number",
    "OwnerId": "sf_owner_id",
    "CreatedDate": "sf_created_date",
})

sf_opps["sf_close_date"] = pd.to_datetime(sf_opps["sf_close_date"], errors="coerce")
sf_opps["sf_created_date"] = pd.to_datetime(sf_opps["sf_created_date"], errors="coerce")
sf_opps["sf_amount"] = pd.to_numeric(sf_opps["sf_amount"], errors="coerce")
sf_opps["sf_installment_number"] = pd.to_numeric(
    sf_opps["sf_installment_number"],
    errors="coerce"
)

sf_opps.head()

# %%
# Validaciones rápidas
sf_opps.shape

# %%
sf_opps["sf_recurring_donation_id"].nunique()

# %%
sf_opps["sf_stage_name"].value_counts(dropna=False)

# %%
sf_opps["sf_auth"].notna().sum(), sf_opps["sf_auth"].isna().sum()

# %%
sf_opps[[
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_recurring_donation_id",
    "sf_close_date",
    "sf_stage_name",
    "sf_amount",
    "sf_currency",
    "sf_auth",
    "sf_installment_number",
]].head(20)

# %%
# Última Opportunity por DR
# Esta será la base/plantilla para crear nuevas Opportunities faltantes.
sf_opps_sorted = sf_opps.sort_values(
    by=[
        "sf_recurring_donation_id",
        "sf_close_date",
        "sf_installment_number",
        "sf_created_date",
    ],
    ascending=[True, True, True, True]
)

last_opp_by_dr = (
    sf_opps_sorted
    .groupby("sf_recurring_donation_id", as_index=False)
    .tail(1)
    .copy()
)

last_opp_by_dr[[
    "sf_recurring_donation_id",
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_close_date",
    "sf_amount",
    "sf_currency",
    "sf_auth",
    "sf_installment_number",
]]

# %%
last_opp_by_dr.shape

# %%
def remove_timezones_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Excel does not support timezone-aware datetimes.
    This returns a copy where timezone-aware datetime columns are made timezone-naive.
    """
    df_clean = df.copy()

    for col in df_clean.columns:
        if isinstance(df_clean[col].dtype, pd.DatetimeTZDtype):
            df_clean[col] = df_clean[col].dt.tz_localize(None)

    return df_clean


sf_opps_excel = remove_timezones_for_excel(sf_opps)
last_opp_by_dr_excel = remove_timezones_for_excel(last_opp_by_dr)
yellow_clean_excel = remove_timezones_for_excel(yellow_clean)

# %%
output_path = OUTPUT_DIR / "04_salesforce_existing_opportunities.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    yellow_clean_excel.to_excel(writer, sheet_name="yellow_clean", index=False)
    sf_opps_excel.to_excel(writer, sheet_name="sf_opportunities", index=False)
    last_opp_by_dr_excel.to_excel(writer, sheet_name="last_opp_by_dr", index=False)

print(f"Archivo creado: {output_path}")