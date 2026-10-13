# %%
import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from google.cloud import bigquery
from google.oauth2 import service_account

# %%
# Base del proyecto
BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_DIR, INPUT_DIR, OUTPUT_DIR

# %%
# Cargar .env
load_dotenv(BASE_DIR / ".env", override=True)

GCP_PROJECT_ID = os.getenv("GCP_PROJECT_ID", "example-project")
GCP_KEY_FILENAME = os.getenv("GCP_KEY_FILENAME")
GCP_KEY_RELATIVE_PATH = os.getenv("GCP_KEY_RELATIVE_PATH")

GCP_PROJECT_ID, GCP_KEY_FILENAME, GCP_KEY_RELATIVE_PATH

# %%
def resolve_service_account_key_path() -> Path:
    configured = os.environ["GOOGLE_APPLICATION_CREDENTIALS"]
    path = Path(configured).expanduser()
    if not path.is_file():
        raise FileNotFoundError("Configured credential file does not exist")
    return path


key_path = resolve_service_account_key_path()

print(f"Using service account key: {key_path}")

# %%
credentials = service_account.Credentials.from_service_account_file(key_path)

client = bigquery.Client(
    credentials=credentials,
    project=GCP_PROJECT_ID,
)

client.project

# %%
# Leer Yellow desde Excel
excel_path = INPUT_DIR / "FH-SF Sync failures.xlsx"

yellow = pd.read_excel(excel_path, sheet_name="Yellow")

yellow_clean = yellow.rename(columns={
    "Id": "sf_recurring_donation_id",
    "Name": "dr_name",
    "Transaction_ID_Unique__c": "agreement_id",
})

yellow_clean = yellow_clean[
    yellow_clean["agreement_id"].notna()
].copy()

yellow_clean["agreement_id"] = (
    yellow_clean["agreement_id"]
    .astype(str)
    .str.strip()
)

agreement_ids = (
    yellow_clean["agreement_id"]
    .dropna()
    .unique()
    .tolist()
)

len(agreement_ids), agreement_ids[:10]

# %%
# Query 1: orders / cobros de PaymentHub
orders_query = """
SELECT
  Agreement_ID,
  Transaction_ID,
  Order_Status,
  Payment_State,
  completed_date,
  completed_time,
  donation_amount,
  Currency
FROM `example-project.paymenthub_mart.orders`
WHERE Agreement_ID IN UNNEST(@agreement_ids)
AND Order_Status = 'Completed'
AND Payment_State IN ('Completed', 'Refunded')
ORDER BY Agreement_ID, completed_date, completed_time
"""

orders_job_config = bigquery.QueryJobConfig(
    query_parameters=[
        bigquery.ArrayQueryParameter(
            "agreement_ids",
            "STRING",
            agreement_ids,
        )
    ]
)

fh_orders = client.query(
    orders_query,
    job_config=orders_job_config,
).to_dataframe()

fh_orders.shape

# %%
fh_orders.head()

# %%
fh_orders[["Order_Status", "Payment_State"]].value_counts(dropna=False)

# %%
# Query 2: subscriptions / estado actual de Agreements
subscriptions_query = """
SELECT
  agreement_id,
  agreement_status,
  amount,
  currency,
  frequency,
  monthly_debit_date,
  next_debit_date,
  date_created,
  date_updated,
  payment_method,
  payment_id,
  payment_gateway,
  merchant_account,
  payment_gateway_message,
  card_last_four_digits,
  card_type,
  card_expiry_date,
  customer_remote_id
FROM `example-project.paymenthub_mart.subscriptions`
WHERE agreement_id IN UNNEST(@agreement_ids)
ORDER BY agreement_id
"""

subscriptions_job_config = bigquery.QueryJobConfig(
    query_parameters=[
        bigquery.ArrayQueryParameter(
            "agreement_ids",
            "STRING",
            agreement_ids,
        )
    ]
)

fh_subscriptions = client.query(
    subscriptions_query,
    job_config=subscriptions_job_config,
).to_dataframe()

fh_subscriptions.shape

# %%
fh_subscriptions.head()

# %%
fh_subscriptions["agreement_status"].value_counts(dropna=False)

# %%
# Exportar CSVs con los mismos nombres que esperan los scripts siguientes
orders_output_path = INPUT_DIR / "fh_orders_completed_and_refunded.csv"
subscriptions_output_path = INPUT_DIR / "fh_subscriptions_status.csv"

fh_orders.to_csv(orders_output_path, index=False, encoding="utf-8-sig")
fh_subscriptions.to_csv(subscriptions_output_path, index=False, encoding="utf-8-sig")

print(f"Orders exported to: {orders_output_path}")
print(f"Subscriptions exported to: {subscriptions_output_path}")

# %%
# Export opcional de control para revisar los Agreements usados
yellow_output_path = OUTPUT_DIR / "02_yellow_agreements_used_for_paymenthub_query.xlsx"

yellow_clean.to_excel(yellow_output_path, index=False)

print(f"Yellow Agreements exported to: {yellow_output_path}")