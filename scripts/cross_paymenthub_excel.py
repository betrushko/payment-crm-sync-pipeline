# %%
import pandas as pd
from pathlib import Path

# %%
BASE_DIR = Path.cwd()

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

excel_path = INPUT_DIR / "FH-SF Sync failures.xlsx"
orders_path = INPUT_DIR / "fh_orders_completed_and_refunded.csv"
subscriptions_path = INPUT_DIR / "fh_subscriptions_status.csv"

excel_path, orders_path, subscriptions_path

# %%
print("Excel exists:", excel_path.exists())
print("Orders CSV exists:", orders_path.exists())
print("Subscriptions CSV exists:", subscriptions_path.exists())

# %%
yellow = pd.read_excel(excel_path, sheet_name="Yellow")
fh_orders_raw = pd.read_csv(orders_path, dtype=str)
fh_subscriptions_raw = pd.read_csv(subscriptions_path, dtype=str)

yellow.shape, fh_orders_raw.shape, fh_subscriptions_raw.shape

# %%
yellow_clean = yellow.rename(columns={
    "Id": "sf_recurring_donation_id",
    "Name": "dr_name",
    "Transaction_ID_Unique__c": "agreement_id",
    "npe03__Open_Ended_Status__c": "sf_dr_status",
    "Fecha_de_cierre_de_DDRR__c": "sf_dr_close_date",
    "Motivo_de_cierre_de_DDRR__c": "sf_dr_close_reason",
    "npe03__Last_Payment_Date__c": "sf_last_payment_date",
    "npe03__Recurring_Donation_Campaign__r.Name": "sf_campaign_name",
})

yellow_clean = yellow_clean[[
    "sf_recurring_donation_id",
    "dr_name",
    "agreement_id",
    "sf_campaign_name",
    "sf_dr_status",
    "sf_dr_close_date",
    "sf_dr_close_reason",
    "sf_last_payment_date",
]].copy()

yellow_clean["agreement_id"] = yellow_clean["agreement_id"].astype(str).str.strip()

yellow_clean.head()

# %%
yellow_clean.info()

# %%
yellow_clean["agreement_id"].nunique(), yellow_clean.shape

# %%
fh_orders = fh_orders_raw.rename(columns={
    "Agreement_ID": "agreement_id",
    "Transaction_ID": "transaction_id",
    "Order_Status": "order_status",
    "Payment_State": "payment_state",
    "Currency": "currency",
})

fh_orders["agreement_id"] = fh_orders["agreement_id"].astype(str).str.strip()
fh_orders["transaction_id"] = fh_orders["transaction_id"].astype(str).str.strip()

fh_orders["completed_date"] = pd.to_datetime(
    fh_orders["completed_date"],
    errors="coerce"
)

fh_orders["completed_at"] = pd.to_datetime(
    fh_orders["completed_date"].dt.strftime("%Y-%m-%d")
    + " "
    + fh_orders["completed_time"].fillna("00:00:00"),
    errors="coerce"
)

fh_orders["donation_amount"] = pd.to_numeric(
    fh_orders["donation_amount"],
    errors="coerce"
)

fh_orders.head()

# %%
fh_orders.info()

# %%
fh_orders["agreement_id"].nunique(), fh_orders.shape

# %%
fh_orders[["order_status", "payment_state"]].value_counts(dropna=False)

# %%
fh_subscriptions = fh_subscriptions_raw.copy()

fh_subscriptions["agreement_id"] = (
    fh_subscriptions["agreement_id"]
    .astype(str)
    .str.strip()
)

subscription_cols = [
    "agreement_id",
    "agreement_status",
    "amount",
    "currency",
    "frequency",
    "monthly_debit_date",
    "next_debit_date",
    "date_created",
    "date_updated",
    "payment_method",
    "payment_id",
    "payment_gateway",
    "merchant_account",
    "payment_gateway_message",
    "card_last_four_digits",
    "card_type",
    "card_expiry_date",
    "customer_remote_id",
]

fh_subscriptions = fh_subscriptions[
    [col for col in subscription_cols if col in fh_subscriptions.columns]
].copy()

fh_subscriptions["amount"] = pd.to_numeric(
    fh_subscriptions["amount"],
    errors="coerce"
)

for col in ["monthly_debit_date", "next_debit_date", "date_created", "date_updated", "card_expiry_date"]:
    if col in fh_subscriptions.columns:
        fh_subscriptions[col] = pd.to_datetime(
            fh_subscriptions[col],
            errors="coerce"
        )

fh_subscriptions.head()

# %%
fh_subscriptions.info()

# %%
fh_subscriptions["agreement_id"].nunique(), fh_subscriptions.shape

# %%
fh_subscriptions["agreement_status"].value_counts(dropna=False)

# %%
orders_joined = fh_orders.merge(
    yellow_clean,
    on="agreement_id",
    how="left",
    indicator=True
)

orders_joined.head()

# %%
orders_joined["_merge"].value_counts(dropna=False)

# %%
orders_not_in_yellow = orders_joined[
    orders_joined["_merge"] == "left_only"
].copy()

orders_not_in_yellow.shape

# %%
payment_counts = (
    fh_orders
    .groupby(["order_status", "payment_state"], dropna=False)
    .size()
    .reset_index(name="records")
)

payment_counts

# %%
tx_summary = (
    fh_orders
    .groupby("agreement_id")
    .agg(
        transactions_count=("transaction_id", "nunique"),
        first_charge_date=("completed_date", "min"),
        last_charge_date=("completed_date", "max"),
        total_charged_amount=("donation_amount", "sum"),
        completed_count=("payment_state", lambda s: (s == "Completed").sum()),
        refunded_count=("payment_state", lambda s: (s == "Refunded").sum()),
    )
    .reset_index()
)

tx_summary.head()

# %%
agreement_summary = yellow_clean.merge(
    tx_summary,
    on="agreement_id",
    how="left"
)

agreement_summary = agreement_summary.merge(
    fh_subscriptions,
    on="agreement_id",
    how="left",
    suffixes=("", "_fh_subscription")
)

agreement_summary["has_fh_orders"] = agreement_summary["transactions_count"].fillna(0).gt(0)
agreement_summary["has_fh_subscription"] = agreement_summary["agreement_status"].notna()

agreement_summary.head()

# %%
agreement_summary[[
    "sf_recurring_donation_id",
    "dr_name",
    "agreement_id",
    "sf_dr_status",
    "agreement_status",
    "transactions_count",
    "completed_count",
    "refunded_count",
    "first_charge_date",
    "last_charge_date",
    "sf_last_payment_date",
]]

# %%
payment_counts = (
    fh_orders
    .groupby(["order_status", "payment_state"], dropna=False)
    .size()
    .reset_index(name="records")
)

payment_counts

# %%
tx_summary = (
    fh_orders
    .groupby("agreement_id")
    .agg(
        transactions_count=("transaction_id", "nunique"),
        first_charge_date=("completed_date", "min"),
        last_charge_date=("completed_date", "max"),
        total_charged_amount=("donation_amount", "sum"),
        completed_count=("payment_state", lambda s: (s == "Completed").sum()),
        refunded_count=("payment_state", lambda s: (s == "Refunded").sum()),
    )
    .reset_index()
)

tx_summary.head()

# %%
tx_summary.shape

# %%
agreement_summary = yellow_clean.merge(
    tx_summary,
    on="agreement_id",
    how="left"
)

agreement_summary = agreement_summary.merge(
    fh_subscriptions,
    on="agreement_id",
    how="left",
    suffixes=("", "_fh_subscription")
)

agreement_summary["has_fh_orders"] = agreement_summary["transactions_count"].fillna(0).gt(0)
agreement_summary["has_fh_subscription"] = agreement_summary["agreement_status"].notna()

agreement_summary.head()

# %%
agreement_summary_review = agreement_summary[[
    "sf_recurring_donation_id",
    "dr_name",
    "agreement_id",
    "sf_campaign_name",
    "sf_dr_status",
    "sf_last_payment_date",
    "agreement_status",
    "frequency",
    "amount",
    "currency",
    "next_debit_date",
    "transactions_count",
    "completed_count",
    "refunded_count",
    "first_charge_date",
    "last_charge_date",
    "has_fh_orders",
    "has_fh_subscription",
]].copy()

agreement_summary_review

# %%
agreement_summary_review[
    agreement_summary_review["has_fh_orders"] == False
]

# %%
agreement_summary_review[
    agreement_summary_review["has_fh_subscription"] == False
]

# %%
agreement_summary_review["agreement_status"].value_counts(dropna=False)

# %%
output_path = OUTPUT_DIR / "03_paymenthub_yellow_cross_review.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    yellow_clean.to_excel(writer, sheet_name="yellow_clean", index=False)
    fh_orders.to_excel(writer, sheet_name="fh_orders", index=False)
    fh_subscriptions.to_excel(writer, sheet_name="fh_subscriptions", index=False)
    orders_joined.to_excel(writer, sheet_name="orders_joined", index=False)
    payment_counts.to_excel(writer, sheet_name="payment_counts", index=False)
    tx_summary.to_excel(writer, sheet_name="tx_summary", index=False)
    agreement_summary_review.to_excel(writer, sheet_name="agreement_summary", index=False)

print(f"Archivo creado: {output_path}")