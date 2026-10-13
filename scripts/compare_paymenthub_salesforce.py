# %%
import pandas as pd
from pathlib import Path

# %%
# =========================
# 1. RUTAS DEL PROYECTO
# =========================

BASE_DIR = Path.cwd()

if BASE_DIR.name == "scripts":
    BASE_DIR = BASE_DIR.parent

INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

orders_path = INPUT_DIR / "fh_orders_completed_and_refunded.csv"
sf_opps_path = OUTPUT_DIR / "04_salesforce_existing_opportunities.xlsx"
excel_path = INPUT_DIR / "FH-SF Sync failures.xlsx"

orders_path, sf_opps_path, excel_path

# %%
# =========================
# 2. CARGAR ARCHIVOS
# =========================

fh_orders_raw = pd.read_csv(orders_path, dtype=str)

sf_opps_raw = pd.read_excel(
    sf_opps_path,
    sheet_name="sf_opportunities",
    dtype=str
)

last_opp_by_dr_raw = pd.read_excel(
    sf_opps_path,
    sheet_name="last_opp_by_dr",
    dtype=str
)

yellow_raw = pd.read_excel(
    excel_path,
    sheet_name="Yellow"
)

fh_orders_raw.shape, sf_opps_raw.shape, last_opp_by_dr_raw.shape, yellow_raw.shape

# %%
# =========================
# 3. LIMPIAR PaymentHub ORDERS
# =========================

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
# Validaciones PaymentHub

fh_orders.shape
# %%
fh_orders["agreement_id"].nunique()

# %%
fh_orders[["order_status", "payment_state"]].value_counts(dropna=False)

# %%
# =========================
# 4. LIMPIAR SALESFORCE OPPORTUNITIES
# =========================

sf_opps = sf_opps_raw.copy()

sf_opps["sf_recurring_donation_id"] = (
    sf_opps["sf_recurring_donation_id"]
    .astype(str)
    .str.strip()
)

sf_opps["sf_auth"] = (
    sf_opps["sf_auth"]
    .astype(str)
    .str.strip()
)

sf_opps["sf_auth"] = sf_opps["sf_auth"].replace({
    "nan": pd.NA,
    "None": pd.NA,
    "NaT": pd.NA,
    "": pd.NA,
})

sf_opps["sf_close_date"] = pd.to_datetime(
    sf_opps["sf_close_date"],
    errors="coerce"
)

sf_opps["sf_amount"] = pd.to_numeric(
    sf_opps["sf_amount"],
    errors="coerce"
)

sf_opps["sf_installment_number"] = pd.to_numeric(
    sf_opps["sf_installment_number"],
    errors="coerce"
)

sf_opps.head()

# %%
# Validaciones Salesforce

sf_opps.shape

# %%
sf_opps["sf_recurring_donation_id"].nunique()

# %%
sf_opps["sf_auth"].notna().sum(), sf_opps["sf_auth"].isna().sum()

# %%
sf_auth_duplicates = (
    sf_opps[sf_opps["sf_auth"].notna()]
    .groupby("sf_auth")
    .size()
    .reset_index(name="count")
    .query("count > 1")
)

sf_auth_duplicates

# %%
# =========================
# 5. LIMPIAR ÚLTIMA OPPORTUNITY POR DR
# =========================

last_opp_by_dr = last_opp_by_dr_raw.copy()

last_opp_by_dr["sf_recurring_donation_id"] = (
    last_opp_by_dr["sf_recurring_donation_id"]
    .astype(str)
    .str.strip()
)

last_opp_by_dr["sf_auth"] = (
    last_opp_by_dr["sf_auth"]
    .astype(str)
    .str.strip()
)

last_opp_by_dr["sf_auth"] = last_opp_by_dr["sf_auth"].replace({
    "nan": pd.NA,
    "None": pd.NA,
    "NaT": pd.NA,
    "": pd.NA,
})

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
# Validación última Opportunity por DR

last_opp_by_dr.shape

# %%
last_opp_by_dr[[
    "sf_recurring_donation_id",
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_close_date",
    "sf_amount",
    "sf_currency",
    "sf_auth",
    "sf_installment_number",
]].head()

# %%
# =========================
# 6. LIMPIAR YELLOW / MAPA AGREEMENT -> DR
# =========================

yellow_map = yellow_raw.rename(columns={
    "Id": "sf_recurring_donation_id",
    "Transaction_ID_Unique__c": "agreement_id",
    "Name": "dr_name",
})[[
    "sf_recurring_donation_id",
    "agreement_id",
    "dr_name",
]].copy()

yellow_map["sf_recurring_donation_id"] = (
    yellow_map["sf_recurring_donation_id"]
    .astype(str)
    .str.strip()
)

yellow_map["agreement_id"] = (
    yellow_map["agreement_id"]
    .astype(str)
    .str.strip()
)

yellow_map.head()

# %%
# Validar Yellow Map

yellow_map.shape, yellow_map["agreement_id"].nunique(), yellow_map["sf_recurring_donation_id"].nunique()

# %%
# =========================
# 7. AGREGAR DR A PaymentHub ORDERS
# =========================

fh_orders_with_dr = fh_orders.merge(
    yellow_map,
    on="agreement_id",
    how="left"
)

fh_orders_with_dr.head()

# %%
# Validar que todos los cobros tengan DR

fh_orders_with_dr["sf_recurring_donation_id"].isna().sum()

# %%
# =========================
# 8. PREPARAR PUNTO DE CORTE: ÚLTIMA OPP DE SF
# =========================

last_opp_cutoff = last_opp_by_dr[[
    "sf_recurring_donation_id",
    "sf_opportunity_id",
    "sf_opportunity_name",
    "sf_close_date",
    "sf_amount",
    "sf_currency",
    "sf_auth",
    "sf_installment_number",
]].copy()

last_opp_cutoff = last_opp_cutoff.rename(columns={
    "sf_opportunity_id": "last_sf_opportunity_id",
    "sf_opportunity_name": "last_sf_opportunity_name",
    "sf_close_date": "last_sf_close_date",
    "sf_amount": "last_sf_amount",
    "sf_currency": "last_sf_currency",
    "sf_auth": "last_sf_auth",
    "sf_installment_number": "last_sf_installment_number",
})

last_opp_cutoff.head()

# %%
# =========================
# 9. AGREGAR PUNTO DE CORTE A CADA COBRO FH
# =========================

fh_orders_with_cutoff = fh_orders_with_dr.merge(
    last_opp_cutoff,
    on="sf_recurring_donation_id",
    how="left"
)

fh_orders_with_cutoff[[
    "agreement_id",
    "sf_recurring_donation_id",
    "dr_name",
    "transaction_id",
    "completed_date",
    "donation_amount",
    "last_sf_close_date",
    "last_sf_auth",
    "last_sf_installment_number",
]].head()

# %%
# Validar que todas tengan punto de corte

fh_orders_with_cutoff["last_sf_close_date"].isna().sum()

# %%
# =========================
# 10. FILTRAR SOLO DESDE LA ÚLTIMA OPP EN SF EN ADELANTE
# =========================

fh_from_last_sf_forward = fh_orders_with_cutoff[
    fh_orders_with_cutoff["completed_date"] >= fh_orders_with_cutoff["last_sf_close_date"]
].copy()

fh_from_last_sf_forward.shape

# %%
fh_from_last_sf_forward[[
    "agreement_id",
    "sf_recurring_donation_id",
    "dr_name",
    "transaction_id",
    "payment_state",
    "completed_date",
    "completed_time",
    "donation_amount",
    "currency",
    "last_sf_close_date",
    "last_sf_auth",
    "last_sf_installment_number",
]].sort_values([
    "agreement_id",
    "completed_date",
    "completed_time",
]).head(100)

# %%
# =========================
# 11. EXCLUIR TRANSACCIONES QUE YA EXISTEN EN SF POR AUTH__C
# =========================

existing_sf_auths = (
    sf_opps["sf_auth"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

len(existing_sf_auths)

# %%
fh_from_last_sf_forward["exists_in_salesforce_by_auth"] = (
    fh_from_last_sf_forward["transaction_id"]
    .astype(str)
    .str.strip()
    .isin(existing_sf_auths)
)

fh_from_last_sf_forward["exists_in_salesforce_by_auth"].value_counts(dropna=False)

# %%
# =========================
# 12. CANDIDATOS REALES A CREAR
# =========================

missing_transactions_after_last_sf = fh_from_last_sf_forward[
    fh_from_last_sf_forward["exists_in_salesforce_by_auth"] == False
].copy()

missing_transactions_after_last_sf.shape

# %%
missing_transactions_after_last_sf[[
    "agreement_id",
    "sf_recurring_donation_id",
    "dr_name",
    "transaction_id",
    "payment_state",
    "completed_date",
    "completed_time",
    "donation_amount",
    "currency",
    "last_sf_close_date",
    "last_sf_auth",
    "last_sf_installment_number",
]].sort_values([
    "agreement_id",
    "completed_date",
    "completed_time",
]).head(100)

# %%
# =========================
# 13. RESUMEN POR AGREEMENT / DR
# =========================

missing_summary_by_agreement = (
    missing_transactions_after_last_sf
    .groupby([
        "agreement_id",
        "sf_recurring_donation_id",
        "dr_name",
        "last_sf_close_date",
        "last_sf_installment_number",
    ], dropna=False)
    .agg(
        missing_transactions=("transaction_id", "nunique"),
        first_missing_charge_date=("completed_date", "min"),
        last_missing_charge_date=("completed_date", "max"),
        total_missing_amount=("donation_amount", "sum"),
        completed_missing=("payment_state", lambda s: (s == "Completed").sum()),
        refunded_missing=("payment_state", lambda s: (s == "Refunded").sum()),
    )
    .reset_index()
    .sort_values([
        "agreement_id",
        "first_missing_charge_date",
    ])
)

missing_summary_by_agreement

# %%
# =========================
# 14. CONTROL DE CALIDAD: RESUMEN DE FALTANTES
# =========================

missing_summary_by_agreement_sorted = missing_summary_by_agreement.sort_values(
    ["missing_transactions", "first_missing_charge_date"],
    ascending=[False, True]
)

missing_summary_by_agreement_sorted

# %%
# =========================
# 15. CONTROL DE CALIDAD: AGREEMENTS SIN FALTANTES
# =========================

agreements_with_missing = set(missing_summary_by_agreement["agreement_id"])
all_agreements = set(yellow_map["agreement_id"])

agreements_without_missing = sorted(all_agreements - agreements_with_missing)

len(agreements_without_missing), agreements_without_missing

# %%
# =========================
# 16. CONTROL DE CALIDAD: DÍAS DESDE LA ÚLTIMA OPP DE SF
# =========================

missing_transactions_after_last_sf["days_after_last_sf"] = (
    missing_transactions_after_last_sf["completed_date"]
    - missing_transactions_after_last_sf["last_sf_close_date"]
).dt.days

missing_transactions_after_last_sf["days_after_last_sf"].value_counts().sort_index()

# %%
# =========================
# 17. FALTANTES ESTRICTAMENTE POSTERIORES A LA ÚLTIMA OPP DE SF
# =========================

missing_transactions_strictly_after_last_sf = missing_transactions_after_last_sf[
    missing_transactions_after_last_sf["completed_date"] > missing_transactions_after_last_sf["last_sf_close_date"]
].copy()

missing_transactions_strictly_after_last_sf.shape

# %%
# =========================
# 18. RESUMEN DE FALTANTES ESTRICTAMENTE POSTERIORES
# =========================

missing_summary_strictly_after = (
    missing_transactions_strictly_after_last_sf
    .groupby([
        "agreement_id",
        "sf_recurring_donation_id",
        "dr_name",
        "last_sf_close_date",
        "last_sf_installment_number",
    ], dropna=False)
    .agg(
        missing_transactions=("transaction_id", "nunique"),
        first_missing_charge_date=("completed_date", "min"),
        last_missing_charge_date=("completed_date", "max"),
        total_missing_amount=("donation_amount", "sum"),
        completed_missing=("payment_state", lambda s: (s == "Completed").sum()),
        refunded_missing=("payment_state", lambda s: (s == "Refunded").sum()),
    )
    .reset_index()
    .sort_values([
        "agreement_id",
        "first_missing_charge_date",
    ])
)

missing_summary_strictly_after

# %%
# =========================
# 19. REVISAR FALTANTES DEL MISMO DÍA QUE LA ÚLTIMA OPP DE SF
# =========================

same_day_missing = missing_transactions_after_last_sf[
    missing_transactions_after_last_sf["completed_date"]
    == missing_transactions_after_last_sf["last_sf_close_date"]
].copy()

same_day_missing[[
    "agreement_id",
    "sf_recurring_donation_id",
    "dr_name",
    "transaction_id",
    "payment_state",
    "completed_date",
    "completed_time",
    "donation_amount",
    "currency",
    "last_sf_close_date",
    "last_sf_auth",
    "last_sf_installment_number",
]].sort_values([
    "agreement_id",
    "completed_date",
    "completed_time",
])

# %%
# =========================
# 20. EXPORTAR REVISIÓN AMPLIADA
# =========================

output_path = OUTPUT_DIR / "05_paymenthub_salesforce_after_last_sf_review.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    fh_orders.to_excel(writer, sheet_name="fh_orders", index=False)
    sf_opps.to_excel(writer, sheet_name="sf_opportunities", index=False)
    last_opp_by_dr.to_excel(writer, sheet_name="last_opp_by_dr", index=False)
    fh_orders_with_cutoff.to_excel(writer, sheet_name="fh_orders_with_cutoff", index=False)
    fh_from_last_sf_forward.to_excel(writer, sheet_name="fh_from_last_sf_forward", index=False)
    missing_transactions_after_last_sf.to_excel(writer, sheet_name="missing_after_last_sf", index=False)
    missing_summary_by_agreement.to_excel(writer, sheet_name="missing_summary", index=False)
    missing_transactions_strictly_after_last_sf.to_excel(writer, sheet_name="missing_strictly_after", index=False)
    missing_summary_strictly_after.to_excel(writer, sheet_name="summary_strictly_after", index=False)
    same_day_missing.to_excel(writer, sheet_name="same_day_missing", index=False)

print(f"Archivo creado: {output_path}")

# %%
# =========================
# 21. DEFINIR BASE PRELIMINAR PARA CREACIÓN
# =========================

opportunity_candidates_base = missing_transactions_strictly_after_last_sf.copy()

opportunity_candidates_base = opportunity_candidates_base.sort_values([
    "agreement_id",
    "completed_date",
    "completed_time",
]).copy()

opportunity_candidates_base.shape

# %%
# =========================
# 22. RESUMEN FINAL DE CANDIDATOS POR AGREEMENT
# =========================

candidate_summary = (
    opportunity_candidates_base
    .groupby([
        "agreement_id",
        "sf_recurring_donation_id",
        "dr_name",
        "last_sf_close_date",
        "last_sf_installment_number",
    ], dropna=False)
    .agg(
        candidate_transactions=("transaction_id", "nunique"),
        first_candidate_charge_date=("completed_date", "min"),
        last_candidate_charge_date=("completed_date", "max"),
        total_candidate_amount=("donation_amount", "sum"),
        completed_candidates=("payment_state", lambda s: (s == "Completed").sum()),
        refunded_candidates=("payment_state", lambda s: (s == "Refunded").sum()),
    )
    .reset_index()
    .sort_values([
        "agreement_id",
        "first_candidate_charge_date",
    ])
)

candidate_summary

# %%
# =========================
# 23. CONTROL: AGREEMENTS CON CANDIDATOS VS SIN CANDIDATOS
# =========================

agreements_with_candidates = set(candidate_summary["agreement_id"])
all_agreements = set(yellow_map["agreement_id"])

agreements_without_candidates = sorted(all_agreements - agreements_with_candidates)

len(agreements_with_candidates), len(agreements_without_candidates), agreements_without_candidates

# %%
# =========================
# 24. EXPORTAR ARCHIVO FINAL DE COMPARACIÓN
# =========================

output_path = OUTPUT_DIR / "05_paymenthub_salesforce_after_last_sf_review.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    fh_orders.to_excel(writer, sheet_name="fh_orders", index=False)
    sf_opps.to_excel(writer, sheet_name="sf_opportunities", index=False)
    last_opp_by_dr.to_excel(writer, sheet_name="last_opp_by_dr", index=False)
    fh_orders_with_cutoff.to_excel(writer, sheet_name="fh_orders_with_cutoff", index=False)
    fh_from_last_sf_forward.to_excel(writer, sheet_name="fh_from_last_sf_forward", index=False)
    missing_transactions_after_last_sf.to_excel(writer, sheet_name="missing_after_last_sf", index=False)
    missing_summary_by_agreement.to_excel(writer, sheet_name="missing_summary", index=False)
    missing_transactions_strictly_after_last_sf.to_excel(writer, sheet_name="missing_strictly_after", index=False)
    missing_summary_strictly_after.to_excel(writer, sheet_name="summary_strictly_after", index=False)
    same_day_missing.to_excel(writer, sheet_name="same_day_missing", index=False)
    opportunity_candidates_base.to_excel(writer, sheet_name="opportunity_candidates", index=False)
    candidate_summary.to_excel(writer, sheet_name="candidate_summary", index=False)

print(f"Archivo creado: {output_path}")