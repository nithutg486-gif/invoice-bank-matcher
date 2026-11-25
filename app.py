import datetime
import pandas as pd
import streamlit as st
from reconcile_engine import export_mis_workbook, reconcile_ledgers

st.set_page_config(page_title="Zoho-to-MIS Reconciliation Engine", layout="wide")
st.title("Automated Invoice Reconciliation & MIS Dashboard")


def guess_index(cols: list[str], candidates: list[str], allow_none: bool = False) -> int:
    offset = 1 if allow_none else 0
    lower_map = {c.lower().strip(): i for i, c in enumerate(cols)}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()] + offset
    return 0


c_up1, c_up2 = st.columns(2)
inv_file = c_up1.file_uploader("1. Upload Invoice Ledger CSV (Zoho Books / ERP)", type=["csv"])
bank_file = c_up2.file_uploader("2. Upload Bank Statement CSV", type=["csv"])

if not inv_file or not bank_file:
    st.info("Upload both CSV files above to auto-map columns and run two-pass reconciliation.")
    st.stop()

inv_df = pd.read_csv(inv_file)
bank_df = pd.read_csv(bank_file)
inv_cols = inv_df.columns.tolist()
bank_cols = bank_df.columns.tolist()

st.sidebar.header("Dynamic Schema Mapping")
inv_map = {
    "inv_id": st.sidebar.selectbox(
        "Invoice Number Column",
        inv_cols,
        index=guess_index(inv_cols, ["Invoice Number", "Invoice_ID", "Inv No"]),
    ),
    "client": st.sidebar.selectbox(
        "Customer Name Column",
        inv_cols,
        index=guess_index(inv_cols, ["Customer Name", "Client Name", "Customer"]),
    ),
    "amount": st.sidebar.selectbox(
        "Total Amount Column",
        inv_cols,
        index=guess_index(inv_cols, ["Total", "Grand Total", "Invoice Amount", "Amount"]),
    ),
    "due_date": st.sidebar.selectbox(
        "Due Date Column",
        inv_cols,
        index=guess_index(inv_cols, ["Due Date", "Invoice Date"]),
    ),
    "tax_id": st.sidebar.selectbox(
        "GSTIN / Tax ID Column (Optional)",
        ["None"] + inv_cols,
        index=guess_index(inv_cols, ["GST Identification Number (GSTIN)", "GSTIN", "Tax ID"], allow_none=True),
    ),
    "gst_treatment": st.sidebar.selectbox(
        "GST Treatment Column (Optional)",
        ["None"] + inv_cols,
        index=guess_index(inv_cols, ["GST Treatment"], allow_none=True),
    ),
}
if inv_map["tax_id"] == "None":
    inv_map["tax_id"] = None
if inv_map["gst_treatment"] == "None":
    inv_map["gst_treatment"] = None

bank_map = {
    "narration": st.sidebar.selectbox(
        "Bank Narration / Description Column",
        bank_cols,
        index=guess_index(bank_cols, ["Description", "Narration", "Particulars"]),
    ),
    "credit": st.sidebar.selectbox(
        "Credit / Deposit Amount Column",
        bank_cols,
        index=guess_index(bank_cols, ["Deposits", "Credit", "Deposit Amount"]),
    ),
    "payee": st.sidebar.selectbox(
        "Payee / Remitter Column (Optional)",
        ["None"] + bank_cols,
        index=guess_index(bank_cols, ["Payee", "Remitter", "Party"], allow_none=True),
    ),
    "date": st.sidebar.selectbox(
        "Transaction Date Column (Optional)",
        ["None"] + bank_cols,
        index=guess_index(bank_cols, ["Date", "Txn Date", "Value Date"], allow_none=True),
    ),
    "ref": st.sidebar.selectbox(
        "Reference / UTR Column (Optional)",
        ["None"] + bank_cols,
        index=guess_index(bank_cols, ["Reference Number", "UTR", "Ref No"], allow_none=True),
    ),
}
for k in ["payee", "date", "ref"]:
    if bank_map[k] == "None":
        bank_map[k] = None

as_of = pd.Timestamp(
    st.sidebar.date_input("Reconciliation Cutoff Date", value=datetime.date(2026, 11, 25))
)

recon, unmatched_credits, duplicates_df = reconcile_ledgers(inv_df, bank_df, inv_map, bank_map, as_of)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Unique Invoices", len(recon), delta=f"-{len(duplicates_df)} dupes removed")
m2.metric("Total Billed", f"₹{recon['Billed_Amount'].sum():,.2f}")
m3.metric("Pending Receivables", f"₹{recon['Balance_Due'].sum():,.2f}")
m4.metric("Invoice Exceptions", int((recon["Recon_Status"] != "Reconciled").sum()))
m5.metric("Unmatched Credits", f"₹{unmatched_credits['Deposit_Amount'].sum():,.2f}")

tab1, tab2, tab3 = st.tabs(["Reconciled Ledger", "Unmatched Bank Credits", "Duplicate Invoices Removed"])

with tab1:
    selected_status = st.multiselect(
        "Filter Reconciliation Status",
        options=recon["Recon_Status"].unique().tolist(),
        default=recon["Recon_Status"].unique().tolist(),
    )
    st.dataframe(recon[recon["Recon_Status"].isin(selected_status)], use_container_width=True)

with tab2:
    st.dataframe(unmatched_credits, use_container_width=True)

with tab3:
    st.dataframe(duplicates_df, use_container_width=True)

st.download_button(
    label="Download Formatted 5-Sheet Excel MIS Pack (.xlsx)",
    data=export_mis_workbook(recon, unmatched_credits, duplicates_df),
    file_name="MIS_Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)