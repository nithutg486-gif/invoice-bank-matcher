import io
import re
import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from rapidfuzz import fuzz, process

GSTIN_REGEX = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$")


def validate_gst_compliance(row: pd.Series, has_tax_col: bool, has_treat_col: bool) -> str:
    if not has_tax_col:
        return "Valid"
    gstin = str(row.get("_gstin", "")).strip().upper()
    treatment = str(row.get("_gst_treatment", "")).strip().lower() if has_treat_col else ""

    if treatment == "business_none" and not gstin:
        return "Unregistered B2B (No GSTIN)"
    if not gstin:
        return "Missing GSTIN"
    if not GSTIN_REGEX.match(gstin):
        return f"Invalid GSTIN ({len(gstin)} chars)"
    return "Valid"


def reconcile_ledgers(
    inv_df: pd.DataFrame,
    bank_df: pd.DataFrame,
    inv_map: dict,
    bank_map: dict,
    as_of_date: pd.Timestamp,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inv = inv_df.copy()
    inv["_inv_id"] = inv[inv_map["inv_id"]].astype(str).str.strip()
    inv["_client"] = inv[inv_map["client"]].astype(str).str.strip()
    inv["_amount"] = pd.to_numeric(
        inv[inv_map["amount"]].astype(str).str.replace(r"[^\d.-]", "", regex=True),
        errors="coerce",
    ).fillna(0.0)
    inv["_due_date"] = pd.to_datetime(inv[inv_map["due_date"]], format="mixed", errors="coerce")

    has_tax_col = bool(inv_map.get("tax_id") and inv_map["tax_id"] in inv.columns)
    has_treat_col = bool(inv_map.get("gst_treatment") and inv_map["gst_treatment"] in inv.columns)

    inv["_gstin"] = inv[inv_map["tax_id"]].fillna("").astype(str).str.strip() if has_tax_col else ""
    inv["_gst_treatment"] = (
        inv[inv_map["gst_treatment"]].fillna("").astype(str).str.strip() if has_treat_col else ""
    )

    dup_mask = inv.duplicated(subset=["_inv_id"], keep="first")
    duplicates_df = inv[dup_mask][["_inv_id", "_client", "_amount", "_due_date"]].copy()
    clean_inv = inv[~dup_mask].copy().reset_index(drop=True)

    clean_inv["GST_Compliance"] = clean_inv.apply(
        validate_gst_compliance, axis=1, args=(has_tax_col, has_treat_col)
    )
    clean_inv["Paid_Amount"] = 0.0
    clean_inv["Match_Method"] = "Unpaid"
    clean_inv["Bank_Refs"] = ""

    bank = bank_df.copy()
    bank["_narration"] = bank[bank_map["narration"]].fillna("").astype(str).str.strip()
    bank["_payee"] = (
        bank[bank_map["payee"]].fillna("").astype(str).str.strip()
        if bank_map.get("payee") and bank_map["payee"] in bank.columns
        else ""
    )
    bank["_ref"] = (
        bank[bank_map["ref"]].fillna("").astype(str).str.strip()
        if bank_map.get("ref") and bank_map["ref"] in bank.columns
        else ""
    )
    bank["_txn_date"] = (
        bank[bank_map["date"]].fillna("").astype(str).str.strip()
        if bank_map.get("date") and bank_map["date"] in bank.columns
        else ""
    )
    bank["_credit"] = pd.to_numeric(
        bank[bank_map["credit"]].astype(str).str.replace(r"[^\d.-]", "", regex=True),
        errors="coerce",
    ).fillna(0.0)

    credits = bank[bank["_credit"] > 0].copy().reset_index(drop=True)
    credits["_matched_inv_id"] = None
    credits["_match_type"] = "Unmatched"

    known_ids = sorted(clean_inv["_inv_id"].dropna().unique().tolist(), key=len, reverse=True)
    inv_id_to_idx = {row["_inv_id"]: idx for idx, row in clean_inv.iterrows()}

    def apply_payment(inv_idx: int, amount: float, method: str, ref: str):
        clean_inv.at[inv_idx, "Paid_Amount"] = round(clean_inv.at[inv_idx, "Paid_Amount"] + amount, 2)
        prev_method = clean_inv.at[inv_idx, "Match_Method"]
        clean_inv.at[inv_idx, "Match_Method"] = (
            method if prev_method == "Unpaid" else f"{prev_method} + {method}"
        )
        if ref:
            prev_ref = clean_inv.at[inv_idx, "Bank_Refs"]
            clean_inv.at[inv_idx, "Bank_Refs"] = ref if not prev_ref else f"{prev_ref}, {ref}"

    for b_idx, b_row in credits.iterrows():
        narr_upper = b_row["_narration"].upper()
        for inv_id in known_ids:
            if inv_id and inv_id.upper() in narr_upper:
                inv_idx = inv_id_to_idx[inv_id]
                apply_payment(inv_idx, b_row["_credit"], "Exact ID Match", b_row["_ref"])
                credits.at[b_idx, "_matched_inv_id"] = inv_id
                credits.at[b_idx, "_match_type"] = "Exact ID Match"
                break

    known_clients = clean_inv["_client"].dropna().unique().tolist()
    for b_idx, b_row in credits[credits["_matched_inv_id"].isna()].iterrows():
        search_text = f"{b_row['_payee']} {b_row['_narration']}".strip()
        if not known_clients or not search_text:
            continue

        match = process.extractOne(search_text, known_clients, scorer=fuzz.token_set_ratio)
        if match and match[1] >= 80:
            matched_client = match[0]
            score = int(match[1])
            open_invoices = clean_inv[
                (clean_inv["_client"] == matched_client)
                & ((clean_inv["_amount"] - clean_inv["Paid_Amount"]) > 0.01)
            ].copy()

            if not open_invoices.empty:
                open_invoices["_open_bal"] = (open_invoices["_amount"] - open_invoices["Paid_Amount"]).round(2)
                exact_amt_hits = open_invoices[(open_invoices["_open_bal"] - b_row["_credit"]).abs() <= 1.0]

                if not exact_amt_hits.empty:
                    target_idx = exact_amt_hits.sort_values("_due_date").index[0]
                    tag = f"Fuzzy + Exact Amt ({score}%)"
                else:
                    target_idx = open_invoices.sort_values("_due_date").index[0]
                    tag = f"Fuzzy + FIFO ({score}%)"

                apply_payment(target_idx, b_row["_credit"], tag, b_row["_ref"])
                credits.at[b_idx, "_matched_inv_id"] = clean_inv.at[target_idx, "_inv_id"]
                credits.at[b_idx, "_match_type"] = tag

    clean_inv["Balance_Due"] = (clean_inv["_amount"] - clean_inv["Paid_Amount"]).round(2)
    clean_inv["Days_Overdue"] = (as_of_date - clean_inv["_due_date"]).dt.days.clip(lower=0).fillna(0).astype(int)

    def classify_status(row: pd.Series) -> str:
        gst_issue = row["GST_Compliance"] != "Valid"
        bal = row["Balance_Due"]
        paid = row["Paid_Amount"]

        if bal <= 0 and gst_issue:
            return f"Settled - {row['GST_Compliance']}"
        if bal <= 0:
            return "Reconciled"
        if gst_issue:
            return f"Open - {row['GST_Compliance']}"
        if paid > 0:
            return "Partial Payment"
        return "Unpaid / Overdue"

    def aging_bucket(row: pd.Series) -> str:
        if row["Balance_Due"] <= 0:
            return "Settled"
        if row["Days_Overdue"] <= 30:
            return "0-30 Days"
        if row["Days_Overdue"] <= 60:
            return "31-60 Days"
        return "60+ Days"

    clean_inv["Recon_Status"] = clean_inv.apply(classify_status, axis=1)
    clean_inv["Aging_Bucket"] = clean_inv.apply(aging_bucket, axis=1)

    unmatched_credits = credits[credits["_matched_inv_id"].isna()][
        ["_txn_date", "_payee", "_narration", "_ref", "_credit"]
    ].rename(
        columns={
            "_txn_date": "Date",
            "_payee": "Payee",
            "_narration": "Description",
            "_ref": "Reference",
            "_credit": "Deposit_Amount",
        }
    )

    output_cols = [
        "_inv_id",
        "_client",
        "_due_date",
        "_amount",
        "Paid_Amount",
        "Balance_Due",
        "Days_Overdue",
        "Aging_Bucket",
        "GST_Compliance",
        "Recon_Status",
        "Match_Method",
        "Bank_Refs",
    ]
    ledger_out = clean_inv[output_cols].rename(
        columns={
            "_inv_id": "Invoice_Number",
            "_client": "Customer_Name",
            "_due_date": "Due_Date",
            "_amount": "Billed_Amount",
        }
    )
    ledger_out["Due_Date"] = ledger_out["Due_Date"].dt.strftime("%Y-%m-%d")

    return ledger_out, unmatched_credits, duplicates_df


def export_mis_workbook(
    recon: pd.DataFrame, unmatched_credits: pd.DataFrame, duplicates_df: pd.DataFrame
) -> bytes:
    buf = io.BytesIO()
    bucket_order = ["0-30 Days", "31-60 Days", "60+ Days"]
    aging = (
        recon[recon["Balance_Due"] > 0]
        .groupby("Aging_Bucket", as_index=False)["Balance_Due"]
        .sum()
    )
    aging["Aging_Bucket"] = pd.Categorical(aging["Aging_Bucket"], categories=bucket_order, ordered=True)
    aging = aging.sort_values("Aging_Bucket").reset_index(drop=True)

    exceptions = recon[recon["Recon_Status"] != "Reconciled"]

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        aging.to_excel(writer, sheet_name="MIS_Summary", index=False, startrow=2)
        recon.to_excel(writer, sheet_name="Reconciled_Ledger", index=False)
        exceptions.to_excel(writer, sheet_name="Audit_Exceptions", index=False)
        unmatched_credits.to_excel(writer, sheet_name="Unmatched_Bank_Credits", index=False)
        duplicates_df.to_excel(writer, sheet_name="Duplicate_Invoices", index=False)

        ws = writer.book["MIS_Summary"]
        ws["A1"] = "Accounts Receivable (AR) Aging & Compliance MIS Summary"
        ws["A1"].font = Font(size=13, bold=True)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        for row in ws.iter_rows(min_row=3, max_row=3):
            for cell in row:
                cell.font = Font(color="FFFFFF", bold=True)
                cell.fill = header_fill

        if not aging.empty:
            chart = BarChart()
            chart.title = "Pending Receivables by Aging Bucket"
            chart.y_axis.title = "Balance Due (INR)"
            data = Reference(ws, min_col=2, min_row=3, max_row=3 + len(aging))
            cats = Reference(ws, min_col=1, min_row=4, max_row=3 + len(aging))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws.add_chart(chart, "D3")

    return buf.getvalue()