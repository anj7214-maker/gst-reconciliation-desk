from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from gst_rules import (
    BANK_COLUMNS,
    CANONICAL_COLUMNS,
    DOCUMENT_COLUMNS,
    calculate_gst,
    concat_frames,
    document_review_to_transactions,
    gst_summary,
    make_excel_report,
    manual_transaction,
    match_sales_receipts,
    read_bank_statement,
    read_document_file,
    read_register,
    reconcile_registers,
    summary_metrics,
    validate_register,
)


DATA_DIR = Path("data")
MANUAL_FILE = DATA_DIR / "manual_transactions.csv"
BANK_FILE = DATA_DIR / "bank_entries.csv"
BANK_CATEGORIES = [
    "Sales receipt",
    "Vendor payment",
    "Owner transfer",
    "Loan",
    "Refund",
    "GST payment",
    "Bank charges",
    "Ignore",
    "Needs review",
]

st.set_page_config(page_title="GST Reconciliation Desk", page_icon="GST", layout="wide")


def main() -> None:
    st.title("GST Reconciliation Desk")
    st.caption("Local-first GST register cleanup, manual billing, bank review, and reconciliation.")

    DATA_DIR.mkdir(exist_ok=True)
    manual_saved = load_csv(MANUAL_FILE, CANONICAL_COLUMNS)
    bank_saved = load_csv(BANK_FILE, BANK_COLUMNS)

    with st.sidebar:
        st.header("Upload")
        sales_files = st.file_uploader("Sales register files", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
        purchase_files = st.file_uploader(
            "Purchase register / books files",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
        )
        counterparty_files = st.file_uploader(
            "GSTR-2B / counterparty files",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
        )
        bank_files = st.file_uploader("Bank statement files", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
        document_files = st.file_uploader(
            "Bills / vouchers",
            type=["pdf", "docx", "doc", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
        )

    sales_uploaded, sales_warnings = load_register_files(sales_files, source_prefix="Sales", invoice_type="Sales")
    purchase_uploaded, purchase_warnings = load_register_files(
        purchase_files,
        source_prefix="Books",
        invoice_type="Purchase",
    )
    counterparty, counterparty_warnings = load_register_files(
        counterparty_files,
        source_prefix="GSTR-2B",
        invoice_type="Purchase",
    )
    bank_uploaded, bank_warnings = load_bank_files(bank_files)
    document_review, document_warnings = load_document_files(document_files)

    bank_entries = merge_bank_entries(bank_saved, bank_uploaded)
    manual_sales = manual_saved[manual_saved["invoice_type"] == "Sales"] if not manual_saved.empty else manual_saved
    manual_purchases = (
        manual_saved[manual_saved["invoice_type"] == "Purchase"] if not manual_saved.empty else manual_saved
    )
    sales = concat_frames([sales_uploaded, manual_sales])
    purchases = concat_frames([purchase_uploaded, manual_purchases])
    all_registers = concat_frames([sales, purchases, counterparty])

    issues = validate_register(all_registers)
    reconciliation = reconcile_registers(purchases, counterparty)
    bank_review = match_sales_receipts(sales, bank_entries)
    summary_frame = gst_summary(sales, purchases, bank_review, issues)

    render_warnings(sales_warnings + purchase_warnings + counterparty_warnings + bank_warnings + document_warnings)
    render_dashboard(sales, purchases, counterparty, issues, reconciliation, bank_review, summary_frame)
    render_tabs(
        sales_uploaded=sales_uploaded,
        purchase_uploaded=purchase_uploaded,
        counterparty=counterparty,
        manual_saved=manual_saved,
        bank_entries=bank_entries,
        document_files=document_files,
        document_review=document_review,
        all_registers=all_registers,
        issues=issues,
        reconciliation=reconciliation,
        bank_review=bank_review,
        summary_frame=summary_frame,
    )


def load_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns)
    frame = pd.read_csv(path)
    for column in columns:
        if column not in frame:
            frame[column] = pd.NA
    return frame


def save_csv(frame: pd.DataFrame, path: Path) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    frame.to_csv(path, index=False)


def load_register_files(files, source_prefix: str, invoice_type: str) -> tuple[pd.DataFrame, list[str]]:
    frames: list[pd.DataFrame] = []
    warnings: list[str] = []

    for index, file in enumerate(files or [], start=1):
        source = f"{source_prefix} {index}: {file.name}"
        try:
            result = read_register(file, source=source, invoice_type=invoice_type)
            frames.append(result.frame)
            warnings.extend([f"{file.name}: {warning}" for warning in result.warnings])
        except Exception as exc:  # pragma: no cover - Streamlit display path
            warnings.append(f"{file.name}: {exc}")

    return concat_frames(frames), warnings


def load_bank_files(files) -> tuple[pd.DataFrame, list[str]]:
    frames: list[pd.DataFrame] = []
    warnings: list[str] = []
    for file in files or []:
        try:
            result = read_bank_statement(file)
            frames.append(result.frame)
            warnings.extend([f"{file.name}: {warning}" for warning in result.warnings])
        except Exception as exc:  # pragma: no cover - Streamlit display path
            warnings.append(f"{file.name}: {exc}")
    return concat_frames(frames), warnings


def load_document_files(files) -> tuple[pd.DataFrame, list[str]]:
    frames: list[pd.DataFrame] = []
    warnings: list[str] = []
    for file in files or []:
        try:
            result = read_document_file(file)
            frames.append(result.frame)
            warnings.extend(result.warnings)
        except Exception as exc:  # pragma: no cover - Streamlit display path
            warnings.append(f"{file.name}: {exc}")
    usable = [frame for frame in frames if not frame.empty]
    if not usable:
        return pd.DataFrame(columns=DOCUMENT_COLUMNS), warnings
    return pd.concat(usable, ignore_index=True), warnings


def merge_bank_entries(saved: pd.DataFrame, uploaded: pd.DataFrame) -> pd.DataFrame:
    if saved.empty and uploaded.empty:
        return pd.DataFrame(columns=BANK_COLUMNS)
    merged = pd.concat([saved, uploaded], ignore_index=True)
    merged = merged.drop_duplicates(subset=["entry_id"], keep="first")
    for column in BANK_COLUMNS:
        if column not in merged:
            merged[column] = pd.NA
    return merged[BANK_COLUMNS]


def render_warnings(warnings: list[str]) -> None:
    if warnings:
        with st.expander("Import warnings", expanded=True):
            for warning in warnings:
                st.warning(warning)


def render_dashboard(
    sales: pd.DataFrame,
    purchases: pd.DataFrame,
    counterparty: pd.DataFrame,
    issues: pd.DataFrame,
    reconciliation: pd.DataFrame,
    bank_review: pd.DataFrame,
    summary_frame: pd.DataFrame,
) -> None:
    sales_summary = summary_metrics(sales)
    purchase_summary = summary_metrics(purchases)
    counterparty_summary = summary_metrics(counterparty)
    summary_lookup = dict(zip(summary_frame["metric"], summary_frame["value"], strict=False))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Approved sales", format_money(sales_summary["Invoice Total"]))
    c2.metric("Possible bank sales", format_money(summary_lookup.get("Possible bank sales receipts", 0)))
    c3.metric("Approved purchases", format_money(purchase_summary["Invoice Total"]))
    c4.metric("Net GST payable", format_money(summary_lookup.get("Net GST payable", 0)))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Output tax", format_money(sales_summary["Output/Input Tax"]))
    c6.metric("Input tax credit", format_money(purchase_summary["Output/Input Tax"]))
    c7.metric("GSTR-2B total", format_money(counterparty_summary["Invoice Total"]))
    c8.metric("Unmatched bank sales", str((bank_review["match_status"] == "Unmatched").sum() if not bank_review.empty else 0))

    if not reconciliation.empty:
        st.caption(f"Validation issues: {len(issues)} | Unresolved purchase reconciliation: {(reconciliation['status'] != 'Matched').sum()}")


def render_tabs(
    sales_uploaded: pd.DataFrame,
    purchase_uploaded: pd.DataFrame,
    counterparty: pd.DataFrame,
    manual_saved: pd.DataFrame,
    bank_entries: pd.DataFrame,
    document_files,
    document_review: pd.DataFrame,
    all_registers: pd.DataFrame,
    issues: pd.DataFrame,
    reconciliation: pd.DataFrame,
    bank_review: pd.DataFrame,
    summary_frame: pd.DataFrame,
) -> None:
    tabs = st.tabs(
        [
            "Upload Registers",
            "Documents",
            "Manual Entry",
            "Bank Statement",
            "GST Summary",
            "Validation",
            "Reconciliation",
            "Export",
        ]
    )

    with tabs[0]:
        render_upload_registers(sales_uploaded, purchase_uploaded, counterparty, all_registers)

    with tabs[1]:
        render_documents(document_files, document_review)

    with tabs[2]:
        render_manual_entry(manual_saved)

    with tabs[3]:
        render_bank_statement(bank_entries, bank_review)

    with tabs[4]:
        render_gst_summary(summary_frame, bank_review)

    with tabs[5]:
        render_validation(issues)

    with tabs[6]:
        render_reconciliation(reconciliation)

    with tabs[7]:
        st.download_button(
            "Download GST review workbook",
            data=make_excel_report(
                all_registers,
                issues,
                reconciliation,
                manual_transactions=manual_saved,
                bank_entries=bank_entries,
                bank_sales_review=bank_review,
                gst_summary_frame=summary_frame,
                document_review=document_review,
            ),
            file_name="gst_review_workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def render_upload_registers(
    sales_uploaded: pd.DataFrame,
    purchase_uploaded: pd.DataFrame,
    counterparty: pd.DataFrame,
    all_registers: pd.DataFrame,
) -> None:
    if all_registers.empty:
        st.info("Upload registers, add manual entries, or upload a bank statement to start.")
        st.write(
            "Accepted register columns include Invoice No, Invoice Date, GSTIN, Taxable Value, IGST, "
            "CGST, SGST, Cess, Total, Supplier Name, and Customer Name."
        )
        st.dataframe(pd.DataFrame(columns=CANONICAL_COLUMNS), use_container_width=True)
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Sales rows uploaded", len(sales_uploaded))
    c2.metric("Purchase rows uploaded", len(purchase_uploaded))
    c3.metric("GSTR-2B rows uploaded", len(counterparty))
    st.dataframe(all_registers, use_container_width=True, hide_index=True)


def render_documents(document_files, document_review: pd.DataFrame) -> None:
    st.warning(
        "PDF and Word extraction works only when text is selectable. JPG/PNG files are accepted for manual review; OCR comes later."
    )
    if document_review.empty:
        st.info("Upload PDF, DOCX, TXT, JPG, or PNG bills/vouchers from the sidebar.")
        st.dataframe(pd.DataFrame(columns=DOCUMENT_COLUMNS), use_container_width=True)
        return

    image_files = [file for file in document_files or [] if file.name.lower().endswith((".jpg", ".jpeg", ".png"))]
    if image_files:
        with st.expander("Image previews", expanded=False):
            for file in image_files:
                st.image(file, caption=file.name, use_container_width=True)

    edited = st.data_editor(
        document_review,
        use_container_width=True,
        hide_index=True,
        column_config={
            "review_status": st.column_config.SelectboxColumn(
                "review_status",
                options=["Needs review", "Approve", "Ignore"],
                required=True,
            ),
            "invoice_type": st.column_config.SelectboxColumn(
                "invoice_type",
                options=["Purchase", "Sales"],
                required=True,
            ),
            "document_type": st.column_config.SelectboxColumn(
                "document_type",
                options=[
                    "Invoice",
                    "Purchase invoice",
                    "Sales invoice",
                    "Payment voucher",
                    "Receipt voucher",
                    "Journal voucher",
                    "Credit note",
                    "Debit note",
                    "Voucher",
                    "Unknown",
                ],
                required=True,
            ),
        },
        disabled=["file_name", "file_type", "extraction_status", "text_preview"],
        key="document_review_editor",
    )

    approved_count = int((edited["review_status"] == "Approve").sum()) if "review_status" in edited else 0
    if st.button(f"Save approved documents ({approved_count})"):
        transactions = document_review_to_transactions(edited)
        if transactions.empty:
            st.info("Mark at least one document as Approve before saving.")
            return
        existing = load_csv(MANUAL_FILE, CANONICAL_COLUMNS)
        save_csv(pd.concat([existing, transactions], ignore_index=True), MANUAL_FILE)
        st.success("Approved document transactions saved into GST records.")
        st.rerun()


def render_manual_entry(manual_saved: pd.DataFrame) -> None:
    st.warning("Manual entries are treated as approved GST records. Enter only reviewed bill details.")
    purchase_col, sales_col = st.columns(2)
    with purchase_col:
        render_transaction_form("Purchase")
    with sales_col:
        render_transaction_form("Sales")

    st.subheader("Saved manual transactions")
    if manual_saved.empty:
        st.info("No manual transactions saved yet.")
    else:
        st.dataframe(manual_saved, use_container_width=True, hide_index=True)
        if st.button("Clear manual transactions"):
            save_csv(pd.DataFrame(columns=manual_saved.columns), MANUAL_FILE)
            st.rerun()


def render_transaction_form(invoice_type: str) -> None:
    party_label = "Supplier" if invoice_type == "Purchase" else "Customer"
    with st.form(f"{invoice_type.lower()}_manual_form", clear_on_submit=True):
        st.subheader(f"Manual {invoice_type.lower()} entry")
        party_name = st.text_input(f"{party_label} name")
        gstin = st.text_input(f"{party_label} GSTIN" + (" optional" if invoice_type == "Sales" else ""))
        invoice_no = st.text_input("Invoice number")
        invoice_date = st.date_input("Invoice date")
        taxable_value = st.number_input("Taxable value", min_value=0.0, step=100.0, format="%.2f")
        gst_type = st.selectbox("GST type", ["CGST+SGST", "IGST"])
        gst_rate = st.selectbox("GST rate", [0.0, 5.0, 12.0, 18.0, 28.0], index=3)
        cess = st.number_input("Cess", min_value=0.0, step=10.0, format="%.2f")
        tax = calculate_gst(taxable_value, gst_type, gst_rate, cess)
        st.caption(
            f"Calculated: IGST {format_money(tax['igst'])}, CGST {format_money(tax['cgst'])}, "
            f"SGST {format_money(tax['sgst'])}, Total {format_money(tax['total'])}"
        )
        payment_received = None
        bank_reference = ""
        if invoice_type == "Sales":
            payment_received = st.checkbox("Payment received")
            bank_reference = st.text_input("Bank reference optional")
        notes = st.text_area("Notes")
        submitted = st.form_submit_button(f"Save {invoice_type.lower()} entry")

    if submitted:
        if not party_name or not invoice_no:
            st.error("Party name and invoice number are required.")
            return
        existing = load_csv(MANUAL_FILE, CANONICAL_COLUMNS)
        new_row = manual_transaction(
            invoice_type=invoice_type,
            party_name=party_name,
            gstin=gstin,
            invoice_no=invoice_no,
            invoice_date=invoice_date,
            taxable_value=taxable_value,
            gst_type=gst_type,
            gst_rate=gst_rate,
            cess=cess,
            notes=notes,
            payment_received=payment_received,
            bank_reference=bank_reference,
        )
        save_csv(pd.concat([existing, new_row], ignore_index=True), MANUAL_FILE)
        st.success(f"Saved manual {invoice_type.lower()} entry.")
        st.rerun()


def render_bank_statement(bank_entries: pd.DataFrame, bank_review: pd.DataFrame) -> None:
    st.warning("Bank credits are not GST sales until you manually categorize them as Sales receipt.")
    if bank_entries.empty:
        st.info("Upload a bank statement from the sidebar to review receipts and payments.")
        st.dataframe(pd.DataFrame(columns=BANK_COLUMNS), use_container_width=True)
        return

    edited = st.data_editor(
        bank_entries,
        use_container_width=True,
        hide_index=True,
        column_config={
            "category": st.column_config.SelectboxColumn("category", options=BANK_CATEGORIES, required=True),
        },
        disabled=["entry_id", "date", "description", "reference", "debit", "credit", "balance", "suggested_category"],
        key="bank_entries_editor",
    )
    if st.button("Save bank review"):
        save_csv(edited, BANK_FILE)
        st.success("Bank review saved.")
        st.rerun()

    st.subheader("Bank sales review")
    if bank_review.empty:
        st.info("Mark credit entries as Sales receipt to include them in possible bank sales.")
    else:
        st.dataframe(bank_review, use_container_width=True, hide_index=True)


def render_gst_summary(summary_frame: pd.DataFrame, bank_review: pd.DataFrame) -> None:
    st.dataframe(summary_frame, use_container_width=True, hide_index=True)
    st.caption("Possible bank sales are for review only. GST output tax is calculated from approved sales invoices/manual sales.")
    unmatched = bank_review[bank_review["match_status"] == "Unmatched"] if not bank_review.empty else pd.DataFrame()
    if unmatched.empty:
        st.success("No unmatched bank sales receipts from the reviewed bank entries.")
    else:
        st.subheader("Unmatched bank sales receipts")
        st.dataframe(unmatched, use_container_width=True, hide_index=True)


def render_validation(issues: pd.DataFrame) -> None:
    if issues.empty:
        st.success("No validation issues found from the current rule set.")
        return
    severity = st.multiselect(
        "Severity",
        options=sorted(issues["severity"].unique()),
        default=sorted(issues["severity"].unique()),
    )
    st.dataframe(issues[issues["severity"].isin(severity)], use_container_width=True, hide_index=True)


def render_reconciliation(reconciliation: pd.DataFrame) -> None:
    if reconciliation.empty:
        st.info("Upload purchase/books and GSTR-2B/counterparty files to reconcile.")
        return
    status = st.multiselect(
        "Status",
        options=sorted(reconciliation["status"].unique()),
        default=sorted(reconciliation["status"].unique()),
    )
    st.dataframe(reconciliation[reconciliation["status"].isin(status)], use_container_width=True, hide_index=True)


def format_money(value: float) -> str:
    return f"Rs {value:,.2f}"


if __name__ == "__main__":
    main()
