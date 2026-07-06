from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from gst_rules import (
    BANK_COLUMNS,
    DOCUMENT_COLUMNS,
    PURCHASE_COLUMNS,
    SALES_COLUMNS,
    build_exceptions,
    calculate_gst,
    concat_frames,
    document_review_to_purchase_sales,
    gst_books_summary,
    make_excel_report,
    manual_purchase_record,
    manual_sales_record,
    match_sales_receipts,
    match_vendor_payments,
    purchase_records_to_register,
    read_bank_statement,
    read_document_file,
    read_register,
    reconcile_registers,
    register_to_purchase_records,
    register_to_sales_records,
    sales_records_to_register,
    validate_register,
)


DATA_DIR = Path("data")
PURCHASE_FILE = DATA_DIR / "purchase_records.csv"
SALES_FILE = DATA_DIR / "sales_records.csv"
BANK_FILE = DATA_DIR / "bank_entries.csv"

BANK_CATEGORIES = [
    "Sales receipt",
    "Vendor payment",
    "Owner transfer",
    "Loan",
    "Refund",
    "GST payment",
    "Bank charges",
    "Salary",
    "Ignore",
    "Needs review",
]


st.set_page_config(page_title="GST Books Builder", page_icon="GST", layout="wide")


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    st.title("GST Books Builder")
    st.caption("Build reviewed GST purchase and sales records from bills, vouchers, and bank statements.")

    saved_purchases = load_csv(PURCHASE_FILE, PURCHASE_COLUMNS)
    saved_sales = load_csv(SALES_FILE, SALES_COLUMNS)
    saved_bank = load_csv(BANK_FILE, BANK_COLUMNS)

    tabs = st.tabs(["Sources", "Review Extraction", "Purchases", "Sales", "Bank", "Exceptions", "GST Summary", "Export"])

    with tabs[0]:
        files = render_sources()

    uploaded_purchase_registers, purchase_warnings = load_register_files(
        files["purchase_registers"],
        "Purchase register",
        "Purchase",
    )
    uploaded_sales_registers, sales_warnings = load_register_files(files["sales_registers"], "Sales register", "Sales")
    gstr_2b_register, gstr_warnings = load_register_files(files["gstr_2b"], "GSTR-2B", "Purchase")
    document_review, document_warnings = load_document_files(files["documents"])
    bank_uploaded, bank_warnings = load_bank_files(files["bank"])

    bank_entries = merge_bank_entries(saved_bank, bank_uploaded)
    purchase_records = concat_record_frames(
        [saved_purchases, register_to_purchase_records(uploaded_purchase_registers)],
        PURCHASE_COLUMNS,
    )
    sales_records = concat_record_frames([saved_sales, register_to_sales_records(uploaded_sales_registers)], SALES_COLUMNS)

    purchase_register = purchase_records_to_register(purchase_records)
    sales_register = sales_records_to_register(sales_records)
    validation_issues = validate_register(concat_frames([purchase_register, sales_register]))
    gstr_2b_reconciliation = reconcile_registers(purchase_register, gstr_2b_register)
    sales_bank_review = match_sales_receipts(sales_records, bank_entries)
    purchase_bank_review = match_vendor_payments(purchase_records, bank_entries)
    exceptions = build_exceptions(
        purchase_records,
        sales_records,
        bank_entries,
        sales_bank_review,
        purchase_bank_review,
        gstr_2b_reconciliation,
        validation_issues,
    )
    summary = gst_books_summary(
        purchase_records,
        sales_records,
        bank_entries,
        sales_bank_review,
        exceptions,
        validation_issues,
    )

    render_warnings(purchase_warnings + sales_warnings + gstr_warnings + document_warnings + bank_warnings)
    render_header_metrics(summary)

    with tabs[1]:
        render_review_extraction(files["documents"], document_review)

    with tabs[2]:
        render_purchases(purchase_records, purchase_bank_review)

    with tabs[3]:
        render_sales(sales_records, sales_bank_review)

    with tabs[4]:
        render_bank(bank_entries, sales_bank_review, purchase_bank_review)

    with tabs[5]:
        render_exceptions(exceptions)

    with tabs[6]:
        render_gst_summary(summary)

    with tabs[7]:
        render_export(
            purchase_records,
            sales_records,
            bank_entries,
            document_review,
            exceptions,
            summary,
            validation_issues,
            gstr_2b_reconciliation,
            sales_bank_review,
        )


def render_sources() -> dict:
    st.subheader("Sources")
    st.write("Upload bills, vouchers, registers, bank statements, and GSTR-2B files here.")
    st.warning(
        "PDF/DOCX/TXT extraction works only for selectable text. JPG/PNG and scanned PDFs require manual review until OCR is added. Bank credits are not sales unless categorized."
    )
    c1, c2 = st.columns(2)
    with c1:
        purchase_registers = st.file_uploader(
            "Purchase bill/register files",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
        )
        sales_registers = st.file_uploader(
            "Sales bill/register files",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
        )
        gstr_2b = st.file_uploader("GSTR-2B / counterparty files", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    with c2:
        documents = st.file_uploader(
            "PDF / Word / image bills and vouchers",
            type=["pdf", "docx", "doc", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
        )
        bank = st.file_uploader("Bank statements", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    return {
        "purchase_registers": purchase_registers,
        "sales_registers": sales_registers,
        "gstr_2b": gstr_2b,
        "documents": documents,
        "bank": bank,
    }


def load_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns)
    frame = pd.read_csv(path)
    for column in columns:
        if column not in frame:
            frame[column] = pd.NA
    return frame[columns]


def save_csv(frame: pd.DataFrame, path: Path) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    frame.to_csv(path, index=False)


def load_register_files(files, source_prefix: str, invoice_type: str) -> tuple[pd.DataFrame, list[str]]:
    frames, warnings = [], []
    for index, file in enumerate(files or [], start=1):
        try:
            result = read_register(file, source=f"{source_prefix} {index}: {file.name}", invoice_type=invoice_type)
            frames.append(result.frame)
            warnings.extend([f"{file.name}: {warning}" for warning in result.warnings])
        except Exception as exc:
            warnings.append(f"{file.name}: {exc}")
    return concat_frames(frames), warnings


def load_document_files(files) -> tuple[pd.DataFrame, list[str]]:
    frames, warnings = [], []
    for file in files or []:
        try:
            result = read_document_file(file)
            frames.append(result.frame)
            warnings.extend(result.warnings)
        except Exception as exc:
            warnings.append(f"{file.name}: {exc}")
    if not frames:
        return pd.DataFrame(columns=DOCUMENT_COLUMNS), warnings
    return pd.concat(frames, ignore_index=True), warnings


def load_bank_files(files) -> tuple[pd.DataFrame, list[str]]:
    frames, warnings = [], []
    for file in files or []:
        try:
            result = read_bank_statement(file)
            frames.append(result.frame)
            warnings.extend([f"{file.name}: {warning}" for warning in result.warnings])
        except Exception as exc:
            warnings.append(f"{file.name}: {exc}")
    if not frames:
        return pd.DataFrame(columns=BANK_COLUMNS), warnings
    return pd.concat(frames, ignore_index=True), warnings


def concat_record_frames(frames: list[pd.DataFrame], columns: list[str]) -> pd.DataFrame:
    usable = [frame for frame in frames if frame is not None and not frame.empty]
    if not usable:
        return pd.DataFrame(columns=columns)
    merged = pd.concat(usable, ignore_index=True)
    for column in columns:
        if column not in merged:
            merged[column] = pd.NA
    id_column = columns[0]
    return merged[columns].drop_duplicates(subset=[id_column], keep="first")


def merge_bank_entries(saved: pd.DataFrame, uploaded: pd.DataFrame) -> pd.DataFrame:
    return concat_record_frames([saved, uploaded], BANK_COLUMNS).drop_duplicates(subset=["entry_id"], keep="first")


def render_warnings(warnings: list[str]) -> None:
    if warnings:
        with st.expander("Import warnings", expanded=True):
            for warning in warnings:
                st.warning(warning)


def render_header_metrics(summary: pd.DataFrame) -> None:
    lookup = dict(zip(summary["metric"], summary["value"], strict=False))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Confirmed sales", format_money(lookup.get("Confirmed sales from approved sales records", 0)))
    c2.metric("Possible unrecorded sales", format_money(lookup.get("Possible unrecorded sales", 0)))
    c3.metric("Approved purchases", format_money(lookup.get("Approved purchases", 0)))
    c4.metric("Net GST payable", format_money(lookup.get("Net GST payable", 0)))


def render_review_extraction(document_files, document_review: pd.DataFrame) -> None:
    st.subheader("Review Extraction")
    st.caption("Approve rows only after checking the extracted bill/voucher fields.")
    if document_review.empty:
        st.info("Upload PDF/DOCX/TXT/JPG/PNG bills or vouchers in Sources.")
        st.dataframe(pd.DataFrame(columns=DOCUMENT_COLUMNS), use_container_width=True)
        return

    image_files = [file for file in document_files or [] if file.name.lower().endswith((".jpg", ".jpeg", ".png"))]
    if image_files:
        with st.expander("Image previews"):
            for file in image_files:
                st.image(file, caption=file.name, use_container_width=True)

    edited = st.data_editor(
        document_review,
        use_container_width=True,
        hide_index=True,
        column_config={
            "review_status": st.column_config.SelectboxColumn("review_status", options=["Needs review", "Approve", "Ignore"]),
            "save_as": st.column_config.SelectboxColumn("save_as", options=["Purchase", "Sales", "Voucher", "Ignore"]),
            "document_type": st.column_config.SelectboxColumn(
                "document_type",
                options=[
                    "Purchase invoice",
                    "Sales invoice",
                    "Payment voucher",
                    "Receipt voucher",
                    "Journal voucher",
                    "Credit note",
                    "Debit note",
                    "Unknown",
                    "Invoice",
                    "Voucher",
                ],
            ),
        },
        disabled=["file_name", "file_type", "extraction_status", "text_preview"],
        key="document_review_editor",
    )
    if st.button("Save approved extraction rows"):
        purchases, sales = document_review_to_purchase_sales(edited)
        if purchases.empty and sales.empty:
            st.info("Approve at least one row as Purchase or Sales.")
            return
        if not purchases.empty:
            save_csv(concat_record_frames([load_csv(PURCHASE_FILE, PURCHASE_COLUMNS), purchases], PURCHASE_COLUMNS), PURCHASE_FILE)
        if not sales.empty:
            save_csv(concat_record_frames([load_csv(SALES_FILE, SALES_COLUMNS), sales], SALES_COLUMNS), SALES_FILE)
        st.success("Approved extraction rows saved into purchase/sales records.")
        st.rerun()


def render_purchases(purchases: pd.DataFrame, payment_review: pd.DataFrame) -> None:
    st.subheader("Purchases")
    render_purchase_form()
    st.data_editor(purchases, use_container_width=True, hide_index=True, key="purchase_editor")
    if st.button("Save edited purchase records"):
        edited = st.session_state.get("purchase_editor", {}).get("edited_rows")
        if edited:
            updated = purchases.copy()
            for row_index, changes in edited.items():
                for key, value in changes.items():
                    updated.loc[int(row_index), key] = value
            save_csv(updated, PURCHASE_FILE)
            st.success("Purchase records saved.")
            st.rerun()
    st.subheader("Vendor payment matching")
    st.dataframe(payment_review, use_container_width=True, hide_index=True)


def render_sales(sales: pd.DataFrame, receipt_review: pd.DataFrame) -> None:
    st.subheader("Sales")
    render_sales_form()
    st.data_editor(sales, use_container_width=True, hide_index=True, key="sales_editor")
    if st.button("Save edited sales records"):
        edited = st.session_state.get("sales_editor", {}).get("edited_rows")
        if edited:
            updated = sales.copy()
            for row_index, changes in edited.items():
                for key, value in changes.items():
                    updated.loc[int(row_index), key] = value
            save_csv(updated, SALES_FILE)
            st.success("Sales records saved.")
            st.rerun()
    st.subheader("Sales receipt matching")
    st.dataframe(receipt_review, use_container_width=True, hide_index=True)


def render_purchase_form() -> None:
    with st.expander("Add manual purchase"):
        with st.form("manual_purchase", clear_on_submit=True):
            supplier = st.text_input("Supplier name")
            gstin = st.text_input("Supplier GSTIN")
            invoice_no = st.text_input("Invoice no")
            invoice_date = st.date_input("Invoice date")
            taxable = st.number_input("Taxable value", min_value=0.0, step=100.0, key="purchase_taxable")
            gst_type = st.selectbox("GST type", ["CGST+SGST", "IGST"], key="purchase_gst_type")
            gst_rate = st.selectbox("GST rate", [0.0, 5.0, 12.0, 18.0, 28.0], index=3, key="purchase_gst_rate")
            cess = st.number_input("Cess", min_value=0.0, step=10.0, key="purchase_cess")
            tax = calculate_gst(taxable, gst_type, gst_rate, cess)
            st.caption(f"Total {format_money(tax['total'])}")
            submitted = st.form_submit_button("Save purchase")
        if submitted:
            if not supplier or not invoice_no:
                st.error("Supplier and invoice number are required.")
                return
            new_row = manual_purchase_record(supplier, gstin, invoice_no, invoice_date, taxable, gst_type, gst_rate, cess)
            save_csv(concat_record_frames([load_csv(PURCHASE_FILE, PURCHASE_COLUMNS), new_row], PURCHASE_COLUMNS), PURCHASE_FILE)
            st.rerun()


def render_sales_form() -> None:
    with st.expander("Add manual sales"):
        with st.form("manual_sales", clear_on_submit=True):
            customer = st.text_input("Customer name")
            gstin = st.text_input("Customer GSTIN optional")
            invoice_no = st.text_input("Invoice no")
            invoice_date = st.date_input("Invoice date")
            sale_type = st.selectbox("Sale type", ["B2B", "B2C", "Export", "Credit note", "Debit note"])
            taxable = st.number_input("Taxable value", min_value=0.0, step=100.0, key="sales_taxable")
            gst_type = st.selectbox("GST type", ["CGST+SGST", "IGST"], key="sales_gst_type")
            gst_rate = st.selectbox("GST rate", [0.0, 5.0, 12.0, 18.0, 28.0], index=3, key="sales_gst_rate")
            cess = st.number_input("Cess", min_value=0.0, step=10.0, key="sales_cess")
            tax = calculate_gst(taxable, gst_type, gst_rate, cess)
            st.caption(f"Total {format_money(tax['total'])}")
            submitted = st.form_submit_button("Save sales")
        if submitted:
            if not customer or not invoice_no:
                st.error("Customer and invoice number are required.")
                return
            new_row = manual_sales_record(customer, gstin, invoice_no, invoice_date, taxable, gst_type, gst_rate, cess, sale_type)
            save_csv(concat_record_frames([load_csv(SALES_FILE, SALES_COLUMNS), new_row], SALES_COLUMNS), SALES_FILE)
            st.rerun()


def render_bank(bank_entries: pd.DataFrame, sales_review: pd.DataFrame, purchase_review: pd.DataFrame) -> None:
    st.subheader("Bank")
    st.warning("Credits and debits remain review items until categorized. Only Sales receipt and Vendor payment affect exception logic.")
    edited = st.data_editor(
        bank_entries,
        use_container_width=True,
        hide_index=True,
        column_config={"category": st.column_config.SelectboxColumn("category", options=BANK_CATEGORIES)},
        disabled=["entry_id", "date", "description", "reference", "debit", "credit", "balance", "suggested_category"],
        key="bank_editor",
    )
    if st.button("Save bank categories"):
        save_csv(edited, BANK_FILE)
        st.success("Bank entries saved.")
        st.rerun()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Sales receipt review")
        st.dataframe(sales_review, use_container_width=True, hide_index=True)
    with c2:
        st.subheader("Vendor payment review")
        st.dataframe(purchase_review, use_container_width=True, hide_index=True)


def render_exceptions(exceptions: pd.DataFrame) -> None:
    st.subheader("Exceptions")
    if exceptions.empty:
        st.success("No current exceptions.")
        return
    for section in exceptions["section"].dropna().unique():
        st.markdown(f"**{section}**")
        st.dataframe(exceptions[exceptions["section"] == section], use_container_width=True, hide_index=True)


def render_gst_summary(summary: pd.DataFrame) -> None:
    st.subheader("GST Summary")
    st.caption("Possible bank sales are not included in output tax until converted into approved sales records.")
    st.dataframe(summary, use_container_width=True, hide_index=True)


def render_export(
    purchases: pd.DataFrame,
    sales: pd.DataFrame,
    bank_entries: pd.DataFrame,
    document_review: pd.DataFrame,
    exceptions: pd.DataFrame,
    summary: pd.DataFrame,
    validation_issues: pd.DataFrame,
    gstr_2b_reconciliation: pd.DataFrame,
    sales_bank_review: pd.DataFrame,
) -> None:
    st.subheader("Export")
    empty_register = pd.DataFrame()
    st.download_button(
        "Download GST Books Builder workbook",
        data=make_excel_report(
            empty_register,
            validation_issues,
            gstr_2b_reconciliation,
            bank_entries=bank_entries,
            bank_sales_review=sales_bank_review,
            gst_summary_frame=summary,
            document_review=document_review,
            purchase_records=purchases,
            sales_records=sales,
            exceptions=exceptions,
        ),
        file_name="gst_books_builder_review.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def format_money(value: float) -> str:
    return f"Rs {value:,.2f}"


if __name__ == "__main__":
    main()
