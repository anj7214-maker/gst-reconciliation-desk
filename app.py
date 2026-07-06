from __future__ import annotations

import pandas as pd
import streamlit as st

from gst_rules import (
    CANONICAL_COLUMNS,
    concat_frames,
    make_excel_report,
    read_register,
    reconcile_registers,
    summary_metrics,
    validate_register,
)


st.set_page_config(page_title="GST Reconciliation Desk", page_icon="GST", layout="wide")


def main() -> None:
    st.title("GST Reconciliation Desk")
    st.caption("Local-first GST register cleanup, validation, and reconciliation.")

    with st.sidebar:
        st.header("Upload")
        sales_files = st.file_uploader(
            "Sales register files",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
        )
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

    sales, sales_warnings = load_files(sales_files, source_prefix="Sales", invoice_type="Sales")
    purchases, purchase_warnings = load_files(purchase_files, source_prefix="Books", invoice_type="Purchase")
    counterparty, counterparty_warnings = load_files(
        counterparty_files,
        source_prefix="GSTR-2B",
        invoice_type="Purchase",
    )
    all_registers = concat_frames([sales, purchases, counterparty])

    warnings = sales_warnings + purchase_warnings + counterparty_warnings
    if warnings:
        with st.expander("Import warnings", expanded=True):
            for warning in warnings:
                st.warning(warning)

    if all_registers.empty:
        render_empty_state()
        return

    issues = validate_register(all_registers)
    reconciliation = reconcile_registers(purchases, counterparty)

    render_dashboard(sales, purchases, counterparty, issues, reconciliation)
    render_tabs(all_registers, issues, reconciliation)


def load_files(files, source_prefix: str, invoice_type: str) -> tuple[pd.DataFrame, list[str]]:
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


def render_empty_state() -> None:
    st.info("Upload at least one sales, purchase, or GSTR-2B CSV/Excel file to start.")
    st.subheader("Expected columns")
    st.write(
        "The app accepts common aliases such as Invoice No, Invoice Date, GSTIN, Taxable Value, "
        "IGST, CGST, SGST, Cess, Total, Supplier Name, and Customer Name."
    )
    st.dataframe(pd.DataFrame(columns=CANONICAL_COLUMNS), use_container_width=True)


def render_dashboard(
    sales: pd.DataFrame,
    purchases: pd.DataFrame,
    counterparty: pd.DataFrame,
    issues: pd.DataFrame,
    reconciliation: pd.DataFrame,
) -> None:
    sales_summary = summary_metrics(sales)
    purchase_summary = summary_metrics(purchases)
    counterparty_summary = summary_metrics(counterparty)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sales total", format_money(sales_summary["Invoice Total"]))
    c2.metric("Purchase total", format_money(purchase_summary["Invoice Total"]))
    c3.metric("GSTR-2B total", format_money(counterparty_summary["Invoice Total"]))
    c4.metric("Validation issues", str(len(issues)))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Output tax", format_money(sales_summary["Output/Input Tax"]))
    c6.metric("Book ITC", format_money(purchase_summary["Output/Input Tax"]))
    c7.metric("2B ITC", format_money(counterparty_summary["Output/Input Tax"]))
    c8.metric("Unresolved recon", str((reconciliation["status"] != "Matched").sum() if not reconciliation.empty else 0))


def render_tabs(all_registers: pd.DataFrame, issues: pd.DataFrame, reconciliation: pd.DataFrame) -> None:
    tabs = st.tabs(["Clean Register", "Validation Issues", "Reconciliation", "Export"])

    with tabs[0]:
        st.dataframe(all_registers, use_container_width=True, hide_index=True)

    with tabs[1]:
        if issues.empty:
            st.success("No validation issues found from the current rule set.")
        else:
            severity = st.multiselect(
                "Severity",
                options=sorted(issues["severity"].unique()),
                default=sorted(issues["severity"].unique()),
            )
            filtered = issues[issues["severity"].isin(severity)]
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[2]:
        if reconciliation.empty:
            st.info("Upload both purchase/books and GSTR-2B/counterparty files to reconcile.")
        else:
            status = st.multiselect(
                "Status",
                options=sorted(reconciliation["status"].unique()),
                default=sorted(reconciliation["status"].unique()),
            )
            filtered = reconciliation[reconciliation["status"].isin(status)]
            st.dataframe(filtered, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.download_button(
            "Download GST review workbook",
            data=make_excel_report(all_registers, issues, reconciliation),
            file_name="gst_review_workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def format_money(value: float) -> str:
    return f"Rs {value:,.2f}"


if __name__ == "__main__":
    main()
