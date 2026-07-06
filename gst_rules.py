from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable
from xml.etree import ElementTree

import pandas as pd
from PIL import Image


CANONICAL_COLUMNS = [
    "source",
    "invoice_type",
    "gstin",
    "party_name",
    "invoice_no",
    "invoice_date",
    "place_of_supply",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
    "cess",
    "total",
]

BANK_COLUMNS = [
    "entry_id",
    "date",
    "description",
    "reference",
    "debit",
    "credit",
    "balance",
    "suggested_category",
    "category",
    "matched_purchase_id",
    "matched_sales_id",
    "match_status",
    "review_status",
]

PURCHASE_COLUMNS = [
    "purchase_id",
    "source",
    "supplier_name",
    "supplier_gstin",
    "invoice_no",
    "invoice_date",
    "place_of_supply",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
    "cess",
    "total",
    "itc_eligible",
    "rcm_applicable",
    "gstr_2b_match_status",
    "bank_payment_status",
    "review_status",
]

SALES_COLUMNS = [
    "sales_id",
    "source",
    "customer_name",
    "customer_gstin",
    "invoice_no",
    "invoice_date",
    "sale_type",
    "place_of_supply",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
    "cess",
    "total",
    "payment_status",
    "bank_match_status",
    "review_status",
]

DOCUMENT_COLUMNS = [
    "file_name",
    "file_type",
    "document_type",
    "save_as",
    "gstin",
    "party_name",
    "invoice_no",
    "invoice_date",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
    "cess",
    "total",
    "extraction_status",
    "review_status",
    "notes",
    "text_preview",
]

COLUMN_ALIASES = {
    "gstin": {
        "gstin",
        "supplier gstin",
        "recipient gstin",
        "customer gstin",
        "vendor gstin",
        "ctin",
    },
    "party_name": {
        "party",
        "party name",
        "supplier",
        "supplier name",
        "customer",
        "customer name",
        "vendor",
        "vendor name",
        "name",
    },
    "invoice_no": {
        "invoice no",
        "invoice number",
        "invoice",
        "inv no",
        "inv number",
        "bill no",
        "document no",
        "doc no",
    },
    "invoice_date": {
        "invoice date",
        "date",
        "inv date",
        "bill date",
        "document date",
        "doc date",
    },
    "place_of_supply": {
        "place of supply",
        "pos",
        "state",
        "supply state",
    },
    "taxable_value": {
        "taxable value",
        "taxable",
        "taxable amount",
        "assessable value",
        "base amount",
    },
    "igst": {"igst", "igst amount", "integrated tax"},
    "cgst": {"cgst", "cgst amount", "central tax"},
    "sgst": {"sgst", "sgst amount", "utgst", "sgst/utgst", "state tax"},
    "cess": {"cess", "cess amount"},
    "total": {"total", "invoice value", "invoice amount", "gross total", "grand total", "amount"},
}

BANK_COLUMN_ALIASES = {
    "date": {"date", "transaction date", "txn date", "value date", "posting date"},
    "description": {
        "description",
        "narration",
        "particulars",
        "transaction remarks",
        "remarks",
        "details",
    },
    "reference": {"reference", "ref", "ref no", "cheque no", "utr", "transaction id", "chq/ref no"},
    "debit": {"debit", "withdrawal", "withdrawals", "paid out", "dr", "debit amount"},
    "credit": {"credit", "deposit", "deposits", "paid in", "cr", "credit amount"},
    "balance": {"balance", "closing balance", "running balance"},
}

GSTIN_RE = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
GSTIN_FIND_RE = re.compile(r"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]")
DATE_RE = re.compile(r"\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b")
AMOUNT_RE = re.compile(r"(?i)(?:rs\.?|inr|â‚¹)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)")


@dataclass(frozen=True)
class ImportResult:
    frame: pd.DataFrame
    warnings: list[str]


def read_register(uploaded_file, source: str, invoice_type: str) -> ImportResult:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    elif name.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(uploaded_file)
    elif name.endswith(".pdf"):
        result = read_document_file(uploaded_file)
        result.frame["save_as"] = invoice_type
        return document_review_to_register_import(result.frame, source=source, invoice_type=invoice_type, warnings=result.warnings)
    else:
        raise ValueError("Upload a CSV, XLS, XLSX, or PDF file.")

    return normalize_register(raw, source=source, invoice_type=invoice_type)


def read_bank_statement(uploaded_file) -> ImportResult:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    elif name.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(uploaded_file)
    elif name.endswith(".pdf"):
        return read_bank_pdf(uploaded_file)
    else:
        raise ValueError("Upload a CSV, XLS, XLSX, or PDF bank statement.")

    return normalize_bank_statement(raw)


def read_document_file(uploaded_file) -> ImportResult:
    file_name = uploaded_file.name
    suffix = file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""
    content = uploaded_file.getvalue()
    warnings: list[str] = []

    if suffix == "pdf":
        text, warning = extract_pdf_text(content)
    elif suffix in {"docx", "doc"}:
        text, warning = extract_docx_text(content) if suffix == "docx" else ("", "Legacy .doc files need conversion to .docx first.")
    elif suffix == "txt":
        text, warning = decode_text(content), ""
    elif suffix in {"jpg", "jpeg", "png"}:
        text, warning = extract_image_text(content)
    else:
        text, warning = "", "Unsupported document format."

    if warning:
        warnings.append(f"{file_name}: {warning}")

    row = extract_document_fields(text, file_name=file_name, file_type=suffix)
    return ImportResult(pd.DataFrame([row], columns=DOCUMENT_COLUMNS), warnings)


def read_bank_pdf(uploaded_file) -> ImportResult:
    text, warning = extract_pdf_text(uploaded_file.getvalue())
    bank_frame = bank_entries_from_text(text)
    warnings = [f"{uploaded_file.name}: {warning}"] if warning else []
    if bank_frame.empty:
        warnings.append(f"{uploaded_file.name}: Could not detect bank rows from PDF text/OCR. Use CSV/XLSX or manual review.")
    return ImportResult(bank_frame, warnings)


def document_review_to_register_import(review: pd.DataFrame, source: str, invoice_type: str, warnings: list[str]) -> ImportResult:
    if review.empty:
        return ImportResult(pd.DataFrame(columns=CANONICAL_COLUMNS), warnings)

    frame = pd.DataFrame(
        {
            "source": source,
            "invoice_type": invoice_type,
            "gstin": review["gstin"],
            "party_name": review["party_name"],
            "invoice_no": review["invoice_no"],
            "invoice_date": pd.to_datetime(review["invoice_date"], errors="coerce").dt.date,
            "place_of_supply": "",
            "taxable_value": pd.to_numeric(review["taxable_value"], errors="coerce").fillna(0.0),
            "igst": pd.to_numeric(review["igst"], errors="coerce").fillna(0.0),
            "cgst": pd.to_numeric(review["cgst"], errors="coerce").fillna(0.0),
            "sgst": pd.to_numeric(review["sgst"], errors="coerce").fillna(0.0),
            "cess": pd.to_numeric(review["cess"], errors="coerce").fillna(0.0),
            "total": pd.to_numeric(review["total"], errors="coerce").fillna(0.0),
        }
    )
    if not warnings:
        warnings = [f"{source}: PDF was converted into one extracted review row; verify fields before using."]
    return ImportResult(frame[CANONICAL_COLUMNS], warnings)


def normalize_register(raw: pd.DataFrame, source: str, invoice_type: str) -> ImportResult:
    warnings: list[str] = []
    normalized = pd.DataFrame()
    raw_columns = {clean_header(col): col for col in raw.columns}

    for canonical, aliases in COLUMN_ALIASES.items():
        source_column = next((raw_columns[a] for a in aliases if a in raw_columns), None)
        if source_column is None:
            normalized[canonical] = pd.NA
            if canonical in {"gstin", "invoice_no", "invoice_date", "taxable_value", "total"}:
                warnings.append(f"Missing expected column: {canonical}")
        else:
            normalized[canonical] = raw[source_column]

    normalized.insert(0, "invoice_type", invoice_type)
    normalized.insert(0, "source", source)

    normalized["gstin"] = normalized["gstin"].map(clean_gstin)
    normalized["invoice_no"] = normalized["invoice_no"].map(clean_invoice_no)
    normalized["invoice_date"] = pd.to_datetime(normalized["invoice_date"], errors="coerce").dt.date

    for column in ["taxable_value", "igst", "cgst", "sgst", "cess", "total"]:
        normalized[column] = pd.to_numeric(normalized[column], errors="coerce").fillna(0.0).round(2)

    for column in CANONICAL_COLUMNS:
        if column not in normalized:
            normalized[column] = pd.NA

    return ImportResult(normalized[CANONICAL_COLUMNS], warnings)


def normalize_bank_statement(raw: pd.DataFrame) -> ImportResult:
    warnings: list[str] = []
    normalized = pd.DataFrame()
    raw_columns = {clean_header(col): col for col in raw.columns}

    for canonical, aliases in BANK_COLUMN_ALIASES.items():
        source_column = next((raw_columns[a] for a in aliases if a in raw_columns), None)
        if source_column is None:
            normalized[canonical] = pd.NA
            if canonical in {"date", "description", "debit", "credit"}:
                warnings.append(f"Missing expected bank column: {canonical}")
        else:
            normalized[canonical] = raw[source_column]

    normalized["date"] = pd.to_datetime(normalized["date"], errors="coerce").dt.date
    for column in ["debit", "credit", "balance"]:
        normalized[column] = pd.to_numeric(normalized[column], errors="coerce").fillna(0.0).round(2)

    normalized["description"] = normalized["description"].fillna("").astype(str)
    normalized["reference"] = normalized["reference"].fillna("").astype(str)
    normalized["entry_id"] = normalized.apply(bank_entry_id, axis=1)
    normalized["suggested_category"] = normalized.apply(classify_bank_entry, axis=1)
    normalized["category"] = "Needs review"
    normalized["matched_purchase_id"] = ""
    normalized["matched_sales_id"] = ""
    normalized["match_status"] = "Unmatched"
    normalized["review_status"] = "Needs review"

    for column in BANK_COLUMNS:
        if column not in normalized:
            normalized[column] = pd.NA

    return ImportResult(normalized[BANK_COLUMNS], warnings)


def calculate_gst(taxable_value: object, gst_type: str, gst_rate: object, cess: object = 0) -> dict[str, float]:
    taxable = money(taxable_value)
    rate = money(gst_rate)
    cess_amount = money(cess)
    gst_amount = round(taxable * rate / 100, 2)

    if gst_type == "IGST":
        igst, cgst, sgst = gst_amount, 0.0, 0.0
    else:
        igst, cgst, sgst = 0.0, round(gst_amount / 2, 2), round(gst_amount / 2, 2)

    total = round(taxable + igst + cgst + sgst + cess_amount, 2)
    return {"igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess_amount, "total": total}


def manual_transaction(
    invoice_type: str,
    party_name: str,
    gstin: str,
    invoice_no: str,
    invoice_date,
    taxable_value: object,
    gst_type: str,
    gst_rate: object,
    cess: object,
    notes: str = "",
    payment_received: bool | None = None,
    bank_reference: str = "",
) -> pd.DataFrame:
    tax = calculate_gst(taxable_value, gst_type, gst_rate, cess)
    source = f"Manual {invoice_type}"
    row = {
        "source": source,
        "invoice_type": invoice_type,
        "gstin": clean_gstin(gstin),
        "party_name": party_name.strip(),
        "invoice_no": clean_invoice_no(invoice_no),
        "invoice_date": invoice_date,
        "place_of_supply": "",
        "taxable_value": money(taxable_value),
        "igst": tax["igst"],
        "cgst": tax["cgst"],
        "sgst": tax["sgst"],
        "cess": tax["cess"],
        "total": tax["total"],
        "gst_type": gst_type,
        "gst_rate": money(gst_rate),
        "notes": notes.strip(),
        "payment_received": payment_received,
        "bank_reference": bank_reference.strip(),
    }
    return pd.DataFrame([row])


def manual_purchase_record(
    supplier_name: str,
    supplier_gstin: str,
    invoice_no: str,
    invoice_date,
    taxable_value: object,
    gst_type: str,
    gst_rate: object,
    cess: object,
) -> pd.DataFrame:
    tax = calculate_gst(taxable_value, gst_type, gst_rate, cess)
    row = {
        "purchase_id": stable_id("purchase", supplier_gstin, invoice_no, invoice_date, tax["total"]),
        "source": "Manual purchase",
        "supplier_name": supplier_name.strip(),
        "supplier_gstin": clean_gstin(supplier_gstin),
        "invoice_no": clean_invoice_no(invoice_no),
        "invoice_date": invoice_date,
        "place_of_supply": "",
        "taxable_value": money(taxable_value),
        "igst": tax["igst"],
        "cgst": tax["cgst"],
        "sgst": tax["sgst"],
        "cess": tax["cess"],
        "total": tax["total"],
        "itc_eligible": True,
        "rcm_applicable": False,
        "gstr_2b_match_status": "Not checked",
        "bank_payment_status": "Unmatched",
        "review_status": "Approved",
    }
    return pd.DataFrame([row], columns=PURCHASE_COLUMNS)


def manual_sales_record(
    customer_name: str,
    customer_gstin: str,
    invoice_no: str,
    invoice_date,
    taxable_value: object,
    gst_type: str,
    gst_rate: object,
    cess: object,
    sale_type: str,
) -> pd.DataFrame:
    tax = calculate_gst(taxable_value, gst_type, gst_rate, cess)
    row = {
        "sales_id": stable_id("sales", customer_gstin, invoice_no, invoice_date, tax["total"]),
        "source": "Manual sales",
        "customer_name": customer_name.strip(),
        "customer_gstin": clean_gstin(customer_gstin),
        "invoice_no": clean_invoice_no(invoice_no),
        "invoice_date": invoice_date,
        "sale_type": sale_type,
        "place_of_supply": "",
        "taxable_value": money(taxable_value),
        "igst": tax["igst"],
        "cgst": tax["cgst"],
        "sgst": tax["sgst"],
        "cess": tax["cess"],
        "total": tax["total"],
        "payment_status": "Unmatched",
        "bank_match_status": "Unmatched",
        "review_status": "Approved",
    }
    return pd.DataFrame([row], columns=SALES_COLUMNS)


def register_to_purchase_records(frame: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in frame.iterrows():
        total = money(row.get("total"))
        rows.append(
            {
                "purchase_id": stable_id("purchase", row.get("gstin"), row.get("invoice_no"), row.get("invoice_date"), total),
                "source": row.get("source", ""),
                "supplier_name": row.get("party_name", ""),
                "supplier_gstin": clean_gstin(row.get("gstin", "")),
                "invoice_no": clean_invoice_no(row.get("invoice_no", "")),
                "invoice_date": row.get("invoice_date", ""),
                "place_of_supply": row.get("place_of_supply", ""),
                "taxable_value": money(row.get("taxable_value")),
                "igst": money(row.get("igst")),
                "cgst": money(row.get("cgst")),
                "sgst": money(row.get("sgst")),
                "cess": money(row.get("cess")),
                "total": total,
                "itc_eligible": True,
                "rcm_applicable": False,
                "gstr_2b_match_status": "Not checked",
                "bank_payment_status": "Unmatched",
                "review_status": "Approved",
            }
        )
    return pd.DataFrame(rows, columns=PURCHASE_COLUMNS)


def register_to_sales_records(frame: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in frame.iterrows():
        total = money(row.get("total"))
        gstin = clean_gstin(row.get("gstin", ""))
        rows.append(
            {
                "sales_id": stable_id("sales", gstin, row.get("invoice_no"), row.get("invoice_date"), total),
                "source": row.get("source", ""),
                "customer_name": row.get("party_name", ""),
                "customer_gstin": gstin,
                "invoice_no": clean_invoice_no(row.get("invoice_no", "")),
                "invoice_date": row.get("invoice_date", ""),
                "sale_type": "B2B" if gstin else "B2C",
                "place_of_supply": row.get("place_of_supply", ""),
                "taxable_value": money(row.get("taxable_value")),
                "igst": money(row.get("igst")),
                "cgst": money(row.get("cgst")),
                "sgst": money(row.get("sgst")),
                "cess": money(row.get("cess")),
                "total": total,
                "payment_status": "Unmatched",
                "bank_match_status": "Unmatched",
                "review_status": "Approved",
            }
        )
    return pd.DataFrame(rows, columns=SALES_COLUMNS)


def document_review_to_purchase_sales(review: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if review.empty:
        return pd.DataFrame(columns=PURCHASE_COLUMNS), pd.DataFrame(columns=SALES_COLUMNS)

    approved = review[review["review_status"] == "Approve"].copy()
    purchases: list[dict] = []
    sales: list[dict] = []
    for _, row in approved.iterrows():
        save_as = row.get("save_as", row.get("invoice_type", "Purchase"))
        total = money(row.get("total"))
        common = {
            "source": "Document: " + str(row.get("file_name", "")),
            "invoice_no": clean_invoice_no(row.get("invoice_no", "")),
            "invoice_date": row.get("invoice_date", ""),
            "place_of_supply": "",
            "taxable_value": money(row.get("taxable_value")),
            "igst": money(row.get("igst")),
            "cgst": money(row.get("cgst")),
            "sgst": money(row.get("sgst")),
            "cess": money(row.get("cess")),
            "total": total,
            "review_status": "Approved",
        }
        if save_as == "Sales":
            gstin = clean_gstin(row.get("gstin", ""))
            sales.append(
                {
                    "sales_id": stable_id("sales", gstin, common["invoice_no"], common["invoice_date"], total),
                    **common,
                    "customer_name": row.get("party_name", ""),
                    "customer_gstin": gstin,
                    "sale_type": "B2B" if gstin else "B2C",
                    "payment_status": "Unmatched",
                    "bank_match_status": "Unmatched",
                }
            )
        elif save_as == "Purchase":
            gstin = clean_gstin(row.get("gstin", ""))
            purchases.append(
                {
                    "purchase_id": stable_id("purchase", gstin, common["invoice_no"], common["invoice_date"], total),
                    **common,
                    "supplier_name": row.get("party_name", ""),
                    "supplier_gstin": gstin,
                    "itc_eligible": True,
                    "rcm_applicable": False,
                    "gstr_2b_match_status": "Not checked",
                    "bank_payment_status": "Unmatched",
                }
            )

    return pd.DataFrame(purchases, columns=PURCHASE_COLUMNS), pd.DataFrame(sales, columns=SALES_COLUMNS)


def document_review_to_transactions(review: pd.DataFrame) -> pd.DataFrame:
    if review.empty:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)

    approved = review[review["review_status"] == "Approve"].copy()
    if approved.empty:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)

    approved["source"] = "Document: " + approved["file_name"].fillna("").astype(str)
    approved["place_of_supply"] = ""
    for column in ["taxable_value", "igst", "cgst", "sgst", "cess", "total"]:
        approved[column] = pd.to_numeric(approved[column], errors="coerce").fillna(0.0).round(2)
    approved["gstin"] = approved["gstin"].map(clean_gstin)
    approved["invoice_no"] = approved["invoice_no"].map(clean_invoice_no)
    approved["invoice_date"] = pd.to_datetime(approved["invoice_date"], errors="coerce").dt.date
    return approved[CANONICAL_COLUMNS + ["file_name", "document_type", "text_preview"]]


def purchase_records_to_register(purchases: pd.DataFrame) -> pd.DataFrame:
    if purchases.empty:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)
    frame = pd.DataFrame(
        {
            "source": purchases["source"],
            "invoice_type": "Purchase",
            "gstin": purchases["supplier_gstin"],
            "party_name": purchases["supplier_name"],
            "invoice_no": purchases["invoice_no"],
            "invoice_date": purchases["invoice_date"],
            "place_of_supply": purchases["place_of_supply"],
            "taxable_value": purchases["taxable_value"],
            "igst": purchases["igst"],
            "cgst": purchases["cgst"],
            "sgst": purchases["sgst"],
            "cess": purchases["cess"],
            "total": purchases["total"],
        }
    )
    return frame[CANONICAL_COLUMNS]


def sales_records_to_register(sales: pd.DataFrame) -> pd.DataFrame:
    if sales.empty:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)
    frame = pd.DataFrame(
        {
            "source": sales["source"],
            "invoice_type": "Sales",
            "gstin": sales["customer_gstin"],
            "party_name": sales["customer_name"],
            "invoice_no": sales["invoice_no"],
            "invoice_date": sales["invoice_date"],
            "place_of_supply": sales["place_of_supply"],
            "taxable_value": sales["taxable_value"],
            "igst": sales["igst"],
            "cgst": sales["cgst"],
            "sgst": sales["sgst"],
            "cess": sales["cess"],
            "total": sales["total"],
        }
    )
    return frame[CANONICAL_COLUMNS]


def validate_register(frame: pd.DataFrame) -> pd.DataFrame:
    issues: list[dict] = []

    if frame.empty:
        return pd.DataFrame(columns=["severity", "source", "invoice_no", "gstin", "issue", "suggestion"])

    duplicate_mask = frame.duplicated(subset=["source", "gstin", "invoice_no"], keep=False)

    for index, row in frame.iterrows():
        row_issues = row_validation_issues(row)
        if duplicate_mask.loc[index]:
            row_issues.append(("High", "Duplicate invoice within same source", "Check whether this invoice was imported twice."))

        for severity, issue, suggestion in row_issues:
            issues.append(
                {
                    "severity": severity,
                    "source": row.get("source", ""),
                    "invoice_no": row.get("invoice_no", ""),
                    "gstin": row.get("gstin", ""),
                    "issue": issue,
                    "suggestion": suggestion,
                }
            )

    return pd.DataFrame(issues, columns=["severity", "source", "invoice_no", "gstin", "issue", "suggestion"])


def row_validation_issues(row: pd.Series) -> list[tuple[str, str, str]]:
    issues: list[tuple[str, str, str]] = []
    gstin = str(row.get("gstin") or "").strip()
    invoice_no = str(row.get("invoice_no") or "").strip()
    taxable = money(row.get("taxable_value"))
    igst = money(row.get("igst"))
    cgst = money(row.get("cgst"))
    sgst = money(row.get("sgst"))
    cess = money(row.get("cess"))
    total = money(row.get("total"))
    tax_total = igst + cgst + sgst + cess

    if not gstin:
        issues.append(("High", "Missing GSTIN", "Add supplier/customer GSTIN before relying on this record."))
    elif not GSTIN_RE.match(gstin):
        issues.append(("High", "Invalid GSTIN format", "Verify the 15-character GSTIN and state code."))

    if not invoice_no:
        issues.append(("High", "Missing invoice number", "Add the invoice number for matching and audit trail."))

    if pd.isna(row.get("invoice_date")):
        issues.append(("Medium", "Missing or invalid invoice date", "Use a valid date from the source invoice/register."))

    if taxable < 0 or total < 0:
        issues.append(("Medium", "Negative amount detected", "Confirm whether this should be a credit note/debit note."))

    if abs((taxable + tax_total) - total) > 2:
        issues.append(("High", "Taxable plus tax does not match total", "Review taxable value, tax columns, cess, and total."))

    if igst > 0 and (cgst > 0 or sgst > 0):
        issues.append(("High", "IGST mixed with CGST/SGST", "Inter-state supplies normally use IGST; intra-state supplies use CGST and SGST."))

    if abs(cgst - sgst) > 1 and (cgst > 0 or sgst > 0):
        issues.append(("Medium", "CGST and SGST are unequal", "For intra-state taxable supply, CGST and SGST are usually equal."))

    # syed-minimal: rule-based checks cover the first 0-cost version; add rate/HSN intelligence after real sample files prove the gap.
    if taxable > 0 and tax_total == 0:
        issues.append(("Low", "No GST amount on taxable value", "Confirm exempt/nil-rated/non-GST treatment."))

    return issues


def reconcile_registers(books: pd.DataFrame, counterparty: pd.DataFrame) -> pd.DataFrame:
    if books.empty and counterparty.empty:
        return pd.DataFrame(
            columns=[
                "status",
                "gstin",
                "invoice_no",
                "book_total",
                "counterparty_total",
                "difference",
                "book_source",
                "counterparty_source",
            ]
        )

    left = aggregate_for_reconciliation(books, "book")
    right = aggregate_for_reconciliation(counterparty, "counterparty")
    merged = left.merge(right, on=["gstin", "invoice_no"], how="outer")

    for column in ["book_total", "counterparty_total"]:
        merged[column] = merged[column].fillna(0.0)

    merged["difference"] = (merged["book_total"] - merged["counterparty_total"]).round(2)
    merged["status"] = merged.apply(reconciliation_status, axis=1)

    columns = [
        "status",
        "gstin",
        "invoice_no",
        "book_total",
        "counterparty_total",
        "difference",
        "book_source",
        "counterparty_source",
    ]
    return merged[columns].sort_values(["status", "gstin", "invoice_no"], na_position="last")


def aggregate_for_reconciliation(frame: pd.DataFrame, prefix: str) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["gstin", "invoice_no", f"{prefix}_total", f"{prefix}_source"])

    grouped = (
        frame.groupby(["gstin", "invoice_no"], dropna=False)
        .agg(total=("total", "sum"), source=("source", lambda values: ", ".join(sorted(set(map(str, values))))))
        .reset_index()
    )
    grouped[f"{prefix}_total"] = grouped.pop("total").round(2)
    grouped[f"{prefix}_source"] = grouped.pop("source")
    return grouped


def reconciliation_status(row: pd.Series) -> str:
    book_total = money(row.get("book_total"))
    counterparty_total = money(row.get("counterparty_total"))
    if book_total == 0:
        return "Only in counterparty file"
    if counterparty_total == 0:
        return "Only in books"
    if abs(book_total - counterparty_total) <= 2:
        return "Matched"
    return "Amount mismatch"


def summary_metrics(frame: pd.DataFrame) -> dict[str, float]:
    if frame.empty:
        return {"Invoices": 0, "Taxable Value": 0.0, "Output/Input Tax": 0.0, "Invoice Total": 0.0}

    return {
        "Invoices": float(len(frame)),
        "Taxable Value": money(frame["taxable_value"].sum()),
        "Output/Input Tax": money(frame[["igst", "cgst", "sgst", "cess"]].sum().sum()),
        "Invoice Total": money(frame["total"].sum()),
    }


def record_tax_total(frame: pd.DataFrame) -> float:
    if frame.empty:
        return 0.0
    return money(frame[["igst", "cgst", "sgst", "cess"]].sum().sum())


def match_sales_receipts(sales: pd.DataFrame, bank_entries: pd.DataFrame, tolerance: float = 2.0) -> pd.DataFrame:
    bank_sales = bank_entries[bank_entries.get("category", "") == "Sales receipt"].copy()
    if bank_sales.empty:
        return empty_bank_review()

    sales_rows = sales.copy()
    if not sales_rows.empty:
        sales_rows["invoice_date"] = pd.to_datetime(sales_rows["invoice_date"], errors="coerce")

    reviews: list[dict] = []
    for _, bank_row in bank_sales.iterrows():
        credit = money(bank_row.get("credit"))
        bank_date = pd.to_datetime(bank_row.get("date"), errors="coerce")
        candidates = pd.DataFrame()
        if not sales_rows.empty:
            candidates = sales_rows[(sales_rows["total"].sub(credit).abs() <= tolerance)].copy()
            if not candidates.empty and not pd.isna(bank_date):
                candidates["date_gap"] = (candidates["invoice_date"] - bank_date).abs().dt.days
                close_candidates = candidates[candidates["date_gap"] <= 7]
                if not close_candidates.empty:
                    candidates = close_candidates.sort_values(["date_gap", "invoice_no"])

        if candidates.empty:
            status, invoice_no, party_name = "Unmatched", "", ""
        else:
            match = candidates.iloc[0]
            status = "Matched"
            invoice_no = match.get("invoice_no", "")
            party_name = match.get("party_name", "")

        reviews.append(
            {
                "entry_id": bank_row.get("entry_id", ""),
                "date": bank_row.get("date", ""),
                "description": bank_row.get("description", ""),
                "credit": credit,
                "category": bank_row.get("category", ""),
                "match_status": status,
                "matched_invoice_no": invoice_no,
                "matched_party": party_name,
            }
        )

    return pd.DataFrame(reviews)


def match_vendor_payments(purchases: pd.DataFrame, bank_entries: pd.DataFrame, tolerance: float = 2.0) -> pd.DataFrame:
    vendor_payments = bank_entries[bank_entries.get("category", "") == "Vendor payment"].copy()
    if vendor_payments.empty:
        return pd.DataFrame(columns=["entry_id", "date", "description", "debit", "match_status", "matched_invoice_no", "matched_party"])

    purchase_rows = purchases.copy()
    if not purchase_rows.empty:
        purchase_rows["invoice_date"] = pd.to_datetime(purchase_rows["invoice_date"], errors="coerce")

    reviews: list[dict] = []
    for _, bank_row in vendor_payments.iterrows():
        debit = money(bank_row.get("debit"))
        bank_date = pd.to_datetime(bank_row.get("date"), errors="coerce")
        candidates = pd.DataFrame()
        if not purchase_rows.empty:
            candidates = purchase_rows[(purchase_rows["total"].sub(debit).abs() <= tolerance)].copy()
            if not candidates.empty and not pd.isna(bank_date):
                candidates["date_gap"] = (candidates["invoice_date"] - bank_date).abs().dt.days
                close_candidates = candidates[candidates["date_gap"] <= 30]
                if not close_candidates.empty:
                    candidates = close_candidates.sort_values(["date_gap", "invoice_no"])

        if candidates.empty:
            status, invoice_no, party_name = "Unmatched", "", ""
        else:
            match = candidates.iloc[0]
            status = "Matched"
            invoice_no = match.get("invoice_no", "")
            party_name = match.get("supplier_name", "")

        reviews.append(
            {
                "entry_id": bank_row.get("entry_id", ""),
                "date": bank_row.get("date", ""),
                "description": bank_row.get("description", ""),
                "debit": debit,
                "match_status": status,
                "matched_invoice_no": invoice_no,
                "matched_party": party_name,
            }
        )

    return pd.DataFrame(reviews)


def gst_summary(sales: pd.DataFrame, purchases: pd.DataFrame, bank_review: pd.DataFrame, issues: pd.DataFrame) -> pd.DataFrame:
    sales_summary = summary_metrics(sales)
    purchase_summary = summary_metrics(purchases)
    possible_bank_sales = money(bank_review[bank_review["category"] == "Sales receipt"]["credit"].sum()) if not bank_review.empty else 0.0
    unmatched_bank_sales = (
        int((bank_review["match_status"] == "Unmatched").sum()) if not bank_review.empty and "match_status" in bank_review else 0
    )
    output_tax = sales_summary["Output/Input Tax"]
    input_tax = purchase_summary["Output/Input Tax"]
    rows = [
        ("Approved sales invoice total", sales_summary["Invoice Total"]),
        ("Possible bank sales receipts", possible_bank_sales),
        ("Approved purchase total", purchase_summary["Invoice Total"]),
        ("Output tax", output_tax),
        ("Input tax credit", input_tax),
        ("Net GST payable", output_tax - input_tax),
        ("Validation issues", float(len(issues))),
        ("Unmatched bank sales receipts", float(unmatched_bank_sales)),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def gst_books_summary(
    purchases: pd.DataFrame,
    sales: pd.DataFrame,
    bank_entries: pd.DataFrame,
    sales_bank_review: pd.DataFrame,
    exceptions: pd.DataFrame,
    issues: pd.DataFrame,
) -> pd.DataFrame:
    approved_sales_total = money(sales[sales["review_status"] == "Approved"]["total"].sum()) if not sales.empty else 0.0
    approved_purchase_total = (
        money(purchases[purchases["review_status"] == "Approved"]["total"].sum()) if not purchases.empty else 0.0
    )
    output_tax = record_tax_total(sales[sales["review_status"] == "Approved"]) if not sales.empty else 0.0
    input_tax = record_tax_total(purchases[purchases["review_status"] == "Approved"]) if not purchases.empty else 0.0
    possible_bank_sales = money(bank_entries[bank_entries["category"] == "Sales receipt"]["credit"].sum()) if not bank_entries.empty else 0.0
    possible_unrecorded = (
        money(sales_bank_review[sales_bank_review["match_status"] == "Unmatched"]["credit"].sum())
        if not sales_bank_review.empty
        else 0.0
    )
    needs_review = count_needs_review(purchases) + count_needs_review(sales) + count_needs_review(bank_entries)
    rows = [
        ("Confirmed sales from approved sales records", approved_sales_total),
        ("Possible bank sales marked Sales receipt", possible_bank_sales),
        ("Possible unrecorded sales", possible_unrecorded),
        ("Approved purchases", approved_purchase_total),
        ("Output tax from sales records", output_tax),
        ("Input tax credit from purchase records", input_tax),
        ("Net GST payable", output_tax - input_tax),
        ("Validation issues", float(len(issues))),
        ("Records needing review", float(needs_review)),
        ("Exception records", float(len(exceptions))),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def build_exceptions(
    purchases: pd.DataFrame,
    sales: pd.DataFrame,
    bank_entries: pd.DataFrame,
    sales_bank_review: pd.DataFrame,
    purchase_bank_review: pd.DataFrame,
    gstr_2b_reconciliation: pd.DataFrame,
    validation_issues: pd.DataFrame,
) -> pd.DataFrame:
    rows: list[dict] = []

    for _, row in sales_bank_review[sales_bank_review.get("match_status", "") == "Unmatched"].iterrows():
        rows.append(exception_row("Possible unrecorded sales", "High", row.get("description", ""), row.get("credit", 0), "Create or match a sales record."))

    matched_sales_invoices = set(sales_bank_review.get("matched_invoice_no", pd.Series(dtype=str)).dropna().astype(str))
    for _, row in sales.iterrows():
        invoice_no = str(row.get("invoice_no", ""))
        if invoice_no and invoice_no not in matched_sales_invoices:
            rows.append(exception_row("Sales invoice not matched to bank credit", "Medium", invoice_no, row.get("total", 0), "Check payment receipt or mark as unpaid."))

    for _, row in purchase_bank_review[purchase_bank_review.get("match_status", "") == "Unmatched"].iterrows():
        rows.append(exception_row("Vendor payment without purchase bill", "High", row.get("description", ""), row.get("debit", 0), "Create or match a purchase record."))

    matched_purchase_invoices = set(purchase_bank_review.get("matched_invoice_no", pd.Series(dtype=str)).dropna().astype(str))
    for _, row in purchases.iterrows():
        invoice_no = str(row.get("invoice_no", ""))
        if invoice_no and invoice_no not in matched_purchase_invoices:
            rows.append(exception_row("Purchase bill not matched to bank debit", "Medium", invoice_no, row.get("total", 0), "Check vendor payment status."))

    if not gstr_2b_reconciliation.empty:
        missing = gstr_2b_reconciliation[gstr_2b_reconciliation["status"] == "Only in books"]
        for _, row in missing.iterrows():
            rows.append(exception_row("Purchase bill missing in GSTR-2B", "High", row.get("invoice_no", ""), row.get("book_total", 0), "Review ITC eligibility before claiming."))

    for _, row in validation_issues.iterrows():
        issue = str(row.get("issue", ""))
        kind = "Duplicate invoices" if "Duplicate" in issue else "Tax calculation errors" if "Taxable plus tax" in issue or "GST" in issue else "Validation issue"
        rows.append(exception_row(kind, row.get("severity", "Medium"), row.get("invoice_no", ""), 0, row.get("suggestion", "")))

    for frame_name, frame, id_column in [
        ("Purchase record needs review", purchases, "invoice_no"),
        ("Sales record needs review", sales, "invoice_no"),
        ("Bank entry needs review", bank_entries, "description"),
    ]:
        if not frame.empty and "review_status" in frame:
            for _, row in frame[frame["review_status"].fillna("") != "Approved"].iterrows():
                rows.append(exception_row("Records needing review", "Medium", row.get(id_column, ""), row.get("total", row.get("credit", row.get("debit", 0))), frame_name))

    return pd.DataFrame(rows, columns=["section", "severity", "reference", "amount", "suggested_action"])


def exception_row(section: str, severity: str, reference: object, amount: object, suggested_action: object) -> dict:
    return {
        "section": section,
        "severity": severity,
        "reference": reference,
        "amount": money(amount),
        "suggested_action": suggested_action,
    }


def make_excel_report(
    normalized: pd.DataFrame,
    issues: pd.DataFrame,
    reconciliation: pd.DataFrame,
    manual_transactions: pd.DataFrame | None = None,
    bank_entries: pd.DataFrame | None = None,
    bank_sales_review: pd.DataFrame | None = None,
    gst_summary_frame: pd.DataFrame | None = None,
    document_review: pd.DataFrame | None = None,
    purchase_records: pd.DataFrame | None = None,
    sales_records: pd.DataFrame | None = None,
    exceptions: pd.DataFrame | None = None,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        (purchase_records if purchase_records is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Purchase Records",
            index=False,
        )
        (sales_records if sales_records is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Sales Records",
            index=False,
        )
        (manual_transactions if manual_transactions is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Manual Transactions",
            index=False,
        )
        (bank_entries if bank_entries is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Bank Entries",
            index=False,
        )
        (bank_sales_review if bank_sales_review is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Bank Sales Review",
            index=False,
        )
        (document_review if document_review is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Document Review",
            index=False,
        )
        (exceptions if exceptions is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="Exceptions",
            index=False,
        )
        issues.to_excel(writer, sheet_name="Validation Issues", index=False)
        reconciliation.to_excel(writer, sheet_name="GSTR-2B Reconciliation", index=False)
        (gst_summary_frame if gst_summary_frame is not None else pd.DataFrame()).to_excel(
            writer,
            sheet_name="GST Summary",
            index=False,
        )
    output.seek(0)
    return output.getvalue()


def classify_bank_entry(row: pd.Series) -> str:
    description = str(row.get("description") or "").lower()
    debit = money(row.get("debit"))
    credit = money(row.get("credit"))
    ignore_keywords = ["self", "transfer", "upi transfer", "neft transfer", "imps transfer"]
    charge_keywords = ["charge", "charges", "fee", "fees"]
    gst_keywords = ["gst", "goods and services tax"]
    loan_keywords = ["loan"]
    salary_keywords = ["salary", "wages"]
    interest_keywords = ["interest"]

    if any(word in description for word in charge_keywords):
        return "Bank charges"
    if any(word in description for word in gst_keywords):
        return "GST payment"
    if any(word in description for word in loan_keywords):
        return "Loan"
    if any(word in description for word in salary_keywords):
        return "Needs review"
    if any(word in description for word in ignore_keywords):
        return "Ignore"
    if any(word in description for word in interest_keywords):
        return "Needs review"
    if credit > 0:
        return "Possible receipt"
    if debit > 0:
        return "Possible payment/expense"
    return "Needs review"


def extract_document_fields(text: str, file_name: str, file_type: str) -> dict:
    normalized_text = normalize_text(text)
    document_type = infer_document_type(normalized_text, file_name)
    invoice_type = infer_invoice_type(normalized_text, document_type)
    gstins = [match.group(0) for match in GSTIN_FIND_RE.finditer(normalized_text.replace(" ", "").upper())]

    taxable = labeled_amount(normalized_text, ["taxable value", "taxable amount", "assessable value", "sub total"])
    igst = labeled_amount(normalized_text, ["igst", "integrated tax"])
    cgst = labeled_amount(normalized_text, ["cgst", "central tax"])
    sgst = labeled_amount(normalized_text, ["sgst", "utgst", "state tax"])
    cess = labeled_amount(normalized_text, ["cess"])
    total = labeled_amount(normalized_text, ["grand total", "invoice value", "invoice total", "total amount", "total"])

    if taxable == 0 and total > 0:
        taxable = max(round(total - igst - cgst - sgst - cess, 2), 0.0)

    return {
        "file_name": file_name,
        "file_type": file_type,
        "document_type": document_type,
        "save_as": invoice_type,
        "gstin": gstins[0] if gstins else "",
        "party_name": "",
        "invoice_no": labeled_text(normalized_text, ["invoice no", "invoice number", "voucher no", "bill no", "receipt no"]),
        "invoice_date": first_date(normalized_text),
        "taxable_value": taxable,
        "igst": igst,
        "cgst": cgst,
        "sgst": sgst,
        "cess": cess,
        "total": total,
        "extraction_status": "Text extracted" if normalized_text else "Manual review needed",
        "review_status": "Needs review",
        "notes": "",
        "text_preview": normalized_text[:1200],
    }


def extract_pdf_text(content: bytes) -> tuple[str, str]:
    warnings: list[str] = []
    try:
        from pypdf import PdfReader
    except ImportError:
        warnings.append("PDF text extraction needs pypdf. It is listed in requirements for Streamlit deployment.")
        text = ""
    else:
        try:
            reader = PdfReader(BytesIO(content))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as exc:
            warnings.append(f"Could not read selectable PDF text: {exc}")
            text = ""

    if text.strip():
        return text, ""

    ocr_text, ocr_warning = extract_pdf_ocr_text(content)
    if ocr_text.strip():
        warning = "Selectable PDF text not found; OCR text was used."
        if warnings:
            warning = " ".join(warnings + [warning])
        return ocr_text, warning

    warnings.append(ocr_warning or "No selectable or OCR-readable PDF text found.")
    return "", " ".join(warnings)


def extract_pdf_ocr_text(content: bytes) -> tuple[str, str]:
    try:
        from pdf2image import convert_from_bytes
    except ImportError:
        return "", "PDF OCR needs pdf2image. It is listed in requirements for Streamlit deployment."

    try:
        pages = convert_from_bytes(content, dpi=200, first_page=1, last_page=3)
    except Exception as exc:
        return "", f"PDF OCR needs Poppler/pdf rendering support: {exc}"

    text_parts: list[str] = []
    warnings: list[str] = []
    for index, page in enumerate(pages, start=1):
        text, warning = ocr_image(page)
        if warning:
            warnings.append(f"page {index}: {warning}")
        text_parts.append(text)
    return "\n".join(text_parts), "; ".join(warnings)


def extract_image_text(content: bytes) -> tuple[str, str]:
    try:
        image = Image.open(BytesIO(content))
    except Exception as exc:
        return "", f"Could not read image for OCR: {exc}"
    text, warning = ocr_image(image)
    if text.strip():
        return text, "Image OCR text was used; verify all fields."
    return "", warning or "Image OCR returned no text. Review manually."


def ocr_image(image: Image.Image) -> tuple[str, str]:
    try:
        import pytesseract
    except ImportError:
        return "", "OCR needs pytesseract. It is listed in requirements for Streamlit deployment."

    try:
        text = pytesseract.image_to_string(image)
    except Exception as exc:
        return "", f"OCR needs the Tesseract binary installed: {exc}"

    return text, ""


def extract_docx_text(content: bytes) -> tuple[str, str]:
    try:
        with zipfile.ZipFile(BytesIO(content)) as docx_zip:
            xml = docx_zip.read("word/document.xml")
    except Exception as exc:
        return "", f"Could not read DOCX text: {exc}"

    root = ElementTree.fromstring(xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    text = "\n".join(node.text or "" for node in root.findall(".//w:t", namespace))
    return text, "" if text.strip() else "No readable text found in DOCX."


def decode_text(content: bytes) -> str:
    for encoding in ["utf-8", "utf-16", "latin-1"]:
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return ""


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def infer_document_type(text: str, file_name: str) -> str:
    lower = f"{file_name} {text}".lower()
    if "credit note" in lower:
        return "Credit note"
    if "debit note" in lower:
        return "Debit note"
    if "payment voucher" in lower:
        return "Payment voucher"
    if "receipt voucher" in lower or "receipt" in lower:
        return "Receipt voucher"
    if "journal voucher" in lower:
        return "Journal voucher"
    if "voucher" in lower:
        return "Voucher"
    if "tax invoice" in lower or "invoice" in lower or "bill" in lower:
        return "Invoice"
    return "Unknown"


def infer_invoice_type(text: str, document_type: str) -> str:
    lower = text.lower()
    if document_type in {"Payment voucher", "Journal voucher", "Voucher"}:
        return "Purchase"
    if "sales" in lower or "customer" in lower:
        return "Sales"
    return "Purchase"


def labeled_text(text: str, labels: list[str]) -> str:
    for label in labels:
        pattern = rf"(?i){re.escape(label)}\s*[:#-]?\s*([A-Z0-9./_-]+)"
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()
    return ""


def labeled_amount(text: str, labels: list[str]) -> float:
    for label in labels:
        pattern = rf"(?i){re.escape(label)}\s*[:#-]?\s*(?:rs\.?|inr|â‚¹)?\s*([0-9][0-9,]*(?:\.\d{{1,2}})?)"
        match = re.search(pattern, text)
        if match:
            return money(match.group(1).replace(",", ""))
    return 0.0


def first_date(text: str) -> str:
    match = DATE_RE.search(text)
    return match.group(1) if match else ""


def bank_entries_from_text(text: str) -> pd.DataFrame:
    rows: list[dict] = []
    for line in (text or "").splitlines():
        normalized = normalize_text(line)
        if not normalized:
            continue
        date = first_date(normalized)
        amounts = [money(match.group(1).replace(",", "")) for match in AMOUNT_RE.finditer(normalized)]
        if not date or not amounts:
            continue
        amount = amounts[-1]
        lower = normalized.lower()
        debit = amount if any(token in lower for token in [" debit ", " dr ", "withdrawal", "paid"]) else 0.0
        credit = amount if debit == 0.0 else 0.0
        rows.append(
            {
                "date": date,
                "description": normalized,
                "reference": "",
                "debit": debit,
                "credit": credit,
                "balance": 0.0,
            }
        )
    if not rows:
        return pd.DataFrame(columns=BANK_COLUMNS)
    return normalize_bank_statement(pd.DataFrame(rows)).frame


def bank_entry_id(row: pd.Series) -> str:
    parts = [
        str(row.get("date") or ""),
        str(row.get("description") or ""),
        str(row.get("reference") or ""),
        f"{money(row.get('debit')):.2f}",
        f"{money(row.get('credit')):.2f}",
    ]
    return str(abs(hash("|".join(parts))))


def stable_id(*parts: object) -> str:
    return str(abs(hash("|".join(str(part) for part in parts))))


def count_needs_review(frame: pd.DataFrame) -> int:
    if frame.empty or "review_status" not in frame:
        return 0
    return int((frame["review_status"].fillna("") != "Approved").sum())


def empty_bank_review() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "entry_id",
            "date",
            "description",
            "credit",
            "category",
            "match_status",
            "matched_invoice_no",
            "matched_party",
        ]
    )


def clean_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())


def clean_gstin(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).upper()


def clean_invoice_no(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def money(value: object) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return round(float(value), 2)
    except (TypeError, ValueError):
        return 0.0


def concat_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    usable = [frame for frame in frames if frame is not None and not frame.empty]
    if not usable:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)
    return pd.concat(usable, ignore_index=True)
