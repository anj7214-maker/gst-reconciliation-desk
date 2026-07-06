from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable
from xml.etree import ElementTree

import pandas as pd


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
    "match_status",
    "matched_invoice_no",
]

DOCUMENT_COLUMNS = [
    "file_name",
    "file_type",
    "document_type",
    "invoice_type",
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
    else:
        raise ValueError("Upload a CSV, XLS, or XLSX file.")

    return normalize_register(raw, source=source, invoice_type=invoice_type)


def read_bank_statement(uploaded_file) -> ImportResult:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    elif name.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(uploaded_file)
    else:
        raise ValueError("Upload a CSV, XLS, or XLSX bank statement.")

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
        text, warning = "", "Image uploaded for manual review. OCR is not included in the zero-cost version yet."
    else:
        text, warning = "", "Unsupported document format."

    if warning:
        warnings.append(f"{file_name}: {warning}")

    row = extract_document_fields(text, file_name=file_name, file_type=suffix)
    return ImportResult(pd.DataFrame([row], columns=DOCUMENT_COLUMNS), warnings)


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
    normalized["category"] = normalized["suggested_category"].where(
        normalized["suggested_category"].isin({"Ignore", "Bank charges"}),
        "Needs review",
    )
    normalized["match_status"] = "Unmatched"
    normalized["matched_invoice_no"] = ""

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


def make_excel_report(
    normalized: pd.DataFrame,
    issues: pd.DataFrame,
    reconciliation: pd.DataFrame,
    manual_transactions: pd.DataFrame | None = None,
    bank_entries: pd.DataFrame | None = None,
    bank_sales_review: pd.DataFrame | None = None,
    gst_summary_frame: pd.DataFrame | None = None,
    document_review: pd.DataFrame | None = None,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        normalized.to_excel(writer, sheet_name="Clean Register", index=False)
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
        issues.to_excel(writer, sheet_name="Validation Issues", index=False)
        reconciliation.to_excel(writer, sheet_name="Reconciliation", index=False)
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
        return "Possible sales receipt"
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
        "invoice_type": invoice_type,
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
        "text_preview": normalized_text[:1200],
    }


def extract_pdf_text(content: bytes) -> tuple[str, str]:
    try:
        from pypdf import PdfReader
    except ImportError:
        return "", "PDF extraction needs pypdf. It is listed in requirements for Streamlit deployment."

    try:
        reader = PdfReader(BytesIO(content))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        return "", f"Could not read PDF text: {exc}"

    if not text.strip():
        return "", "No selectable PDF text found. Use image/manual review now; add OCR later."
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


def bank_entry_id(row: pd.Series) -> str:
    parts = [
        str(row.get("date") or ""),
        str(row.get("description") or ""),
        str(row.get("reference") or ""),
        f"{money(row.get('debit')):.2f}",
        f"{money(row.get('credit')):.2f}",
    ]
    return str(abs(hash("|".join(parts))))


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
