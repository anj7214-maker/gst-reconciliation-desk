import pandas as pd

from gst_rules import (
    calculate_gst,
    gst_summary,
    match_sales_receipts,
    normalize_bank_statement,
    normalize_register,
    reconcile_registers,
    validate_register,
)


def test_normalize_and_validate_flags_bad_gstin_and_tax_mismatch():
    raw = pd.DataFrame(
        [
            {
                "GSTIN": "bad",
                "Invoice No": "A1",
                "Invoice Date": "2026-04-01",
                "Taxable Value": 100,
                "CGST": 9,
                "SGST": 0,
                "Total": 109,
            }
        ]
    )

    result = normalize_register(raw, source="sample", invoice_type="Purchase")
    issues = validate_register(result.frame)

    assert "Invalid GSTIN format" in issues["issue"].tolist()
    assert "CGST and SGST are unequal" in issues["issue"].tolist()


def test_reconcile_marks_matched_missing_and_mismatch():
    books = pd.DataFrame(
        [
            {"gstin": "07ABCDE1234F1Z5", "invoice_no": "A1", "total": 118},
            {"gstin": "27ABCDE1234F1Z7", "invoice_no": "A2", "total": 200},
        ]
    )
    books["source"] = "books"
    counterparty = pd.DataFrame(
        [
            {"gstin": "07ABCDE1234F1Z5", "invoice_no": "A1", "total": 118},
            {"gstin": "27ABCDE1234F1Z7", "invoice_no": "A2", "total": 205},
            {"gstin": "29ABCDE1234F1Z1", "invoice_no": "A3", "total": 300},
        ]
    )
    counterparty["source"] = "2b"

    result = reconcile_registers(books, counterparty)

    assert set(result["status"]) == {"Matched", "Amount mismatch", "Only in counterparty file"}


def test_calculate_gst_splits_cgst_sgst():
    tax = calculate_gst(1000, "CGST+SGST", 18, 0)

    assert tax == {"igst": 0.0, "cgst": 90.0, "sgst": 90.0, "cess": 0.0, "total": 1180.0}


def test_bank_normalization_and_classification():
    raw = pd.DataFrame(
        [
            {"Date": "2026-04-01", "Narration": "UPI customer receipt", "Deposit": 1180, "Withdrawal": 0},
            {"Date": "2026-04-02", "Narration": "Bank charges", "Deposit": 0, "Withdrawal": 25},
        ]
    )

    result = normalize_bank_statement(raw)

    assert result.frame.loc[0, "suggested_category"] == "Possible sales receipt"
    assert result.frame.loc[1, "suggested_category"] == "Bank charges"


def test_sales_receipt_matching_by_amount_and_date():
    sales = pd.DataFrame(
        [
            {
                "invoice_no": "S-1",
                "party_name": "Customer",
                "invoice_date": "2026-04-01",
                "total": 1180,
            }
        ]
    )
    bank = pd.DataFrame(
        [
            {
                "entry_id": "1",
                "date": "2026-04-04",
                "description": "customer",
                "credit": 1181,
                "category": "Sales receipt",
            }
        ]
    )

    result = match_sales_receipts(sales, bank)

    assert result.loc[0, "match_status"] == "Matched"
    assert result.loc[0, "matched_invoice_no"] == "S-1"


def test_gst_summary_uses_sales_purchase_tax_and_bank_review():
    sales = pd.DataFrame([{"taxable_value": 1000, "igst": 0, "cgst": 90, "sgst": 90, "cess": 0, "total": 1180}])
    purchases = pd.DataFrame([{"taxable_value": 500, "igst": 0, "cgst": 45, "sgst": 45, "cess": 0, "total": 590}])
    bank_review = pd.DataFrame(
        [{"category": "Sales receipt", "credit": 1180, "match_status": "Matched"}]
    )
    issues = pd.DataFrame(columns=["issue"])

    result = gst_summary(sales, purchases, bank_review, issues)
    lookup = dict(zip(result["metric"], result["value"], strict=False))

    assert lookup["Output tax"] == 180
    assert lookup["Input tax credit"] == 90
    assert lookup["Net GST payable"] == 90
