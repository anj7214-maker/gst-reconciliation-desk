import pandas as pd

from gst_rules import normalize_register, reconcile_registers, validate_register


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
