# GST Reconciliation Desk

A zero-cost local prototype for GST register cleanup, validation, reconciliation, and Excel reporting.

## What It Does

- Upload sales register files as CSV/XLS/XLSX.
- Upload purchase/books files as CSV/XLS/XLSX.
- Upload GSTR-2B or counterparty files as CSV/XLS/XLSX.
- Normalizes common GST column names into one clean format.
- Flags practical GST data issues.
- Reconciles purchase/books data against GSTR-2B-style data.
- Exports a review workbook with clean data, issues, and reconciliation.

## Run

```powershell
python -m streamlit run app.py
```

## Deploy

See `DEPLOY_STREAMLIT.md` for Streamlit Community Cloud deployment steps.

## Expected Columns

The app accepts common names and aliases for:

- GSTIN
- Party / Supplier / Customer / Vendor Name
- Invoice No
- Invoice Date
- Place of Supply
- Taxable Value
- IGST
- CGST
- SGST
- Cess
- Total / Invoice Value

## Deliberately Skipped In This Version

- Login and user roles.
- OCR/PDF extraction.
- Cloud database and hosting.
- GSTN/GSP filing integration.
- Paid AI APIs.

Those can be added after real sample files prove which workflow matters most.
