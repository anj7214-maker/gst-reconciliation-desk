# GST Books Builder

A zero-cost Streamlit prototype for turning bills, vouchers, registers, and bank statements into reviewed GST purchase and sales records.

## What It Does

- Upload purchase registers, sales registers, GSTR-2B files, bank statements, and bill/voucher documents.
- Extract selectable text from PDF/DOCX/TXT bills and vouchers.
- Accept JPG/PNG files for manual review and preview.
- Review extracted rows before saving them as purchase or sales records.
- Keep purchase records and sales records separate.
- Manually add purchase and sales records.
- Categorize bank entries without treating every credit as sales.
- Match sales invoices to bank credits categorized as `Sales receipt`.
- Match purchase bills to bank debits categorized as `Vendor payment`.
- Highlight possible unrecorded sales and missing purchase bills in Exceptions.
- Export a GST-ready review workbook.

## Accounting Rule

The app does not calculate sales blindly from bank credits.

- Confirmed sales = approved sales records.
- Possible bank sales = bank credits manually categorized as `Sales receipt`.
- Possible unrecorded sales = `Sales receipt` bank credits not matched to approved sales records.
- Possible bank sales are not included in output tax until converted into approved sales records.

## Run

```powershell
python -m streamlit run app.py
```

## Deploy

See `DEPLOY_STREAMLIT.md` for Streamlit Community Cloud deployment steps.

## Persistence

Manual purchase records, sales records, and reviewed bank categories are saved locally under `data/`. On Streamlit Community Cloud this storage is runtime-local, so use exports as the durable handoff for now.

## Current Limits

- OCR is not included yet for scanned PDFs or image-only bills.
- Legacy `.doc` files need conversion to `.docx`.
- No login, billing, direct GST filing, or GSTN/GSP integration.
- No paid AI APIs.
