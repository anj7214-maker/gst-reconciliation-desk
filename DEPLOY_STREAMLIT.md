# Deploy To Streamlit Community Cloud

This app is ready for Streamlit Community Cloud.

## Files Streamlit Needs

- `app.py` is the app entrypoint.
- `requirements.txt` declares Python dependencies.
- `.streamlit/config.toml` contains basic theme settings.
- `gst_rules.py` contains the GST validation and reconciliation logic.

## Deploy Steps

1. Create a GitHub repository, for example `gst-reconciliation-desk`.
2. Upload all files from this folder to the repository root.
3. Go to `https://share.streamlit.io`.
4. Click **Create app**.
5. Choose the GitHub repository and branch.
6. Set the main file path to:

```text
app.py
```

7. Keep Python as the default unless you need to change it later.
8. Click **Deploy**.

## Important Privacy Note

This first version processes uploaded files during the active Streamlit session. Do not upload sensitive client data to a public demo app unless access is restricted and the data handling policy is clear.

## Recommended App URL

```text
gst-reconciliation-desk.streamlit.app
```
