# Billing Submission Automation — Streamlit

## Deploy on Streamlit

This version is designed for Streamlit Cloud/web deployment.

It does **not** start Flask, pywebview, a desktop window, or create an EXE.

### Files required in the GitHub repository

```text
streamlit_app.py
web_app.py
requirements.txt
```

### `requirements.txt`

```txt
streamlit
flask
pandas
numpy
openpyxl
python-calamine
```

### Streamlit Cloud

Set the main file to:

```text
streamlit_app.py
```

Then deploy.

### Workflow

1. Upload BDR
2. Upload Case_AR
3. Upload Master
4. Click **PROCESS FILES**
5. Download `Billing_Submission_Master_File.xlsx`

The Streamlit interface reuses the existing `process_one()` processing engine from `web_app.py`. The desktop/Flask startup code is not used by the Streamlit page.

## Important

Keep the existing `web_app.py` in the same repository/folder as `streamlit_app.py`.

The EXE build can remain separate. Use `web_app.py` + PyInstaller for the Windows desktop version, and `streamlit_app.py` for Streamlit deployment.
