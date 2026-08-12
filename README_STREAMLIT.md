# Billing Submission Automation — Streamlit

This is the web version of the tool. It runs on Streamlit Cloud and does **not**
open a desktop window, start pywebview, or build an EXE — that's what the
separate `web_app.py` + PyInstaller path is for.

## Files needed in the repo (all in the same folder)

```text
streamlit_app.py
web_app.py
requirements.txt
```

**Important:** the file must be named exactly `requirements.txt`. Streamlit
Cloud only auto-installs from that exact filename — `requirements_streamlit.txt`
or any other name will be ignored, and your app will fail on the first import
that needs one of those packages.

## requirements.txt

```txt
streamlit
flask
pandas
numpy
openpyxl
python-calamine
```

Flask is listed here even though the Flask server itself never runs on
Streamlit. That's because `streamlit_app.py` loads `web_app.py` as a module
to reuse its `process_one()` function, and `web_app.py` imports Flask at the
top of the file — so Flask still needs to be installed for that import to
succeed, even though `app.run()` is never called.

## How it works

1. Upload BDR
2. Upload Case_AR
3. Upload Master
4. Click **Process files**
5. Download `Billing_Submission_Master_File.xlsx`
