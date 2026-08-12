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

## Deploying on Streamlit Cloud

1. Push `streamlit_app.py`, `web_app.py`, and `requirements.txt` to the repo, all at the same level.
2. In Streamlit Cloud, set the main file to `streamlit_app.py`.
3. Deploy. If you previously deployed with a different requirements filename, use **Manage app → Reboot** so it reinstalls dependencies from the correct file.

## How it works

1. Upload BDR
2. Upload Case_AR
3. Upload Master
4. Click **Process files**
5. Download `Billing_Submission_Master_File.xlsx`

Under the hood, `streamlit_app.py` imports `web_app.py` purely to call its
`process_one()` function — the Flask/desktop startup code in `web_app.py`
never executes on Streamlit.

## Troubleshooting

**`ModuleNotFoundError: No module named 'flask'`** (or pandas, numpy, etc.)
This means Streamlit Cloud didn't install from your requirements file —
almost always because the file isn't named exactly `requirements.txt`, or
the app needs a reboot to pick up a fresh copy of it. Fix the filename,
commit, and reboot the app from **Manage app**.

**`web_app.py wasn't found next to streamlit_app.py`**
Both files need to live in the same folder in the repo. Check the path
Streamlit Cloud is using as the app root.

**Processing runs but fails partway through**
Expand "Technical details" under the error message in the app — it shows
the full Python traceback from `process_one()`, which is the fastest way to
tell whether the issue is in an uploaded file's format or in the processing
logic itself.

## Keeping the desktop build separate

The EXE build stays independent of this. Use `web_app.py` + PyInstaller for
the Windows desktop version, and `streamlit_app.py` for the web deployment —
same processing engine, two different front ends.
