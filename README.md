# Billing Submission Automation

> **Automate your daily Billing Submission workflow — without manually changing file sources.**

## 📌 What is this application?

**Billing Submission Automation** is a desktop/web utility designed to simplify the daily billing submission preparation process.

Instead of manually opening multiple workbooks, changing sources, copying data, and applying the same processing logic every day, the application lets you provide the three required source files:

1. **BDR (Billing Detail Report)**
2. **Case_AR**
3. **Master**

The application processes these files using the established billing-submission logic and creates a new:

**`Billing_Submission_Master_File.xlsx`**

You can then open the generated workbook in Excel and manually refresh your Pivot/Data Model as part of the existing workflow.

---

## 🚀 How the workflow works

```text
        ┌──────────────┐
        │   BDR File   │
        └──────┬───────┘
               │
               │
        ┌──────▼───────┐
        │ Billing      │
        │ Submission   │
        │ Automation   │
        └──────▲───────┘
               │
        ┌──────┴───────┐
        │              │
┌───────┴───────┐ ┌────┴────────┐
│   Case_AR     │ │    Master   │
└───────────────┘ └─────────────┘
               │
               ▼
┌────────────────────────────────┐
│ Billing_Submission_Master_File │
│             .xlsx              │
└────────────────────────────────┘
               │
               ▼
       Refresh Pivot/Data Model
```

---

## 📂 Required input files

### 1. BDR

Upload the current **Billing Detail Report**.

The application reads the BDR data and uses the relevant Master headers/fields according to the processing logic.

### 2. Case_AR

Upload the current **Case_AR** Excel file.

This file is used for the applicable Case ID / EDI Status lookup logic.

### 3. Master

Upload the current **Billing Submission Master** workbook.

The Master workbook provides the destination structure and existing workbook/template.

---

## ⚙️ What the application does

The processing workflow includes the established billing-submission logic, including:

- Reading BDR data
- Reading Case_AR data
- Reading the Master workbook
- Matching/populating applicable BDR data into the Master structure
- Case_AR lookup logic for EDI Status
- Submission Date handling
- Billing Month calculation
- Billing Week calculation
- Verification Status / Bill Type / EDI Service Type / EDI Status cascade logic
- Existing status conditions and business rules
- Creation of a separate output workbook

### Important

The application does **not** require the provider name to be part of the filename.

You simply select the files you want to process.

---

## 🖥️ User Interface

The application is designed around a simple three-file workflow:

### Step 1 — Upload BDR

Click:

**`Upload BDR`**

and select the current BDR file.

### Step 2 — Upload Case_AR

Click:

**`Upload Case_AR`**

and select the current Case_AR workbook.

### Step 3 — Upload Master

Click:

**`Upload Master`**

and select the Master workbook.

### Step 4 — Process

Click:

**`PROCESS FILES`**

The application validates the selected files and starts processing.

### Step 5 — Output

The application generates:

```text
Billing_Submission_Master_File.xlsx
```

The original input files are not intended to be overwritten.

---

## 📊 Output

The generated workbook is the prepared **Billing Submission Master File**.

After generation, continue with the normal Excel workflow, including:

```text
Open Billing_Submission_Master_File.xlsx
              ↓
Review output
              ↓
Refresh Pivot / Data Model
              ↓
Continue normal reporting workflow
```

---

## 🔐 Safety / validation

The application validates input files before processing.

Examples of validation include:

- File exists
- File can be read
- BDR is not empty
- Required workbook inputs are available
- Output does not overwrite an input file

If a file cannot be processed, the application should stop rather than silently generating an incomplete billing file.

---

## 🛠️ Running from Python

### Requirements

Python 3.x with the packages listed in:

```text
requirements.txt
```

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

Run the application:

```bash
python web_app.py
```

---

## 🪟 Windows EXE

The application can be packaged as a standalone Windows executable using PyInstaller.

Typical build command:

```bash
python -m PyInstaller --clean --noconfirm --onefile --windowed --name Billing_Submission_Automation web_app.py
```

After a successful build:

```text
dist/
└── Billing_Submission_Automation.exe
```

The EXE is intended to allow normal users to run the application without opening Jupyter.

---

## 🧩 Project dependencies

Main Python dependencies:

- Flask
- pandas
- NumPy
- openpyxl
- python-calamine
- PyInstaller (for EXE builds)

See `requirements.txt` for the dependency list.

---

## ❗ Troubleshooting

### `ModuleNotFoundError`

Example:

```text
ModuleNotFoundError: No module named 'flask'
```

Run:

```bash
python -m pip install -r requirements.txt
```

### PyInstaller not recognized

Run:

```bash
python -m pip install pyinstaller
```

Then build using:

```bash
python -m PyInstaller --clean --noconfirm --onefile --windowed --name Billing_Submission_Automation web_app.py
```

### BDR is empty

If the application reports that the BDR file is empty or contains no readable data, verify that the selected BDR file is the actual exported BDR and contains its headers/data.

---

## 🎯 Designed for

This application is intended for teams that repeatedly perform the same Billing Submission preparation process and want to reduce:

- Manual source changes
- Repetitive copying/pasting
- File-selection mistakes
- Repetitive formula/lookup work
- Manual processing steps

The goal is simple:

> **Select the three files → Process → Download/receive the prepared Master File.**

---

## 📋 Current workflow philosophy

The application separates the workflow into two parts:

### Automation

```text
BDR
Case_AR
Master
   ↓
Automated processing
   ↓
Billing_Submission_Master_File.xlsx
```

### Existing Excel workflow

```text
Generated Master
   ↓
Manual review
   ↓
Manual Pivot/Data Model refresh
```

This keeps the automated portion focused on the established billing-submission logic while allowing the existing Excel reporting process to remain unchanged.

---

## 📞 Support / Maintenance

When modifying the application, preserve the established business rules and test the generated Master file against a known-good output before deploying a new version.

**Version:** 1.0  
**Application:** Billing Submission Automation
