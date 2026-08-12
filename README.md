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

---

## 🖥️ UI

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



## 🧩 Project dependencies

Main Python dependencies:

- Flask
- pandas
- NumPy
- openpyxl
- python-calamine
- PyInstaller (for EXE builds)

See `requirements.txt` for the dependency list.

