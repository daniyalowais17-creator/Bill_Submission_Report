# 📊 Billing Submission Automation

> **Make your daily Billing Submission process faster, simpler, and more consistent.**

## What is Billing Submission Automation?

If you prepare a Billing Submission report every day, you probably know how repetitive the process can become — opening multiple files, updating data, checking statuses, matching information, and preparing the final Master file.

**Billing Submission Automation** is built to simplify that entire process.

Instead of manually changing file sources or repeating the same steps every day, you simply upload your three required files:

1. **BDR (Billing Detail Report)**
2. **Case_AR**
3. **Master**

The application processes the files using the existing Billing Submission business logic and generates a ready-to-use:

**`Billing_Submission_Master_File.xlsx`**

Your normal Excel workflow can then continue from there, including reviewing the output and refreshing the Pivot/Data Model.

---

# 📂 Required Files

Before starting, make sure you have these three files ready.

### 1. BDR — Billing Detail Report

Upload the latest BDR file for the current reporting cycle.

The application reads the BDR data and uses the relevant fields to populate the Billing Submission Master according to the existing processing logic.

### 2. Case_AR

Upload the latest **Case_AR** Excel file.

This file is used for the applicable **Case ID / EDI Status lookup** during processing.

### 3. Master

Upload your current **Billing Submission Master** workbook.

The Master file provides the structure and template that the application uses to create the final output.

> **You don't need to manually change file sources inside the application. Simply upload the current files each time you run the process.**

---

# ⚙️ What the Application Does

Once the three files are uploaded, the application handles the repetitive processing steps for you.

The workflow includes:

- Reading the BDR data
- Reading the Case_AR data
- Reading the Master workbook
- Matching the applicable BDR data with the Master structure
- Using Case_AR for the required EDI Status lookups
- Handling Submission Date information
- Calculating Billing Month
- Calculating Billing Week
- Applying Verification Status logic
- Applying Bill Type logic
- Applying EDI Service Type logic
- Applying EDI Status logic
- Applying the existing business rules and status conditions
- Creating a new Billing Submission Master workbook

The goal is simple:

> **Upload the files once, let the application do the repetitive work, and use the generated Master file for the next step of your normal workflow.**

---

# 🖥️ How to Use the Application

The application is designed to keep the process as simple as possible.

## Step 1 — Upload BDR

Select:

**`Upload BDR`**

and choose the latest Billing Detail Report.

---

## Step 2 — Upload Case_AR

Select:

**`Upload Case_AR`**

and choose the latest Case_AR workbook.

---

## Step 3 — Upload Master

Select:

**`Upload Master`**

and choose the current Billing Submission Master workbook.

---

## Step 4 — Process the Files

Once all three files have been uploaded, select:

**`PROCESS FILES`**

The application will validate the files and begin processing.

You can monitor the processing status directly from the application.

---

## Step 5 — Download the Result

Once processing is complete, the application generates:

```text
Billing_Submission_Master_File.xlsx

## 📊 Final Output

Upload BDR
     +
Upload Case_AR
     +
Upload Master
     ↓
Billing Submission Automation
     ↓
Billing_Submission_Master_File.xlsx
     ↓
Review the output
     ↓
Refresh Pivot / Data Model
     ↓
Continue normal reporting workflow
