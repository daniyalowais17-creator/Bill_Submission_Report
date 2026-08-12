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

