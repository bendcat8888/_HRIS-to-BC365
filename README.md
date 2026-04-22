# 🏦 HRIS to MS Dynamics BC365 Converter
**Automated Financial Reconciliation & ERP Data Pipeline for Complex Payroll Systems**

[![Python 3.12](https://img.shields.io/badge/Python-3.12-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![MS Dynamics](https://img.shields.io/badge/Dynamics_365-0078D4?style=for-the-badge&logo=microsoft-dynamics-365&logoColor=white)](https://dynamics.microsoft.com/)

## 🛠 Tech Stack

| Category | Tools |
| :--- | :--- |
| **Language** | **Python 3.12** |
| **Data Engine** | **Pandas** (ETL - Extract, Transform, Load) |
| **File Handling** | **OpenPyXL** (Excel XLSX Generation) |
| **GUI** | **Tkinter** (Lightweight Desktop Interface) |
| **Business Logic** | Custom Rules Engine for Affiliate Cost-Sharing & Tax Calculations |

---

## 🎯 Project Overview
This specialized ETL tool automates the bridge between **HRIS Payroll Exports** and **Microsoft Dynamics 365 Business Central (BC365)**. It transforms raw employee compensation data into a structured, validated General Ledger (GL) import file, handling complex affiliate splits and government contribution logic that would otherwise require hours of manual calculation.

### 🌟 High-Value Business Logic
* **Affiliate Cost-Sharing:** Automatically splits salary expenses across multiple legal entities (IPI, SPI, Aldril, Randril, Totalcare) based on per-employee percentage mappings.
* **Government Contribution Engine:** Real-time calculation of **SSS, PhilHealth (PHIC), and Pag-IBIG (HDMF)** employer/employee shares, including specialized logic for EC (Employees' Compensation) contributions.
* **Smart Sign Handling:** Automatically manages Debit/Credit signs for payables vs. deductions to ensure the trial balance remains consistent during ERP import.

---

## 🚀 Key Professional Capabilities

### 📊 Financial Data Transformation
* **Vendor Line Mapping:** Intelligently identifies deductions for specific vendors (e.g., Maxicare, Insurance) and maps them to verified Vendor Account IDs in BC365.
* **Complex Mapping Tables:** Cross-references HRIS employee codes to BC365 PMR codes and Department codes using external lookup tables (`hr_pmr_list.csv`).
* **Validation & Safety:** Includes fallback rules (e.g., Department/Section defaults) to ensure the script never fails on missing or unexpected data points.

### 🛡️ Enterprise Data Engineering
* **Statutory Compliance:** Built-in logic for 15th/30th period variations in HDMF contributions.
* **Automated BC365 Formatting:** Auto-fills required enterprise fields: Line Numbers, Recurring Methods, Posting Dates, and Document Types.
* **Secure Local Execution:** Designed as a "Zero-Cloud" local utility to ensure sensitive payroll and compensation data never leaves the secure internal network.

---

## ⚙️ Development & Quick Start

### Requirements
- **Python 3.9+**
- **Libraries:** `pandas`, `openpyxl`

### Installation
```powershell
pip install pandas openpyxl
```

**Usage Workflow**
1. Initialize: Run the script to launch the Tkinter GUI.
2. Configure: Enter the target Document Number (e.g., PAYROLL_MARCH_2026).
3. Process: Select the HRIS CSV export.
4. Review: The system generates a validated .xlsx file and opens it automatically for final audit.

---

## 📜 License & Intellectual Property
**Copyright (c) 2026 Benedic Cater / InnoGen Pharmaceuticals Inc.**

**All Rights Reserved.**
This repository is published for **portfolio review and technical demonstration purposes only.**

**Strict Restrictions:**
- **No Reproduction:** No part of this code may be copied, modified, or distributed.
- **Brand Protection:** Use of the "InnoGen" or "Solvang" name, branding, or logos is strictly prohibited.
- **Data Privacy:** Use of any proprietary data or business logic contained herein for commercial or personal projects is strictly prohibited.

_For professional inquiries or permission requests, please contact Benedic Cater._

