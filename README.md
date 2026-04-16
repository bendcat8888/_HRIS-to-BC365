# HRIS Overall Payroll to MS Dynamic BC365 Converter

Converts an HRIS-exported payroll CSV into an Excel file formatted for import to Microsoft Dynamics 365 Business Central (BC365).

Includes:

- Auto mapping for salary affiliate sharing
- Automatic calculation of government-mandated contributions (SSS, Pag-IBIG, PhilHealth)
- Share split rules from `hr_pmr_list.csv` (IPI/SPI/Aldril/Randril/Totalcare) with a built-in fallback rule for Distribution (dept_code=5, sect_code=11)
- PMR code mapping (HRIS employee code -> BC code) from `hr_pmr_list.csv`
- Department code mapping (HRIS dept_code -> BC DEPT Code)
- Auto debit/credit sign handling for payables/deductions (e.g., tax, SSS/PHIC/HDMF, loans, savings)
- Vendor line handling for selected items (e.g., Maxicare and misc deductions) using `hr_vendor_list.csv`
- Auto-fills key BC365 fields (Line No, Recurring Method/Frequency, Posting Date, Document No, Description)
- Simple GUI file picker (Tkinter) and auto-opens the generated Excel output

## Output Columns

- Line No
- Recurring Method
- Recurring Frequency
- External Document No.
- Posting Date
- Document Type
- Document No.
- Account Type
- Account No
- Description
- Amount
- PMR Code
- DEPT Code

## Input CSV (HRIS Export)

The script expects an HRIS payroll export CSV that contains (at minimum) these columns:

```text
first_name, last_name, mid_name, emp_code,
dept_code, dept_name, group_code, sect_code,
basic_pay, w_tax,
sss_ee, sss_er, ecc_er,
med_ee, med_er,
pagibig_ee, pagibig_er,
late_amt, abs_amt, under_amt,
ot_amt, net_pay, other_tax_earn,
loan_savings, savings, loan_pagibig, loan_sss,
loan_housing, bond_tablet, bond_cash
```

If your export is missing any of these, the script may fail with a KeyError during processing.

## Business Rules (Highlights)

- Affiliate sharing: reads per-employee share percentages from `hr_pmr_list.csv` (IPI/SPI/Aldril/Randril/Totalcare) and applies a default split for Distribution (dept_code=5, sect_code=11) when no explicit share is found.
- Government-mandated contributions:
  - SSS Contribution Payable = sss_ee + sss_er + ecc_er
  - PHIC Contribution Payable = med_ee + med_er
  - HDMF Contribution Payable = pagibig_ee + pagibig_er
  - EC - PHIC = PHIC Contribution Payable / 2
  - EC - SSS = sss_er + ecc_er
  - EC - HDMF = 200 when Document No contains "15", else 0
- Vendor lines:
  - MAXICARE / MISC SMART / MISCELLANEOUS / PERSONAL MEDICINES are converted to Vendor lines using `hr_vendor_list.csv` (by employee FullName).
  - Affiliate categories are mapped to Vendor accounts: SPI=AF0019, Aldril=AF0025, Randril=AF0007, Totalcare=AF0020.

## Data Privacy & Security

- The script does not require credentials and does not call external services; it reads local CSV files and writes a local Excel output.
- Payroll exports contain personal and compensation data; avoid committing real HRIS exports and real mapping files to public repositories.

## What It Produces

- An `.xlsx` file named `Converted_To_BC365_Format_<timestamp>.xlsx`
- Output opens automatically after generation

## Requirements

- Windows
- Python 3.9+
- Python packages:
  - pandas
  - openpyxl

Install dependencies:

```powershell
py -3.9 -m pip install pandas openpyxl
```

## Required Supporting Files

Place these CSV files in the same folder as the script (or run the script from the folder that contains them):

- `hr_pmr_list.csv`
  - Required columns: `IHRIS`, `BC`
  - Optional share columns (used for account splitting): `IPI`, `SPI`, `Aldril`, `Randril`, `Totalcare`
- `hr_vendor_list.csv`
  - Required columns: `Name`, `No_`

Do not commit real employee/vendor mapping files to a public repository.

## Suggested .gitignore

```gitignore
__pycache__/
*.pyc
Converted_To_BC365_Format_*.xlsx
hr_pmr_list.csv
hr_vendor_list.csv
```

## Usage

Run the script:

```powershell
py -3.9 "HRIS_to_BC365 V3.py"
```

Workflow:

1. Enter a Document No. (example: `PAYJANUARY`)
2. Click `Browse and Process...` and select the HRIS CSV export
3. The converted Excel file is generated and opened

## Quick Syntax Check

```powershell
py -3.9 -m py_compile "HRIS_to_BC365 V3.py"
```
