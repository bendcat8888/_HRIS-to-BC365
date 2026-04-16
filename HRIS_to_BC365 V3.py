# HRIS raw data to BC365 format
# Last Update: Jan 27, 2022 ~ DONE ~
# UPDATE : Jan 28, 2022 - Sort by Name
# UPDATE : Jan 31, 2022 - Negative and Positive per Account No.
# UPDATE: Feb 18, 2022 - Add mapping for PMR Code and Dept CODE
# UPDATE: Feb 18, 2022 - Add Loans mapping, Loan SSS, Loan pagibig, Loan Savings, savings
# Used by Sir Pochie for encoding at BC365s
# UPDATE: Mar 4, 2022 added Savings : -1 in amount
# UPDATE: Mar 10, 2022 added EC computation and mapping
# UPDATE: April 5, 2022 added 2 items in Dept_Code (TradeDev and Solvang)
# UPDATE: April 5, 2022 added MAXICARE, MISC SMART, MISCELLANEOUS, PERSONAL MEDICINES
# UPDATE: June 23, 2022 added list in pmr_dict, account_dict, amount_dict, DEPT_dict
# UPDATE: June 23, 2022 added condition if DocNo have "15" EC-HDMF = 100 else 0

# import pyodbc
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import filedialog as fd
import os
import sys
import csv 

print("\n+++++++++++++++++++++++++++++++++++++++++++")
n=input("\nPlease input your Document No. [ ex. PAYJANUARY ] : ")
print("+++++++++++++++++++++++++++++++++++++++++++")

# Default shares for Distribution condition: dept_name="Distribution", dept_code=5, sect_code=11
DEFAULT_SHARE_DISTRIBUTION = {'IPI': 0.80, 'SPI': 0.10, 'Aldril': 0.065, 'Randril': 0.035, 'Totalcare': 0}

def share_distribution(df2a, Share_dict, Share_Account_dict):
    """
    Apply share split for Share_Account_dict accounts.
    Two sources: (1) Share_dict from hr_pmr_list.csv, (2) Distribution condition: dept_name="Distribution", dept_code=5, sect_code=11
    """
    new_rows = []
    for _, row in df2a.iterrows():
        if row['Account Name'] not in Share_Account_dict:
            new_row = row.copy()
            new_row['category'] = ''
            new_rows.append(new_row)
            continue

        # Get shares: from Share_dict (IHRIS in hr_pmr_list) or from Distribution condition
        shares = None
        if row['IHRIS'] in Share_dict:
            shares = Share_dict[row['IHRIS']]
        else:
            dept_name = str(row.get('dept_name', '')).strip()
            dept_code = row.get('_share_dept_code')
            sect_code = row.get('_share_sect_code')
            if dept_name == "Distribution" and dept_code is not None and sect_code is not None:
                try:
                    if str(dept_code).strip() == '5' and str(sect_code).strip() == '11':
                        shares = DEFAULT_SHARE_DISTRIBUTION
                except (ValueError, TypeError):
                    pass

        if shares:
            for category, share_val in shares.items():
                new_row = row.copy()
                new_row['Amount1'] = row['Amount1'] * share_val
                new_row['category'] = category
                new_rows.append(new_row)
        else:
            new_row = row.copy()
            new_row['category'] = ''
            new_rows.append(new_row)
    return pd.DataFrame(new_rows)

def get_file_name(file_entry):
    file_name = fd.askopenfilename(title = "Select file",filetypes = (("CSV Files","*.csv"),))
    file_entry.delete(0,END)
    file_entry.insert(0,file_name)
    print("\n+++++++++++++++++++++++++++++++++++++++++++")
    print("Start Converting...")
    print("Processing the file:\n"+file_name)
    
    df = pd.read_csv(file_name,encoding = 'unicode_escape', engine ='python')
    
    # print(df)
    print(df.columns.tolist())
    print("\nRemoving unnecessary fields...")
    # Preserve dept_code, sect_code for share_distribution (Distribution condition)
    df["_share_dept_code"] = df["dept_code"] if 'dept_code' in df.columns else None
    df["_share_sect_code"] = df["sect_code"] if 'sect_code' in df.columns else None
    # del df['emp_code']
    del df['group_code']
    # del df['dept_code']
    del df['sect_code']
    del df['tax_allow']
    del df['nontax_allow']
    del df['other_tax_ded']
    del df['other_nontax_ded']
    del df['other_nontax_earn']
    # del df['dept_name']
    del df['group_name']
    del df['sect_name']
    del df['co_name']
    del df['co_address']
    del df['dept_rank']
    del df['atm_no']
    del df['with_atm']
    del df['last_pay']
    del df['payroll_date']

    with open('hr_pmr_list.csv', 'r') as file:
        reader = csv.DictReader(file)
        rows = list(reader)
        PMR_dict = {row['IHRIS']: row['BC'] for row in rows}
        # Share_dict: IHRIS -> {category: share_value} for IPI, SPI, Aldril, Randril, Totalcare (only non-zero values)
        Share_dict = {}
        share_columns = ['IPI', 'SPI', 'Aldril', 'Randril', 'Totalcare']
        for row in rows:
            ihris = row['IHRIS']
            shares = {}
            for csv_key, val in row.items():
                cat = csv_key.strip()
                if cat in share_columns and val:
                    try:
                        v = float(val)
                        if v > 0:
                            shares[cat] = v
                    except (ValueError, TypeError):
                        pass
            if shares:
                Share_dict[ihris] = shares
        
    DEPT_dict = {
    'DRZ' : 'DRZEN',
    'FIN' : 'FINANCE',
    'SOLV' : 'SALES09',
    'MEDS' : 'SALES08',
    '1' : 'SALES04',
    '2' : 'SALES05',
    '3' : 'HR',
    '4' : 'SALES06',
    '5' : 'DIST',
    '7' : 'TRAINING',
    '8' : 'REG',
    '9' : 'EXEC',
    'TRADE' : 'TRADEDEV',
    'SOLVANG' : 'SALES09',
    '20' : 'BUSDEV',
    'REG' : 'MEDICAL'
    }

    SSS_dict = {
    400 : 265,
    465 : 307.50,
    530 : 350,
    595 : 392.50,
    660 : 435,
    725 : 477.50,
    790 : 520,
    855 : 562.50,
    920 : 605,
    985 : 647.50,
    1050 : 690,
    1115 : 732.50,
    1180 : 775,
    1245 : 817.50,
    1310 : 860,
    1375 : 902.50,
    1440 : 945,
    1505 : 987.50,
    1570 : 1030,
    1635 : 1072.50,
    1700 : 1115,
    1765 : 1157.50,
    1830 : 1200,
    1895 : 1242.50,
    1980 : 1305,
    2045 : 1347.50,
    2110 : 1390,
    2175 : 1432.50,
    2240 : 1475,
    2305 : 1517.50,
    2370 : 1560,
    2435 : 1602.50,
    2500 : 1645,
    2565 : 1687.50,
    2630 : 1730,
    2695 : 1772.50,
    2760 : 1815,
    2825 : 1857.50,
    2890 : 1900,
    2955 : 1942.50,
    3020 : 1985,
    3085 : 2027.50,
    3150 : 2070,
    3215 : 2112.50,
    3280 : 2155
    }

    with open('hr_vendor_list.csv', 'r') as file:
        reader = csv.DictReader(file)
        Vendor_dict = {row['Name']: row['No_'] for row in reader}
        

    df["FullName"] = df["first_name"] + " " + df["last_name"]

    df["Sal. & Wages"] = df["basic_pay"] # 6101
    df["W/Tax comp."] = df["w_tax"] # 2301

    df["SSS Contribution Payable"] = df["sss_ee"].add(df["sss_er"].add(df["ecc_er"])) # 2505


    df["PHIC Contribution Payable"] = df["med_ee"].add(df["med_er"]) # 2508
    df["HDMF Contribution Payable"] = df["pagibig_ee"].add(df["pagibig_er"]) # 2506

    df["Sal. & Wages - Absent/Late"] = df["late_amt"].add(df["abs_amt"].add(df["under_amt"])) # 6101
    df["Sal. & Wages - OT"] = df["ot_amt"] # 6101
    df["Accrued Salaries & Wages"] = df["net_pay"] # 2401
    df["Accrued Commission & Incentive"] = df["other_tax_earn"] # 2402

    # df["EC - SSS"] = df["SSS Contribution Payable"].map(SSS_dict).fillna(0)
    # Update: Feb 08, 20023
    df["EC - SSS"] = df["sss_er"].add(df["ecc_er"]) # 6111

    df["EC - HDMF"] = 200 if '15' in n else 0
    df["EC - PHIC"] = df["PHIC Contribution Payable"].div(2)
    df["Loan Savings"] = df["loan_savings"].fillna(0) if 'loan_savings' in df else 0
    df["Savings"] = df["savings"].fillna(0) if 'savings' in df else 0
    df["Loan HDMF"] = df["loan_pagibig"].fillna(0) if 'loan_pagibig' in df else 0
    df["Loan SSS"] = df["loan_sss"].fillna(0) if 'loan_sss' in df else 0
    df["PMR Code1"] =  df["emp_code"].map(PMR_dict).fillna('')
    df["DEPT Code1"] =  df["dept_code"].map(DEPT_dict).fillna('')
    df["IHRIS"] = df["emp_code"]  # Keep IHRIS for share lookup before emp_code is deleted

    # Additional June 23, 2022
    df["Housing Loan"] = df["loan_housing"] if 'loan_housing' in df else 0
    df["Tablet Bond"] = df["bond_tablet"] if 'bond_tablet' in df else 0
    df["Cash Bond"] = df["bond_cash"] if 'bond_cash' in df else 0
    
    del df["loan_savings"]
    del df["loan_pagibig"]
    del df["loan_sss"]
    del df["savings"]
    del df["loan_housing"]
    del df["bond_tablet"]
    del df["bond_cash"]

    # del df["loan_savings"] if 'loan_savings' in df else tkinter.messagebox.showerror(title="Column Name Does Not Exist", message="Column Name [ loan_savings ] didn't exists in your CSV file, do you want to continue? ", **options)
    # del df["loan_pagibig"] if 'loan_pagibig' in df
    # del df["loan_sss"] if 'loan_sss' in df
    # del df["savings"] if 'savings' in df

    del df["emp_code"]
    del df["dept_code"]
    del df['first_name']
    del df['last_name']
    del df['mid_name']

    del df['basic_pay']
    del df['w_tax']

    del df['sss_ee']
    del df['sss_er']
    del df['ecc_er']
    del df['med_ee']
    del df['med_er']
    del df['pagibig_ee']
    del df['pagibig_er']

    del df['late_amt']
    del df['abs_amt']
    del df['under_amt']

    del df['ot_amt']
    del df['net_pay']
    del df['other_tax_earn']

    print("\n+++++++++++++++++++++++++++++++++++++++++++")
    print(df.columns.tolist())
    print("+++++++++++++++++++++++++++++++++++++++++++")

    Share_Account_dict = {
    'Sal. & Wages': '6101',
    'Sal. & Wages - OT': '6101',
    'EC - SSS' : '6111',
    'EC - HDMF' : '6113',
    'EC - PHIC' : '6112',
    }

    now = datetime.now()
    df2a = df.melt(id_vars=["FullName", "dept_name","PMR Code1","DEPT Code1","IHRIS","_share_dept_code","_share_sect_code"], var_name='Account Name', value_name='Amount1')

    df2a = share_distribution(df2a, Share_dict, Share_Account_dict)

    Account_dict = {
    'Sal. & Wages': '6101',
    'W/Tax comp.': '2301',
    'SSS Contribution Payable': '2505',
    'PHIC Contribution Payable': '2508',
    'HDMF Contribution Payable': '2506',
    'Sal. & Wages - Absent/Late': '6101',
    'Sal. & Wages - OT': '6101',
    'Accrued Salaries & Wages': '2401',
    'Accrued Commission & Incentive': '2402',
    'EC - SSS' : '6111',
    'EC - HDMF' : '6113',
    'EC - PHIC' : '6112',
    'Loan Savings' : '2901',
    'Savings' : '2900',
    'Loan HDMF' : '2507',
    'Loan SSS' : '2504',
    'Housing Loan' : '2901',
    'Tablet Bond' : '2902',
    'Cash Bond' : '2902',
    '13th Month' : '2404',
    'Absent_Lates' : '6101'
    }


    Amount_dict = {
    'W/TAX COMP.': -1,
    'SSS CONTRIBUTION PAYABLE': -1,
    'PHIC CONTRIBUTION PAYABLE': -1,
    'HDMF CONTRIBUTION PAYABLE': -1,
    'SAL. & WAGES - ABSENT/LATE': -1,
    'ACCRUED SALARIES & WAGES': -1,
    'SAVINGS': -1,
    'LOAN SAVINGS': -1,
    'LOAN HDMF': -1,
    'LOAN SSS': -1,
    'MAXICARE': -1,
    'MISC SMART': -1,
    'MISCELLANEOUS': -1,
    'PERSONAL MEDICINES': -1,
    'HOUSING LOAN': -1,
    'TABLET BOND': -1,
    'CASH BOND': -1,
    '13TH MONTH': -1,
    'ABSENT_LATES' : -1
    }
    #
    # 'SAL. & WAGES': 1,
    # 'SAL. & WAGES - OT': 1,
    # 'ACCRUED COMMISSION & INCENTIVE': 1,
    #
    
    df2a['Account No_'] = df2a["Account Name"].map(Account_dict).fillna('')

    # The purpose of [dept_name] is only for sorting thats why it was retain
    df2 = df2a.sort_values(by=['dept_name','FullName','Account No_'],ascending=[True, True, False],ignore_index=True)

    df2 = df2.drop(df2[ ((df2['Account Name'] == 'Loan Savings') & (df2['Amount1'] == 0)) |
    ((df2['Account Name'] == 'Savings') & (df2['Amount1'] == 0)) |
    ((df2['Account Name'] == 'Loan HDMF') & (df2['Amount1'] == 0)) |
    ((df2['Account Name'] == 'Loan SSS') & (df2['Amount1'] == 0))
    ].index)

    df2["Account Name_upper"] = df2["Account Name"].str.upper()
    df2['Line No'] = (df2.index + 1) * 10000
    df2['Recurring Method'] = 'V  Variable'
    df2['Recurring Frequency'] = '30D'
    df2['External Document No.'] = ''
    df2['Posting Date'] = now.strftime("%m/%d/%Y")
    df2['Document Type'] = ''
    df2['Document No.'] = n
    df2['Account Type'] = 'G/L Account'
    df2['Account No'] = df2['Account No_']
    df2['Description'] = df2["Account Name"] + ' - ' + df2["FullName"]
    df2['Amount'] = df2['Amount1'] * df2["Account Name_upper"].map(Amount_dict).fillna(1)
    df2['PMR Code'] = df2['PMR Code1'].fillna('')
    df2['DEPT Code'] = df2['DEPT Code1'].fillna('')

    #Locate name "MAXICARE" inside Data Frame df2["Account Name_upper"] then display/Select "Account Type" change value to "Vendor" (only row equal to "MAXICARE")
    #Same to others
    df2.loc[df2["Account Name_upper"] == 'MAXICARE', 'Account Type'] = 'Vendor'
    df2.loc[df2["Account Name_upper"] == 'MISC SMART', 'Account Type'] = 'Vendor'
    df2.loc[df2["Account Name_upper"] == 'MISCELLANEOUS', 'Account Type'] = 'Vendor'
    df2.loc[df2["Account Name_upper"] == 'PERSONAL MEDICINES', 'Account Type'] = 'Vendor'

    df2["FullName"] = df2["FullName"].str.upper()
    # Locate name "MAXICARE" inside Data Frame df2["Account Name_upper"] then display/Select "Account No" change value from map "Vendor_dict" (only row equal to "MAXICARE")
    # mapping the value of FullName if found it will display the equal value of in the Vendor_dict
    df2.loc[df2["Account Name_upper"] == 'MAXICARE', 'Account No'] = df2["FullName"].map(Vendor_dict).fillna('')
    df2.loc[df2["Account Name_upper"] == 'MISC SMART', 'Account No'] = df2["FullName"].map(Vendor_dict).fillna('')
    df2.loc[df2["Account Name_upper"] == 'MISCELLANEOUS', 'Account No'] = df2["FullName"].map(Vendor_dict).fillna('')
    df2.loc[df2["Account Name_upper"] == 'PERSONAL MEDICINES', 'Account No'] = df2["FullName"].map(Vendor_dict).fillna('')

    # Share category mapping: SPI, Aldril, Randril, Totalcare -> Vendor with specific Account No
    # IPI keeps Account Name, Account No, Account Type as is
    Category_Account_No_dict = {'SPI': 'AF0019', 'Aldril': 'AF0025', 'Randril': 'AF0007', 'Totalcare': 'AF0020'}
    for cat, acct_no in Category_Account_No_dict.items():
        mask = df2['category'] == cat
        df2.loc[mask, 'Account Type'] = 'Vendor'
        df2.loc[mask, 'Account No'] = acct_no

    del df2["Account Name_upper"]
    del df2['dept_name']
    del df2['Account Name']
    del df2['Amount1']
    del df2['FullName']
    del df2['Account No_']
    del df2['PMR Code1']
    del df2['DEPT Code1']
    del df2['category']
    del df2['IHRIS']
    del df2['_share_dept_code']
    del df2['_share_sect_code']

    # DEPT Code "REG" -> "MEDICAL" in final output
    df2.loc[df2['DEPT Code'] == 'REG', 'DEPT Code'] = 'MEDICAL'

    # Account Type "Vendor" -> PMR Code and DEPT Code blank
    df2.loc[df2['Account Type'] == 'Vendor', 'PMR Code'] = ''
    df2.loc[df2['Account Type'] == 'Vendor', 'DEPT Code'] = ''

    df2 = df2.drop(df2[ ((df2['Amount'] == 0)) ].index)
    df2 = df2.dropna(subset=['Amount'])

    dt_string = now.strftime("%m-%d-%Y %H.%M.%S")
    xlsFilename = r"Converted_To_BC365_Format_{}.xlsx".format(dt_string)
    df2.to_excel(xlsFilename, index=False)
    print("\nDone")
    print("\nSuccessfully converted to BC365 format")

    os.startfile(xlsFilename)

    close()


def run_and_close(event=None):
    print(entry_csv.get())
    ######################################
    ## EXECUTE OR CALL OTHER PYTHON FILE##
    ######################################
    close()

def close(event=None):
    master.withdraw() # if you want to bring it back
    # close()
    sys.exit() # if you want to exit the entire thing



master = Tk()
master.title("Convert CSV File to BC365 format")
entry_csv=Entry(master, text="", width=50)
entry_csv.grid(row=0, column=1, sticky=W, padx=5)

Label(master, text="Input CSV").grid(row=0, column=0 ,sticky=W)
Button(master, text="Browse and Process...", width=20, command=lambda:get_file_name(entry_csv)).grid(row=0, column=2, sticky=W)

# Button(master, text="Ok",     command=run_and_close, width=10).grid(row=3, column=1, sticky=E, padx=5)
Button(master, text="Close", command=close, width=20).grid(row=3, column=2, sticky=W)

master.bind('<Return>', run_and_close)
master.bind('<Escape>', close)
mainloop()
