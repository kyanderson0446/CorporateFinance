import re
import os
import time


import duckdb
import pandas as pd
from glob import glob
import openpyxl
import xlwings as xw


start_date = str(input("What is the start date? MM/DD/YYYY"))

path_bool = str(input("Do you know the exact file path of the proformas? Y/N")).lower()
if path_bool == 'y':
    folder_entry = str(input("Paste the exact path of the folder:"))
    folder_entry = folder_entry.replace(fr"\PACS", "")
    print("-" * 20)
    print()
else:
    print("-"*20)
    print("Case sensitive, Please Match the Exact Folder Name")
    print("-"*20)
    print()
    acquisition_group = str(input("Enter the acquisition group folder name: "))
    acquisition_type = str(input("Enter the acquisition type\n Active, Closing or Completed Buildings: "))
    lower_folder = str(input("Is there a folder that is further down from the Proforma's folder? Y or N")).lower()
    if lower_folder == 'Y':
        yes = str(input("Please enter the subfolder:"))
        general_path = fr"P:\Finance\Acquistions & New Build\{acquisition_type}\{acquisition_group}\Proformas\{yes}\*.xlsx"
    else:
        general_path = fr"P:\Finance\Acquistions & New Build\{acquisition_type}\{acquisition_group}\Proformas\*.xlsx"
        pass

print("-"*20)
print()
print("The files will be saved at:  Finance\Acquistions & New Build\REPORTING_COMPILE")
print("-"*20)
save_path = fr"P:\Finance\Acquistions & New Build\REPORTING_COMPILE"

#######################################################################

def create_excel_if_not_exists(save_path, acquisition_group):
    master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING.xlsx"
    if not os.path.isfile(master_xl_file):
        wb = openpyxl.Workbook()
        wb.save(master_xl_file)
        wb.close()
        time.sleep(1)
        print(f"File {master_xl_file} created successfully.")
    else:
        print(f"File {master_xl_file} already exists.")


create_excel_if_not_exists(save_path, acquisition_group)


master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING.xlsx"

conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')


for file in glob(general_path):
    print(file)
    query_str = fr"""SELECT * FROM st_read('{file}', layer='REPORTING');"""
    df = conn.query(query_str).df()
    facility_name = df.iloc[0, 0]
    facility_name = re.sub(r'[^\w\s]', '', facility_name)
    with pd.ExcelWriter(master_xl_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, sheet_name=facility_name, index=False)

print()
print(fr'{acquisition_group} has been loaded')

############# xlwings #############
file = master_xl_file
# Open the workbook
wb = xw.Book(file, update_links=False)

wb.sheets['Sheet1'].delete()

for sheet in wb.sheets:
    # Delete the first row (header row)
    sheet.range('1:1').Delete()

values = [
    "Medicare Revenue",
    "Medicaid Revenue",
    "Private Revenue",
    "Managed Care Revenue",
    "Hospice Revenue",
    "Insurance Revenue",
    "Veterans Revenue",
    "Other Payers Revenue",
    "",
    "",
    "",
    "",
    "",
    "Assisted Living Revenue",
    "",
    "",
    "Ancillary Medicare B",
    "Ancillary Other Revenue",
    "Other Revenue",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Nursing Management Wages",
    "Nurses with Admin Duties Wages",
    "RN Wages",
    "LVN/LPN Wages",
    "CNA Wages",
    "Other Nursing Wages",
    "Assisted Living Wages",
    "",
    "Maintenance Wages",
    "Housekeeping Wages",
    "Laundry Wages",
    "Dietary Wages",
    "Social Services Wages",
    "Activities Wages",
    "Education Wages",
    "Administration Wages",
    "Barber & Beauty Wages",
    "",
    "",
    "",
    "Payroll Taxes",
    "Workers Comp Insurance",
    "Health Insurance",
    "Paid Time Off (Vac, Hol, Sick)",
    "Other Benefits",
    "",
    "",
    "",
    "RN Agency",
    "LVN/LPN Agency",
    "CNA Agency",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Nurs - Supplies & Minor Equipment",
    "Nurs - Medical Director",
    "Nurs - Medical Waste",
    "Nurs - Other Expenses",
    "",
    "",
    "",
    "Maint - Supplies & Minor Equipment",
    "Maint - Repairs & Maintenance",
    "Maint - Purchased Services",
    "Maint - Other Expenses",
    "",
    "",
    "",
    "Plant Op - Grounds",
    "Plant Op - Utilities",
    "",
    "",
    "",
    "Hskp - Supplies & Paper Goods",
    "Hskp - Other Expenses",
    "",
    "",
    "",
    "Lndy - Supplies & Cleaning Supplies",
    "Lndy - Linen & Bedding",
    "Lndy - Repairs & Maintenance",
    "Lndy - Other Expenses",
    "",
    "",
    "",
    "Diet - Raw Food",
    "Diet - Tube Feeding",
    "Diet - Dietician Fees",
    "Diet - Supplies & Minor Equipment",
    "Diet - Repairs & Maintenance",
    "Diet - Other Expenses",
    "",
    "",
    "",
    "Soc Srvc - Property Replacement",
    "Soc Srvc - Other Expenses",
    "",
    "",
    "",
    "Act - Supplies",
    "Act - Outside Entertainment",
    "Act - Other Expenses",
    "",
    "",
    "",
    "Educ - Supplies",
    "Educ - Other Supplies",
    "",
    "",
    "",
    "Adm - Supplies & Minor Equipment",
    "Adm - Repairs & Maintenance",
    "Adm - Professional Fees",
    "Adm - Service Agreements",
    "Adm - Software Support",
    "Adm - Licenses",
    "Adm - Liability Insurance",
    "Adm - Marketing & Advertising",
    "Adm - Admin Services Fees",
    "Adm - Direct Expense Allocation",
    "Adm - Quality Assurance Fees",
    "Adm - Other Expenses",
    "",
    "",
    "",
    "Prop - Depreciation & Amortization",
    "Prop - Property Leases",
    "Prop - Property Taxes",
    "Prop - Property Insurance",
    "Prop - Interest",
    "",
    "",
    "Total Bad Debt Expense",
    "",
    "",
    "Anc - Patient Supplies",
    "Anc - Physical Therapy",
    "Anc - Occupational Therapy",
    "Anc - Speech Therapy",
    "Anc - Pharmacy",
    "Anc - Laboratory",
    "Anc - Physician Fees",
    "Anc - Radiology",
    "Anc - Respiratory Therapy",
    "Anc - Transportation",
    "Anc - Other Expenses",
    "",
    "",
    "",
    "",
    "",
    "",
    "Total NonOperating Income",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Number of SNF beds",
    "",
    "Number of ALF beds",
    "",
    "",
    "Medicare Census",
    "Medicaid Census",
    "Private Census",
    "Managed Care Census",
    "Hospice Census",
    "Insurance Census",
    "Veterans Census",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Assisted Living Census"
]

# Loop through each sheet
for sheet in wb.sheets:
    # Paste the values starting from cell A8
    sheet.range('A8').value = [[value] for value in values]



