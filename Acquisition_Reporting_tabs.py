import re
import os
import time


import duckdb
import pandas as pd
from glob import glob
import openpyxl
import xlwings as xw
import numpy as np


start_date = str(input("What is the start date? MM/DD/YYYY:  "))

path_bool = str(input("Do you know the exact file path of the proformas? Y/N: ")).lower()
if path_bool == 'y':
    acquisition_group = str(input("Enter the acquisition group folder name: "))
    folder_entry = input("Paste the exact path of the acquisition folder: ").strip()

    # Remove any leading or trailing quotation marks
    if folder_entry.startswith('"') and folder_entry.endswith('"'):
        folder_entry = folder_entry[1:-1]

    folder_entry = folder_entry + r'\*.xlsx'
    # general_path = folder_entry.replace(fr"\PACS", "")
    general_path = folder_entry


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
        general_path = fr"P:\PACS\Finance\Acquistions & New Build\{acquisition_type}\{acquisition_group}\Proformas\{yes}\*.xlsx"
    else:
        general_path = fr"P:\PACS\Finance\Acquistions & New Build\{acquisition_type}\{acquisition_group}\Proformas\*.xlsx"
        pass

print("-"*20)
print()
print("The files will be saved at:  Finance\Acquistions & New Build\REPORTING_COMPILE")
print("-"*20)
save_path = fr"P:\PACS\Finance\Acquistions & New Build\REPORTING_COMPILE"

#######################################################################

# Create the template file to enter the sheet data into
def create_excel_if_not_exists(save_path, acquisition_group):
    master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING_test.xlsx"
    if not os.path.isfile(master_xl_file):
        wb = openpyxl.Workbook()
        wb.save(master_xl_file)
        wb.close()
        time.sleep(1)
        print(f"File {master_xl_file} created successfully.")
    else:
        print(f"File {master_xl_file} already exists.")

# Create the file if the code hasn't been run already for this acquisition
create_excel_if_not_exists(save_path, acquisition_group)

master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING_test.xlsx"

# Connect to duckdb and extract data
conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')

# File names to be sheet names
name_list = []
seen_names = set()

for file in glob(general_path):
    print(file)
    file_name = os.path.basename(file).split(' - ')[1]
    print(file_name)

    # If the name is not in the set, it's unique
    if file_name not in seen_names:
        name_list.append(file_name)
        seen_names.add(file_name)
    else:
        # Handle duplicates by appending (1), (2), etc.
        base_name = file_name
        counter = 1
        new_name = f"{base_name}({counter})"
        while new_name in seen_names:
            counter += 1
            new_name = f"{base_name}({counter})"
        name_list.append(new_name)
        seen_names.add(new_name)

    query_str = fr"""SELECT * FROM st_read('{file}', layer='DW Upload');"""
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

# wb.sheets['Sheet1'].delete()

x = 0
for sheet in wb.sheets:
    # Delete the first row (header row)
    if 'Sheet' in sheet.name:
        sheet.delete()
        continue
    sheet.range('1:1').delete()
    sheet.range('A1').value = name_list[x]

    sheet.name = sheet.range('A1').value
    x += 1

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
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "AL - SSI Revenue",
    "AL - Private Revenue",
    "",
    "",
    "Ancillary Medicare B",
    "Ancillary Other Revenue",
    "AL - Private Revenue",
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
    "Assisted Living Caregiver Wages",
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
    "",
    "",
    "Medicare Census",
    "Medicaid Census",
    "Private Census",
    "Managed Care Census",
    "Hospice Census",
    "Insurance Census",
    "Veterans Census",
    "AL Census"
]

# Loop through each sheet
for sheet in wb.sheets:
    # Paste the values starting from cell A8
    sheet.range('A8').value = [[value] for value in values]
    sheet.range('B:B').delete()
    sheet.range('B:B').delete()
    # Convert format to float
    cell_range = sheet.range('C8:BZ400')
    data = pd.DataFrame(sheet.range(cell_range.address).value)
    def convert_to_number(value):
        try:
            return pd.to_numeric(value, errors='coerce')
        except (ValueError, TypeError):
            return value
    data = data.applymap(convert_to_number)
    sheet.range(cell_range.address).value = data.values.tolist()
wb.save()
wb.close()

# delete 191 and onwards?

