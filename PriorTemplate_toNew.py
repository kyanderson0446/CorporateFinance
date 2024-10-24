import pandas as pd
import xlwings as xw
from glob import glob
import os
import openpyxl
import concurrent.futures
import time


"""
Map old proformas to new template with the Workday ledgers
"""
##############
# Input
plan = 'Budget 2024 Q4 - Consolidated'
structure = 'Consolidated'
##############

# paths
#################### ENTER FILE PATH TO FOLDER HERE ############################
path = fr""
map = '96linemap.xlsx'
map = pd.read_excel(map, index_col=False)
##############

app = xw.App(visible=False)

def column_index_to_letter(index):
    letter = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


for file in glob(path):
    wb = xw.Book(file, update_links=False)
    # Get the facility name
    try:
        base_name = os.path.basename(file)
        base_name.replace(".xlsx", "")
        file_name = os.path.basename(base_name).split(' - ')[1]
    except:
        base_name = os.path.basename(file)
        base_name.replace(".xlsx", "")
        file_name = base_name.replace(".xlsx","")
    print(file_name)

    # Fix data DW Upload
    dwupload = wb.sheets("DW Upload")
    sixtyMonth = wb.sheets('60-Month Report')
    dwupload.visible = False
    dwupload.api.Unprotect(Password='PGHCbudget!')

    # Fix dates for DW Upload
    try:
        dwupload.range('E5').formula = fr"='BUDGET WORKSHEET'!G9"

    except:
        dwupload.range('E5').formula = fr"='FORECAST WORKSHEET'!G9"

    # Add whatever from medicare to 3100040...
    # Define the formulas for each row

    # E8 Medicare
    for col in range(5, 64): # E through BL
        # Get the Excel column letter from its index
        col_letter = column_index_to_letter(col)
        value_row_8 = dwupload.range(f'{col_letter}8').value or 0  # Use 0 if cell is empty (None)
        value_row_19 = dwupload.range(f'{col_letter}19').value or 0  # Use 0 if cell is empty (None)
        E8_medicare = value_row_8 + value_row_19
        dwupload.range(f'{col_letter}8').value = E8_medicare

    # E9_medicaid
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_9 = dwupload.range(f'{col_letter}9').value or 0
        value_row_20 = dwupload.range(f'{col_letter}20').value or 0
        E9_medicaid = value_row_9 + value_row_20
        dwupload.range(f'{col_letter}9').value = E9_medicaid

    # E10_private
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_10 = dwupload.range(f'{col_letter}10').value or 0
        value_row_21 = dwupload.range(f'{col_letter}21').value or 0
        E10_private = value_row_10 + value_row_21
        dwupload.range(f'{col_letter}10').value = E10_private

    # E11_managed
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_11 = dwupload.range(f'{col_letter}11').value or 0
        value_row_22 = dwupload.range(f'{col_letter}22').value or 0
        E11_managed = value_row_11 + value_row_22
        dwupload.range(f'{col_letter}11').value = E11_managed

    # E12_hospice
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_12 = dwupload.range(f'{col_letter}12').value or 0
        value_row_23 = dwupload.range(f'{col_letter}23').value or 0
        E12_hospice = value_row_12 + value_row_23
        dwupload.range(f'{col_letter}12').value = E12_hospice

    # E13_insurance
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_13 = dwupload.range(f'{col_letter}13').value or 0
        value_row_24 = dwupload.range(f'{col_letter}24').value or 0
        E13_insurance = value_row_13 + value_row_24
        dwupload.range(f'{col_letter}13').value = E13_insurance

    # E14_veterans
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_14 = dwupload.range(f'{col_letter}14').value or 0
        value_row_25 = dwupload.range(f'{col_letter}25').value or 0
        E14_veterans = value_row_14 + value_row_25
        dwupload.range(f'{col_letter}14').value = E14_veterans

    # E15_other
    for col in range(5, 64):
        col_letter = column_index_to_letter(col)
        value_row_15 = dwupload.range(f'{col_letter}15').value or 0
        value_row_26 = dwupload.range(f'{col_letter}26').value or 0
        E15_other = value_row_15 + value_row_26
        dwupload.range(f'{col_letter}15').value = E15_other


    e31 = "=IFNA(INDEX(Budget60Month, MATCH('DW Upload'!$A31, '60-Month Report'!$A$5:$A$197, 0), MATCH('DW Upload'!E$5, '60-Month Report'!$A$5:$BJ$5, 0)), 0)"
    dwupload.range('E31').formula = e31
    e35 = "=IFNA(INDEX(Budget60Month, MATCH('DW Upload'!$A35, '60-Month Report'!$A$5:$A$197, 0), MATCH('DW Upload'!E$5, '60-Month Report'!$A$5:$BJ$5, 0)), 0)"
    dwupload.range('E35').formula = e35
    e36 = "=IFNA(INDEX(Budget60Month, MATCH('DW Upload'!$A36, '60-Month Report'!$A$5:$A$197, 0), MATCH('DW Upload'!E$5, '60-Month Report'!$A$5:$BJ$5, 0)), 0)"
    dwupload.range('E36').formula = e36
    e37 = "=IFNA(INDEX(Budget60Month, MATCH('DW Upload'!$A37, '60-Month Report'!$A$5:$A$197, 0), MATCH('DW Upload'!E$5, '60-Month Report'!$A$5:$BJ$5, 0)), 0)"
    dwupload.range('E37').formula = e37

    formula_e31 = dwupload.range('E31').formula
    formula_e35 = dwupload.range('E35').formula
    formula_e36 = dwupload.range('E36').formula
    formula_e37 = dwupload.range('E37').formula
    dwupload.range('E31:BL31').formula = formula_e31
    dwupload.range('E35:BL35').formula = formula_e35
    dwupload.range('E36:BL36').formula = formula_e36
    dwupload.range('E37:BL37').formula = formula_e37


    # Assign SA lines to Routine Revenue
    dwupload.range('C17:C27').value = ''
    dwupload.range('A17:A27').value = ''
    dwupload.range('C30:C30').value = ''
    dwupload.range('D29').value = 'Ancillary and Other Revenue'
    dwupload.range('C197:C204').value = ''
    dwupload.range('C208:C274').value = ''
    dwupload.range('C281:C372').value = ''
    dwupload.range('19:19').api.Font.Underline = False

    # Fix 60-Month Report for the AL private revenue
    sixtyMonth.visible = False
    sixtyMonth.api.Unprotect(Password='budgetreport')
    all_lines = wb.sheets('RPT - All Lines')
    all_lines.visible = False
    all_lines.api.Unprotect(Password='budgetreport')
    c31 = "='RPT - All Lines'!C31+C30"
    sixtyMonth.range('D31:BJ31').formula = c31

    # Old ledger range
    ledger_values = dwupload.range('A8:A372').value

    # Iterate over the values and update based on the DataFrame
    for i, value in enumerate(ledger_values):
        if value in map['96line'].values:
            # Get the corresponding Ledger value
            ledger_value = map.loc[map['96line'] == value, 'Ledger'].values[0]
            # Update the corresponding cell in column B (or whichever column you want)
            dwupload.range(f'C{8 + i}').value = ledger_value
    # Reformat values to number
    dwupload.range('C:C').number_format = '0'
    # Upload section
    dwupload.range('B3').value = 'Budget'
    dwupload.range('B4').value = 'REF_ID'
    dwupload.range('B5').value = plan
    dwupload.range('B6').value = structure
    print("done")

    # dwupload = dwupload.copy(after=dwupload)
    used_range = dwupload.used_range
    used_range.value = used_range.value
    dwupload.range('17:27').delete()
    dwupload.range('19:19').delete()
    dwupload.range('D18').value = 'Ancillary and Other Revenue'
    dwupload.range('20:22').delete()
    dwupload.range('19:19').api.Font.Underline = False

    # Save
    time.sleep(1)
    wb.save()
    time.sleep(2)
    wb.close()

