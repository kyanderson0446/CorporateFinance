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
plan = 'Budget 2024 Q3 - Consolidated'
structure = 'Consolidated'
##############

# paths
#################### ENTER FILE PATH TO FOLDER HERE ############################
path = ""
map = '96linemap.xlsx'
map = pd.read_excel(map, index_col=False)
##############

app = xw.App(visible=False)

for file in glob(path):
    wb = xw.Book(file, update_links=False)
    file_name = os.path.basename(file).split(' - ')[1]

    # Fix data DW Upload
    dwupload = wb.sheets("DW Upload")
    dwupload.visible = True
    dwupload.api.Unprotect(Password='PGHCbudget!')
    dwupload.range('17:27').delete()
    dwupload.range('19:19').delete()
    dwupload.range('D18').value = 'Ancillary and Other Revenue'
    dwupload.range('20:22').delete()
    dwupload.range('C174:C181').value = ''
    dwupload.range('C190:C259').value = ''
    dwupload.range('C266:C360').value = ''
    dwupload.range('19:19').api.Font.Underline = False
    # Fix formulas for DW Upload
    try:
        dwupload.range('E5').formula = fr"='BUDGET WORKSHEET'!G9"

    except:
        dwupload.range('E5').formula = fr"='FORECAST WORKSHEET'!G9"

    f19 = "=IFNA(INDEX(Budget60Month,MATCH('DW Upload'!$A19,'60-Month Report'!$A$5:$A$197,0),MATCH('DW Upload'!F$5,'60-Month Report'!$A$5:$BJ$5,0)),0)"
    dwupload.range('F19:BL19').formula = f19
    f20 = "=IFNA(INDEX(Budget60Month,MATCH('DW Upload'!$A20,'60-Month Report'!$A$5:$A$197,0),MATCH('DW Upload'!F$5,'60-Month Report'!$A$5:$BJ$5,0)),0)"
    dwupload.range('F20:BL20').formula = f20
    f21 = "=IFNA(INDEX(Budget60Month,MATCH('DW Upload'!$A21,'60-Month Report'!$A$5:$A$197,0),MATCH('DW Upload'!F$5,'60-Month Report'!$A$5:$BJ$5,0)),0)"
    dwupload.range('F21:BL21').formula = f21
    f22 = "=IFNA(INDEX(Budget60Month,MATCH('DW Upload'!$A22,'60-Month Report'!$A$5:$A$197,0),MATCH('DW Upload'!F$5,'60-Month Report'!$A$5:$BJ$5,0)),0)"
    dwupload.range('F22:BL22').formula = f22
    e23 = "=SUM(E19:E22)"
    dwupload.range('E23:E23').formula = e23
    f23 = "=SUM(F19:F22)"
    dwupload.range('F23:BL23').formula = f23
    e25 = "=SUM(E23,E16)"
    dwupload.range('E25:E25').formula = e25
    f25 = "=SUM(f23,f16)"
    dwupload.range('F25:BL25').formula = f25

    # Fix Formulas for 60-Month Report tab
    # Unlock and Unhide Sheets
    sixtyMonth = wb.sheets('60-Month Report')
    sixtyMonth.visible = True
    sixtyMonth.api.Unprotect(Password='budgetreport')
    all_lines = wb.sheets('RPT - All Lines')
    all_lines.visible = True
    all_lines.api.Unprotect(Password='budgetreport')

    c31 = "='RPT - All Lines'!C31+C30"
    sixtyMonth.range('D31:BJ31').formula = c31


    ledger_values = dwupload.range('A8:A357').value

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
    dwupload.range('B4').value = 'ref'
    dwupload.range('B5').value = plan
    dwupload.range('B6').value = structure
    print("done")
    # Save
    time.sleep(1)
    wb.save(fr"path\Upload\{file_name}.xlsx")
    time.sleep(2)
    wb.close()

