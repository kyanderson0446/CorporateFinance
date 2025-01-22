# Data and DB Login
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
from deal_query import DatabaseFetcher
import json
import duckdb
import warnings
# File Management
import time
from datetime import datetime
import os
from glob import glob
# Excel
import xlwings as xw
# Algo
from matchOptions import FacilityNameMatcher
import rapidfuzz
import logging
from color_forms import bcolors
import numpy
from thefuzz import fuzz
from rapidfuzz import fuzz as rfuzz, process
from difflib import SequenceMatcher
import jellyfish
import re
from typing import List, Tuple, Optional, Union


"""

Intro:
Import the SQL Query
Source:
Utilize the Deal Acquisition Table to Maintain Integrity When Assigning Deal ID's to Files Stored on the Pdrive
Navigate:
Scrape Through Deals that Have an Assigned File Path and Assess if That File is the Same Deal in the Table
Match:
If There is a Confident Match, Assign the Deal ID to the DW Upload Tab in the Excel File and Load to Sharepoint.

#### ALGO LOGIC ####

Summary of Code:
This code aims to match a "potential deal" from a database (`df1`) to a corresponding entry in an Excel file 
based on facility name similarity and bed count proximity. The workflow adjusts thresholds dynamically to prioritize confident matches while allowing manual intervention for ambiguous cases.

1. Primary Matching Criteria:
   - Matches are determined by comparing:
     - Facility Name (string similarity, high threshold)
     - Bed Count (numerical proximity, within 3 beds).

2. Detailed Matching Steps:
   - Start with Name Match:
     - If the facility name similarity meets a high threshold, check the bed count proximity.
       - If the bed count is within 6, lower the name similarity threshold to confirm the match.
     - If no match is found, prompt for manual input to verify.

   - Fallback Matching:
     - If the initial name match fails, switch to a backup facility name ("in-file name").
     - Use a stricter name similarity threshold for this fallback check.

3. Output on Successful Match:
   - If a match is confirmed (automatically or manually), the corresponding `deal_id` is retrieved and 
     written to a designated range in the Excel sheet.

4. Output on No Match:
   - If no match is found at any stage, a message is logged, flagging the deal as unmatched.

Key Notes:
- The algorithm prioritizes reducing manual input by refining matching confidence through thresholds and 
  multiple fallback checks.
- Manual intervention serves as a last resort when automated matching cannot confidently confirm a deal.
- The use of dynamically adjusted thresholds (70%, 75%, and 85%) balances precision with flexibility in the matching process.

"""

save_proforma = r"E:\Unacquired"

# DB Login
with open('conn_string.json') as login_file:
    data = json.load(login_file)
url = data.get('url')
# Step 2: Initialize DatabaseFetcher and connect to DB
db_fetcher = DatabaseFetcher(url)
db_fetcher.connect()
# Step 3
df = db_fetcher.fetch_data(query="SELECT * FROM acquisition.public.deal", chunksize=10000)
# Proforma Table
df_loaded = db_fetcher.fetch_data(query="SELECT Distinct(acquisition_id) FROM public.proforma", chunksize=1000)
df1 = df.copy()
print(df1['file_path'].info())
try:
    df1['file_path'] = df1['file_path'].str.replace(r'\\PACS', '', regex=True) #VM doesn't play nice
except:
    df1['file_path'] = df1['file_path'].str.replace(r'\PACS', '', regex=True)  # VM doesn't play nice

# The 96 line proforma table we are popuplating
df2 = df_loaded.copy()
# Step 4: Close the database connection
db_fetcher.close_connection()
print("Data fetching complete")

"""
Logging
"""

# Configure logging to write to a log file
logging.basicConfig(
    filename=fr"{save_proforma}\no_match_log.txt",  # Name of the log file
    level=logging.INFO,           # Logging level
    format="%(asctime)s - %(message)s",  # Log format
    datefmt="%Y-%m-%d %H:%M:%S"         # Date and time format
)
"""
Log No Matches
"""
def log_no_match(potential_deal_id):
    logging.info(f"No Possible Match for: {potential_deal_id}")

"""
Determine if already logged:

Stop duplicates within the process from any match attempts or loads. Less work.
"""
def is_duplicate_deal(df_log, in_file_beds, file_facility_name, proforma):
    if df_log.empty:
        return False

    match_condition = (
            (df_log['Beds'] == in_file_beds) &
            (df_log['File Name'] == file_facility_name) &
            (df_log['File Path'] == proforma)
    )

    return match_condition.any()

"""
Log Matches and save:
Create in-process match logs and export to csv for review
"""
current_date = datetime.now().strftime("%Y-%m-%d")
csv_path = fr"{save_proforma}\{current_date} - matched_deals.csv" # Save out the matches log

def matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma):
    new_row = {
        "Deal ID": deal_id,
        "Beds": in_file_beds,
        "File Name": file_facility_name,
        "File Path": proforma
    }
    # Create Dataframe for each Deal Folder
    df_log = pd.concat([df_log, pd.DataFrame([new_row])], ignore_index=True)
    df_log.to_csv(csv_path, index=False)
    return df_log  # Return the updated dataframe

# Log files that have already been matched
df_log = pd.DataFrame(columns=["Deal ID", "Beds", "File Name", "File Path"])

"""
Start of Processing and File Loops
"""
warnings.simplefilter(action='ignore', category=FutureWarning)

for potential_deal in df1['facility_name']: # Loop through all deals as there will be duplicate names but unique deals tied to them
    # Is the load already pushed to the table? Stop the proforma that already exists
    
    deal_data = df1[df1['facility_name'] == potential_deal]
    potential_deal_path = deal_data['file_path'].iloc[0]
    potential_deal_id = deal_data['id'].iloc[0]
    
    if potential_deal_id in df2['acquisition_id']:
        continue
    
    try:
        if potential_deal_path is not None:
            pass
        else:
            print(f"Deal: {potential_deal} no file path.")
            continue

        # Extract the licensed_op_beds value from that row
        if not deal_data.empty:
            potential_deal_beds = deal_data['licensed_op_beds'].iloc[0]
            print()
            print(f"{bcolors.OKGREEN}Deal ID: {bcolors.OKBLUE}{potential_deal_id}{bcolors.OKGREEN}, Facility: {bcolors.OKBLUE}{potential_deal}{bcolors.OKGREEN}, Licensed Beds: {bcolors.OKBLUE}{potential_deal_beds}\n")
        else:
            print()
            print(f"{bcolors.FAIL}Facility: {potential_deal} not found in the DataFrame.\n")

        # for example: P:\PACS\Finance\Acquistions & New Build\Active\2024 - RealSG AZ-4\Proforma

        # Extract the file information
        for files in glob(potential_deal_path):
            proformas_path = fr"{potential_deal_path}\*Proforma.xlsx"
            print(f"{bcolors.OKGREEN}Looping Through Files in Folder")

            files_in_folder = []
            df_excel = pd.DataFrame(columns=["In File Name", "Beds", "File Name", "File Path"])


            for proforma in glob(proformas_path):
                proforma_file = os.path.basename(proforma)
                proforma_file = proforma_file.replace('.xlsx', '')
                file_facility_name = proforma_file.split(" - ")[1] if len(proforma_file.split(" - ")) > 1 else None
                files_in_folder.append(file_facility_name)

                # Extract the exact details from the excel workbook

                conn = duckdb.connect(database=':memory:', read_only=False)
                conn.execute('INSTALL spatial;')
                conn.execute('LOAD spatial;')

                # Forecast or Budget Worksheet
                try:
                    query_str = fr"""SELECT * FROM st_read('{proforma}', layer='FACILITY INFO');"""
                    df_file = conn.execute(query_str).df()

                except:
                    continue


                in_file_name = df_file.iloc[6, 1]
                in_file_beds = float(df_file.iloc[9, 1])
                new_row = {
                    "In File Name": in_file_name,
                    "Beds": in_file_beds,
                    "File Name": file_facility_name,
                    "File Path": proforma
                }
                # Create Dataframe for each Deal Folder

                df_excel = pd.concat([df_excel, pd.DataFrame([new_row])], ignore_index=True)

                if is_duplicate_deal(df_log, in_file_beds, file_facility_name, proforma):
                    print(f"{bcolors.WARNING}Deal already logged {file_facility_name}. Skipping...")
                    continue  # Skip to next iteration


                ###
                # Execute a matcher based on file name here
                ####

                # Match Deal to Files
                """
                1. Proforma -> File Name 
                """
                match_1 = FacilityNameMatcher.match_facility_name(potential_deal, files_in_folder, threshold=98)
                if match_1[0] is not None:  # Check explicitly if match is not None
                    matched_value = match_1[0]

                    temp_df = df_excel[df_excel['File Name'] == matched_value]
                    file_beds = float(temp_df['Beds'].iloc[0])
                    index = df_excel[df_excel['File Name'] == matched_value].index[0]
                    """
                    1. Proforma -> File Name -> Bed
                    """
                    if abs(potential_deal_beds - file_beds) < 6:
                        # Lower Threshold Based on Bed Count Proximity, Drill Down to In-File Name
                        """
                        1. Proforma -> File Name -> Bed -> In File Name
                        """
                        file_name_choices = temp_df['In File Name'].to_list()
                        match_5 = FacilityNameMatcher.match_facility_name(potential_deal, file_name_choices, threshold=95)
                        if match_5[0] is not None:  # Check explicitly if match is not None
                            in_file_name_match = match_5[0]
                            # If there a duplicate In File Names from the files
                            in_file_index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {in_file_name_match}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                            else:
                                conn.close()
                                continue
                            # ID
                            locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                            open_file = locate_path['File Path'].iloc[0]
                            app = xw.App(visible=False)
                            xw.Visible = False
                            xw.ScreenUpdating = False
                            xw.Interactive = False
                            wb = xw.Book(open_file, update_links=False)
                            wb.sheets("DW Upload").range("B3").value = potential_deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                            wb.save(fr"{save_proforma}\{potential_deal_id} - {potential_deal}.xlsx")
                            wb.close()
                            app.quit()
                            # create a matches log
                            df_log = matches_log(df_log, potential_deal_id, in_file_beds, file_facility_name, proforma)
                            conn.close()
                            break
                        # If No Match, Flag and Manually Determine
                        else:
                            """
                            2. Proforma -> In File Name -> Bed -> Manual
                            """
                            manual_match = str(
                                input(
                                    fr"{bcolors.WARNING} Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))

                            if manual_match == True or str(manual_match).lower() == 'true':
                                print(f"{bcolors.OKGREEN}Using 'Manual': Match Found\n")
                                # ID
                                locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                                open_file = locate_path['File Path'].iloc[0]
                                app = xw.App(visible=False)
                                xw.Visible = False
                                xw.ScreenUpdating = False
                                xw.Interactive = False
                                wb = xw.Book(open_file, update_links=False)
                                wb.sheets("DW Upload").range("B3").value = potential_deal_id
                                wb.sheets("DW Upload").range("B4").value = 'default'
                                wb.save(fr"{save_proforma}\{potential_deal_id} - {potential_deal}.xlsx")
                                wb.close()
                                app.quit()
                                df_log = matches_log(df_log, potential_deal_id, in_file_beds, file_facility_name, proforma)
                                conn.close()
                                break
                            else:
                                print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")
                                conn.close()

                    # This is When the File Name was a Close Match but the Bed Count Fails
                    # Increase the Threshold Here as it Requires Matching on File Name and In_File Name
                    else:
                        """
                        3. Proforma -> File Name -> In File Name. Beds were not close but try to match
                        """
                        file_name_choices = temp_df['In File Name'].to_list()
                        match_4 = FacilityNameMatcher.match_facility_name(potential_deal, file_name_choices, threshold=95)
                        if match_4[0] is not None:  # Check explicitly if match is not None
                            in_file_name_match = match_4[0]
                            index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]
                            in_file_index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {in_file_name_match}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                            else:
                                conn.close()
                                continue
                            # ID
                            locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                            open_file = locate_path['File Path'].iloc[0]
                            app = xw.App(visible=False)
                            xw.Visible = False
                            xw.ScreenUpdating = False
                            xw.Interactive = False
                            wb = xw.Book(open_file, update_links=False)
                            wb.sheets("DW Upload").range("B3").value = potential_deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                            wb.save(fr"{save_proforma}\{potential_deal_id} - {potential_deal}.xlsx")
                            wb.close()
                            app.quit()
                            df_log = matches_log(df_log, potential_deal_id, in_file_beds, file_facility_name, proforma)
                            conn.close()
                            break
                        else:
                            print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")
                            conn.close()

                # If There is a Weak Match on Deal to File Name, Try the In-File Name as an Alternative
                else:
                    """
                    2. Proforma -> In File name
                    """
                    file_name_choices = df_excel['In File Name'].to_list()
                    match_2 = FacilityNameMatcher.match_facility_name(potential_deal, file_name_choices, threshold=85)
                    if match_2[0] is not None:  # Check explicitly if match is not None
                        in_file_name_match = match_2[0]
                        temp_df = df_excel[df_excel['In File Name'] == in_file_name_match]
                        file_beds = float(temp_df['Beds'].iloc[0])
                        index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]

                        # Higher Threshold as This is the Backup
                        """
                        2. Proforma -> In File Name -> Bed
                        """
                        if abs(potential_deal_beds - in_file_beds) < 6:
                            file_name_choices = temp_df['In File Name'].to_list()
                            in_file_index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {in_file_name_match}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                                # ID
                                locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                                open_file = locate_path['File Path'].iloc[0]
                                app = xw.App(visible=False)
                                xw.Visible = False
                                xw.ScreenUpdating = False
                                xw.Interactive = False
                                wb = xw.Book(open_file, update_links=False)
                                wb.sheets("DW Upload").range("B3").value = potential_deal_id
                                wb.sheets("DW Upload").range("B4").value = 'default'
                                wb.save(fr"{save_proforma}\{potential_deal_id} - {potential_deal}.xlsx")
                                wb.close()
                                app.quit()
                                df_log = matches_log(df_log, potential_deal_id, in_file_beds, file_facility_name, proforma)
                                conn.close()
                                break
                            else:
                                """
                                2. Proforma -> In File Name -> Bed -> Manual
                                """
                                manual_match = str(
                                    input(
                                        fr"{bcolors.WARNING}Deal:{potential_deal} -> {in_file_name}, is it a match? True or False: "))
                                if manual_match == True or str(manual_match).lower() == 'true':
                                    print(f"{bcolors.OKGREEN}Using 'Manual': Match Found\n")
                                    # ID
                                    locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                                    open_file = locate_path['File Path'].iloc[0]
                                    app = xw.App(visible=False)
                                    xw.Visible = False
                                    xw.ScreenUpdating = False
                                    xw.Interactive = False
                                    wb = xw.Book(open_file, update_links=False)
                                    wb.sheets("DW Upload").range("B3").value = potential_deal_id
                                    wb.sheets("DW Upload").range("B4").value = 'default'
                                    wb.save(fr"{save_proforma}\{potential_deal_id} - {potential_deal}.xlsx")
                                    wb.close()
                                    app.quit()
                                    df_log = matches_log(df_log, potential_deal_id, in_file_beds, file_facility_name, proforma)
                                    conn.close()
                                    break
                                else:
                                    print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")
                                    conn.close()

                        else:
                            print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")
                            conn.close()


    except:
        print(f"{bcolors.FAIL}***No Possible Match for: {potential_deal_id}***")
        log_no_match(potential_deal_id)

