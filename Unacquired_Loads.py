# Data and DB Login
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
from deal_query import DatabaseFetcher
import json
# File Management
import time
import os
from glob import glob
# Excel
import xlwings as xw
# Algo
import rapidfuzz
from rapidfuzz import process, fuzz


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
     - Facility Name (string similarity, thresholds of 70%-85%).
     - Bed Count (numerical proximity, within 3 beds).

2. Detailed Matching Steps:
   - Start with Name Match:
     - If the facility name similarity meets a 75% threshold, check the bed count proximity.
       - If the bed count is within 3, lower the name similarity threshold to 70% to confirm the match.
     - If no match is found, prompt for manual input to verify.

   - Fallback Matching:
     - If the initial name match fails, switch to a backup facility name ("in-file name").
     - Use a stricter 85% name similarity threshold for this fallback check.

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


# DB Login
with open('conn_string.json') as login_file:
    data = json.load(login_file)
url = data.get('url')
# Step 2: Initialize DatabaseFetcher and connect to DB
db_fetcher = DatabaseFetcher(url)
db_fetcher.connect()
# Step 3
df = db_fetcher.fetch_data(query="SELECT * FROM acquisition.public.deal", chunksize=10000)
df1 = df.copy()
# Step 4: Close the database connection
db_fetcher.close_connection()
print("Data fetching complete.")


# deal_id = df1['id']
# deal_name = df1['deal']
# deal_facility_name = df1['facility_name']
# deal_beds = df1['licensed_op_beds']

# deal_df = df1[['id', 'deal', 'facility_name', 'licensed_op_beds']]
# Load Excel file
# Get the filename and the name inside the file to confirm.

# Create a function to match facility names


def match_facility_name(row, choices, threshold=80):
    """
    Matches a single facility name (row) against a list of choices using RapidFuzz.

    Args:
        row (str): The facility name to match.
        choices (list): List of possible facility names to match against.
        threshold (int): Minimum match score to consider a match.

    Returns:
        str: The best matching facility name if score >= threshold, else None.
    """
    if not choices or not isinstance(row, str):
        return None  # No choices or invalid input

    # Use process.extractOne for fuzzy matching
    match, score, idx = process.extractOne(row, choices, scorer=fuzz.partial_ratio)
    return match if score >= threshold else None


path = fr"P:\PACS\Finance\Acquistions & New Build\Active\2024 - RealSG AZ-4\Proforma\*Proforma.xlsx"

# for potential_deal in df1['facility_name']:
for potential_deal in glob(path):
    # uncomment when tyler adds it
    # potential_deal_path = df1[df1['file_path'] == potential_deal]
    potential_deal = 'Devon Gables Rehabilitation Center'
    matching_row = df1[df1['facility_name'] == potential_deal]

    # Extract the licensed_op_beds value from that row
    if not matching_row.empty:
        potential_deal_beds = matching_row['licensed_op_beds'].iloc[0]
        print(f"Facility: {potential_deal}, Licensed Beds: {potential_deal_beds}")
    else:
        print(f"Facility: {potential_deal} not found in the DataFrame.")

    potential_deal_path = fr"P:\PACS\Finance\Acquistions & New Build\Active\2024 - RealSG AZ-4\Proforma"
    # File Path
    # if not potential_deal_path.empty:
    #     proforma_path = potential_deal_path['file_path'].iloc[0]
    # else:
    #     print(f"Deal: {potential_deal} no file path.")
    #     continue

    # for example: P:\PACS\Finance\Acquistions & New Build\Active\2024 - RealSG AZ-4\Proforma

    # Extract the file information
    proforma_path = fr"P:\PACS\Finance\Acquistions & New Build\Active\2024 - RealSG AZ-4\Proforma"
    for files in glob(proforma_path):
        proformas_path = fr"{proforma_path}\*Proforma.xlsx"

        for proforma in glob(proformas_path):
            proforma_file = os.path.basename(proforma)
            proforma_file = proforma_file.replace('.xlsx', '')
            facility_name = proforma_file.split(" - ")[1] if len(proforma_file.split(" - ")) > 1 else None
            print(facility_name)

            # Extract the exact details from the excel workbook
            app = xw.App(visible=False)
            xw.Visible = False
            xw.ScreenUpdating = False
            xw.Interactive = False
            wb = xw.Book(proforma, update_links=False)
            in_file_name = wb.sheets("FACILITY INFO").range("B7").value
            in_file_beds = wb.sheets("FACILITY INFO").range("B10").value

            print(fr"Validating Deal:")
            print(fr"Deal: {facility_name} -> Facility {in_file_name}")
            print(potential_deal_beds)
            print(in_file_beds)

            # Match Deal to File Name
            if match_facility_name(potential_deal, facility_name, threshold=75):
                if abs(potential_deal_beds - in_file_beds) < 3:
                    print("Close Match on Beds, Moving Forward to See if Name Matches")

                    # Lower Threshold Based on Bed Count Proximity
                    if match_facility_name(potential_deal, in_file_name, threshold=63):
                        print("Match Found")

                        # ID
                        deal_id = df1[df1['id'] == potential_deal]
                        wb.sheets("DW Upload").range("B3").value = deal_id
                        wb.sheets("DW Upload").range("B4").value = 'default'

                    # If No Match, Flag and Manually Determine
                    else:
                        manual_match = str(
                            input(fr"Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))
                        if manual_match == True:
                            print("Match Found")
                            # ID
                            deal_id = df1[df1['id'] == potential_deal]
                            wb.sheets("DW Upload").range("B3").value = deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                            # wb.save()
                        else:
                            print(fr"***No Possible Match for: {potential_deal}***")
                else:
                    # Match on Deal and In-File Name
                    if match_facility_name(potential_deal, in_file_name, threshold=85):
                        print("Match Found")

                        # ID
                        deal_id = df1[df1['id'] == potential_deal]
                        wb.sheets("DW Upload").range("B3").value = deal_id
                        wb.sheets("DW Upload").range("B4").value = 'default'
                    # Manual Match
                    else:
                        print(fr"***No Possible Match for: {potential_deal}***")

            # If There is a Weak Match on Deal to File Name, Try the In-File Name as an Alternative
            else:
                # Higher Threshold as This is the Backup
                if abs(potential_deal_beds - in_file_beds) < 3:
                    print("Close Match on Beds, Moving Forward to See if Name Matches")
                    # Lower Threshold Based on Bed Count Proximity

                    if match_facility_name(potential_deal, in_file_name, threshold=85):
                        print("Match Found")

                        # ID
                        deal_id = df1[df1['id'] == potential_deal]
                        wb.sheets("DW Upload").range("B3").value = deal_id
                        wb.sheets("DW Upload").range("B4").value = 'default'
                    else:
                        manual_match = str(
                            input(fr"Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))
                        if manual_match == True:
                            print("Match Found")
                            # ID
                            deal_id = df1[df1['id'] == potential_deal]
                            wb.sheets("DW Upload").range("B3").value = deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                        else:
                            print(fr"No Possible Match for {potential_deal}")
                else:
                    if match_facility_name(potential_deal, in_file_name, threshold=85):
                        print("Match Found")

                        # ID
                        deal_id = df1[df1['id'] == potential_deal]
                        wb.sheets("DW Upload").range("B3").value = deal_id
                        wb.sheets("DW Upload").range("B4").value = 'default'
                    else:
                        print(fr"No Possible Match for {potential_deal}")












            #
            #
            #
            #
            # # Start of Algo Match
            # if abs(potential_deal_beds - in_file_beds) < 3:
            #     print("Close Match on Beds, Moving Forward to See if Name Matches")
            #     # Lower Threshold Based on Bed Count Proximity
            #     if match_facility_name(potential_deal, in_file_name, threshold=70):
            #         print("Match Found")
            #
            #         # ID
            #         deal_id = df1[df1['id']==potential_deal]
            #         wb.sheets("DW Upload").range("B3").value = deal_id
            #         wb.sheets("DW Upload").range("B4").value = 'default'
            #
            #     else:
            #         manual_match = str(input(fr"Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))
            #         if manual_match == True:
            #             print("Match Found")
            #             # ID
            #             deal_id = df1[df1['id'] == potential_deal]
            #             wb.sheets("DW Upload").range("B3").value = deal_id
            #             wb.sheets("DW Upload").range("B4").value = 'default'
            # else:
            #     # Increase the Threshold if This is the Last Resort
            #     if match_facility_name(potential_deal, in_file_name, threshold=85):
            #         print("Match Found")
            #
            #         # ID
            #         deal_id = df1[df1['id'] == potential_deal]
            #         wb.sheets("DW Upload").range("B3").value = deal_id
            #         wb.sheets("DW Upload").range("B4").value = 'default'
            #     else:
            #         print(fr"No Possible Match for {potential_deal}")
            #

