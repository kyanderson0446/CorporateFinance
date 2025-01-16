# Data and DB Login
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
from deal_query import DatabaseFetcher
import json
import duckdb
# File Management
import time
import os
from glob import glob
# Excel
import xlwings as xw
# Algo
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
print("Data fetching complete")

# Configure logging to write to a log file
logging.basicConfig(
    filename="no_match_log.txt",  # Name of the log file
    level=logging.INFO,           # Logging level
    format="%(asctime)s - %(message)s",  # Log format
    datefmt="%Y-%m-%d %H:%M:%S"         # Date and time format
)

def log_no_match(potential_deal_id):
    logging.info(f"No Possible Match for: {potential_deal_id}")

# Comparison Matchers Here
def compare_matchers(name1, name2, threshold=85):
    """
    Compare different text matching algorithms for facility name matching.

    Args:
        name1 (str): First facility name
        name2 (str): Second facility name
        threshold (int): Score threshold for considering a match

    Returns:
        dict: Scores from different matching algorithms
    """

    def normalize_name(name):
        # Convert to lowercase and remove common words/punctuation
        name = name.lower()
        name = re.sub(r'\b(center|centre|rehabilitation|rehab|facility)\b', '', name)
        name = re.sub(r'[^\w\s]', '', name)
        return ' '.join(name.split())

    # Normalize names
    norm1 = normalize_name(name1)
    norm2 = normalize_name(name2)

    return {
        # Token-based comparisons (better for word rearrangement)
        "token_sort_ratio": fuzz.token_sort_ratio(norm1, norm2),
        "token_set_ratio": fuzz.token_set_ratio(norm1, norm2),

        # Sequence-based comparisons (better for typos/spelling variations)
        "levenshtein_distance": jellyfish.levenshtein_distance(norm1, norm2),
        "jaro_winkler": int(jellyfish.jaro_winkler_similarity(norm1, norm2) * 100),

        # Hybrid approaches
        "sequence_matcher": int(SequenceMatcher(None, norm1, norm2).ratio() * 100),
        "metaphone": jellyfish.metaphone(norm1) == jellyfish.metaphone(norm2),
    }

# Matching Engine Begins Here
def match_facility_name(
        row: str,
        choices: List[str],
        threshold: int = 85,
        debug: bool = False
) -> Tuple[Optional[str], float]:
    """
    Enhanced facility name matcher using multiple algorithms.

    Args:
        row: The facility name to match
        choices: List of possible facility names to match against
        threshold: Minimum score to consider a match
        debug: If True, print detailed matching information

    Returns:
        tuple: (best_match, confidence_score) or (None, 0) if no match
    """
    # Input validation
    if not isinstance(choices, (list, tuple)):
        raise ValueError(f"choices must be a list or tuple, got {type(choices)}")

    if not choices:
        return None, 0

    if not isinstance(row, str) or len(row.strip()) == 0:
        return None, 0

    # Filter out any non-string or empty choices
    valid_choices = [c for c in choices if isinstance(c, str) and len(c.strip()) > 0]

    if debug:
        print(f"\nMatching facility: {row}")
        print(f"Number of valid choices: {len(valid_choices)}")

    best_match = None
    best_score = 0

    for choice in valid_choices:
        if len(choice) <= 1:  # Skip single characters
            continue

        scores = compare_matchers(row, choice, threshold)

        # Calculate composite score
        token_score = max(scores['token_sort_ratio'], scores['token_set_ratio'])
        sequence_score = scores['jaro_winkler']

        # Weighted average (emphasize token-based matching)
        composite_score = (token_score * 0.7 + sequence_score * 0.3)

        if scores['metaphone']:
            composite_score = min(100, composite_score + (10 if scores['metaphone'] else 0))

        if debug:
            print(f"\nComparing with: {choice}")
            print(f"Token score: {token_score}")
            print(f"Sequence score: {sequence_score}")
            print(f"Composite score: {composite_score}")
            print(f"All scores: {scores}")

        # Early rejection for clearly different names
        if scores['levenshtein_distance'] > min(len(row), len(choice)) / 2:
            if debug:
                print(f"Rejected due to high Levenshtein distance: {scores['levenshtein_distance']}")
            continue

        if composite_score > best_score and composite_score >= threshold:
            best_score = composite_score
            best_match = choice

    if debug:
        print(f"\nBest match: {best_match}")
        print(f"Best score: {best_score}")

    return best_match, best_score


def is_duplicate_deal(df_log, in_file_beds, file_facility_name, proforma):
    if df_log.empty:
        return False

    match_condition = (
            (df_log['Beds'] == in_file_beds) &
            (df_log['File Name'] == file_facility_name) &
            (df_log['File Path'] == proforma)
    )

    return match_condition.any()


def matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma):
    new_row = {
        "Deal ID": deal_id,
        "Beds": in_file_beds,
        "File Name": file_facility_name,
        "File Path": proforma
    }
    # Create Dataframe for each Deal Folder
    df_log = pd.concat([df_log, pd.DataFrame([new_row])], ignore_index=True)
    return df_log  # Return the updated dataframe


# Log files that have already been matched
df_log = pd.DataFrame(columns=["Deal ID", "Beds", "File Name", "File Path"])


for potential_deal in df1['facility_name']: # Loop through all deals as there will be duplicate names but unique deals tied to them
    deal_data = df1[df1['facility_name'] == potential_deal]
    potential_deal_path = deal_data['file_path'].iloc[0]
    potential_deal_id = deal_data['id'].iloc[0]
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
                in_file_beds = df_file.iloc[9, 1]
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
                match_1 = match_facility_name(potential_deal, files_in_folder, threshold=98)
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
                        match_5 = match_facility_name(potential_deal, file_name_choices, threshold=95)
                        if match_5[0] is not None:  # Check explicitly if match is not None
                            in_file_name_match = match_5[0]
                            # If there a duplicate In File Names from the files
                            in_file_index = df_excel[df_excel['In File Name'] == in_file_name_match].index[0]
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {matched_value}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                            else:
                                continue
                            # ID
                            locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                            open_file = locate_path['File Path'].iloc[0]
                            app = xw.App(visible=False)
                            xw.Visible = False
                            xw.ScreenUpdating = False
                            xw.Interactive = False
                            wb = xw.Book(open_file, update_links=False)
                            locate_deal_id = df1[df1['facility_name'] == potential_deal]
                            deal_id = locate_deal_id['id'].iloc[0]

                            wb.sheets("DW Upload").range("B3").value = deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                            # wb.save()
                            wb.close()
                            app.quit()
                            # create a matches log
                            df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)

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
                                locate_deal_id = df1[df1['facility_name'] == potential_deal]
                                deal_id = locate_deal_id.iloc[0]
                                wb.sheets("DW Upload").range("B3").value = deal_id
                                wb.sheets("DW Upload").range("B4").value = 'default'
                                # wb.save()
                                wb.close()
                                app.quit()
                                df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)
                                break
                            else:
                                print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")

                    # This is When the File Name was a Close Match but the Bed Count Fails
                    # Increase the Threshold Here as it Requires Matching on File Name and In_File Name
                    else:
                        """
                        3. Proforma -> File Name -> In File Name. Beds were not close but try to match
                        """
                        file_name_choices = temp_df['In File Name'].to_list()
                        match_4 = match_facility_name(potential_deal, file_name_choices, threshold=95)
                        if match_4[0] is not None:  # Check explicitly if match is not None
                            in_file_name_match = match_4[0]
                            index = df_excel[df_excel['File Name'] == in_file_name_match].index[0]
                            in_file_index = df_excel[df_excel['File Name'] == in_file_name_match].index[0]
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {matched_value}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                            else:
                                continue
                            # ID
                            locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                            open_file = locate_path['File Path'].iloc[0]
                            app = xw.App(visible=False)
                            xw.Visible = False
                            xw.ScreenUpdating = False
                            xw.Interactive = False
                            wb = xw.Book(open_file, update_links=False)
                            locate_deal_id = df1[df1['facility_name'] == potential_deal]
                            deal_id = locate_deal_id.iloc[0]
                            wb.sheets("DW Upload").range("B3").value = deal_id
                            wb.sheets("DW Upload").range("B4").value = 'default'
                            # wb.save()
                            wb.close()
                            app.quit()
                            df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)
                            break
                        else:
                            print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")

                # If There is a Weak Match on Deal to File Name, Try the In-File Name as an Alternative
                else:
                    """
                    2. Proforma -> In File name
                    """
                    file_name_choices = df_excel['In File Name'].to_list()
                    match_2 = match_facility_name(potential_deal, file_name_choices, threshold=85)
                    if match_2[0] is not None:  # Check explicitly if match is not None
                        matched_value = match_2[0]
                        temp_df = df_excel[df_excel['In File Name'] == matched_value]
                        file_beds = float(temp_df['Beds'].iloc[0])
                        index = df_excel[df_excel['In File Name'] == matched_value].index[0]

                        # Higher Threshold as This is the Backup
                        """
                        2. Proforma -> In File Name -> Bed
                        """
                        if abs(potential_deal_beds - in_file_beds) < 6:
                            file_name_choices = temp_df['In File Name'].to_list()
                            if index == in_file_index:
                                print(
                                    f"{bcolors.OKGREEN}Found Match:{bcolors.OKBLUE} {matched_value}{bcolors.OKGREEN}, {bcolors.OKBLUE}{file_beds}")
                                # ID
                                locate_path = df_excel[df_excel['In File Name'] == in_file_name_match]
                                open_file = locate_path['File Path'].iloc[0]
                                app = xw.App(visible=False)
                                xw.Visible = False
                                xw.ScreenUpdating = False
                                xw.Interactive = False
                                wb = xw.Book(open_file, update_links=False)
                                locate_deal_id = df1[df1['facility_name'] == potential_deal]
                                deal_id = locate_deal_id.iloc[0]
                                wb.sheets("DW Upload").range("B3").value = deal_id
                                wb.sheets("DW Upload").range("B4").value = 'default'
                                # wb.save()
                                wb.close()
                                app.quit()
                                df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)
                                break
                            else:
                                """
                                2. Proforma -> In File Name -> Bed -> Manual
                                """
                                manual_match = str(
                                    input(
                                        fr"{bcolors.WARNING}Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))
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
                                    locate_deal_id = df1[df1['facility_name'] == potential_deal]
                                    deal_id = locate_deal_id.iloc[0]
                                    wb.sheets("DW Upload").range("B3").value = deal_id
                                    wb.sheets("DW Upload").range("B4").value = 'default'
                                    # wb.save()
                                    wb.close()
                                    app.quit()
                                    df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)
                                    break
                                else:
                                    print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")

                        else:
                            """
                            No Match on File Name, In File Name or Beds. Try Manual
                            """
                            manual_match = str(
                                input(
                                    fr"{bcolors.WARNING}Deal:{potential_deal} -> {in_file_name}, is it a match? True or False"))
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
                                locate_deal_id = df1[df1['facility_name'] == potential_deal]
                                deal_id = locate_deal_id.iloc[0]
                                wb.sheets("DW Upload").range("B3").value = deal_id
                                wb.sheets("DW Upload").range("B4").value = 'default'
                                # wb.save()
                                wb.close()
                                app.quit()
                                df_log = matches_log(df_log, deal_id, in_file_beds, file_facility_name, proforma)
                                break
                            else:
                                print(f"{bcolors.FAIL}***{potential_deal} Did NOT Match: {file_facility_name}***")

    except:
        print(f"{bcolors.FAIL}***No Possible Match for: {potential_deal_id}***")
        log_no_match(potential_deal_id)
