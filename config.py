data_folder_path = ".\Data"  # Use Data folder to download and create excel sheets and csv

ishares_russell_midcap_etf_download_url = 'https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf/1515394931018.ajax?fileType=xls&fileName=iShares-Russell-Mid-Cap-ETF_fund&dataType=fund'
ishares_russell_midcap_value_etf_download_url = 'https://www.ishares.com/ch/professionals/en/products/239719/ishares-russell-midcap-value-etf/1535604580403.ajax?fileType=xls&fileName=iShares-Russell-Mid-Cap-Value-ETF_fund&dataType=fund'
ishares_fund_prefix = "iShares-Russell-Mid-Cap-Value-ETF_fund"
ishares_out_prefix = "iShares-Russell-Mid-Cap-Value-ETF_Out"
gurufocus_prefix = "iShares-Russell-Mid-Cap-Value-ETF_Guru"
gurufocus_attributes_xlsx = 'GuruFocus_attributes.xlsx'

ishares_manual_fund_csv = 'iShares-Russell-Mid-Cap-ETF_fund_20240428.csv'
manual_download = False

file_ishares_init_extension = ".xls"
file_ishares_out_extension = ".xlsx"

ishares_holdings_sheet_name = "Holdings"
ishares_out_sheet_name = "Mid-Cap-Value-ETF"
gf_sheet_name = "Sheet1"  # Name of the sheet to import
ishares_gurufocus_sheet_name = "GuruFocus_Data"  # New name for the imported sheet
analysis_sheet_name = "Analysis"

six_month_price_column_name = "6M Price"
price_column_name = "Price"
six_month_rps_column_name = '6M RPS'
upside_potencial_column_name = 'Upside Potencial'
ticker_column_name = 'Issuer Ticker' #midcap value etf: 'Issuer Ticker', Others: 'Ticker'

import os
from datetime import datetime, date
import pandas as pd
import csv

# Define print colors functions:
def prRed(skk): print("\033[91m {}\033[00m" .format(skk))
 
 
def prGreen(skk): print("\033[92m {}\033[00m" .format(skk))
 
 
def prYellow(skk): print("\033[93m {}\033[00m" .format(skk))
 
 
def prLightPurple(skk): print("\033[94m {}\033[00m" .format(skk))
 
 
def prPurple(skk): print("\033[95m {}\033[00m" .format(skk))
 
 
def prCyan(skk): print("\033[96m {}\033[00m" .format(skk))
 
 
def prLightGray(skk): print("\033[97m {}\033[00m" .format(skk))
 
 
def prBlack(skk): print("\033[98m {}\033[00m" .format(skk))

def character_count(text, start_index):
  """
  This function calculates the number of characters in a string after slicing it from a specific starting point.

  Args:
    text: The string you want to analyze.
    start_index: An integer representing the index from where the slicing should start.

  Returns:
    An integer representing the number of characters in the sliced string.
  """

  if start_index < 0:
    raise ValueError("Index cannot be negative.")

  if start_index >= len(text):
    return 0

  return len(text[start_index:])

# Generate filenames and relative or full path to the data files: ishares_fund_csv_path, ishares_out_xlsx_path, guru_csv_path, guru_xlsx_path
def create_file_paths():
    # Generate the current date in the format "yyyymmdd"
    current_date = date.today().strftime("%Y%m%d")

     # Construct output file names
    output_xlsx = f"{ishares_out_prefix}_{current_date}.xlsx"
    fund_csv = f"{ishares_fund_prefix}_{current_date}.csv"
    output_file = f"{gurufocus_prefix}_{current_date}.xlsx"
    csv_file = f"{gurufocus_prefix}_{current_date}.csv"
    
    # Join the folder path with the file names
    ishares_fund_csv_path = os.path.join(data_folder_path, fund_csv)
    ishares_out_xlsx_path = os.path.join(data_folder_path, output_xlsx)
    
    #Gurufocus
    guru_csv_path = os.path.join(data_folder_path, csv_file)
    guru_xlsx_path = os.path.join(data_folder_path, output_file)
    
    # Create the directory if it doesn't exist
    if not os.path.exists(data_folder_path):
        os.makedirs(data_folder_path)
        print(f"Folder {data_folder_path} created successfully!")

    return ishares_fund_csv_path, ishares_out_xlsx_path, guru_csv_path, guru_xlsx_path


def find_newest_file(data_folder_path, file_input_prefix, file_input_extension, file_output_prefix, file_out_extension, input_sheet_name):
    # Get a list of files matching the criteria
    files = [
        f for f in os.listdir(data_folder_path)
        if f.startswith(file_input_prefix) and f.endswith(file_input_extension)
    ]

    # Sort files by modification time in descending order
    files.sort(key=lambda f: os.path.getmtime(os.path.join(data_folder_path, f)), reverse=True)

    #prCyan(f"iShares files matching criteria: \n{files}\n")
    # Get the newest file
    newest_file = files[0]

    # Construct the full file path
    source_file = os.path.join(data_folder_path, newest_file)
    print(f"{source_file} was identified as the newest one to parse Holdings sheet from")

    # Generate the current date in the format "yyyymmdd"
    current_date = date.today().strftime("%Y%m%d")

    # Read the first cell (A1) from the "Holdings" sheet to check for a date
    source_df = pd.read_excel(source_file, sheet_name=input_sheet_name, header=None, nrows=1)
    date_cell = source_df.iloc[0, 0]

    # Try to parse the date in "03-Jul-2023" format
    try:
        parsed_date = datetime.strptime(date_cell, "%d-%b-%Y")
        current_date = parsed_date.strftime("%Y%m%d")
    except ValueError:
        # If parsing fails, use the current date
        current_date = date.today().strftime("%Y%m%d")

    output_file = f"{file_output_prefix}_{current_date}{file_out_extension}"

    return output_file, current_date, source_file

def find_newest_file_simple(data_folder_path, file_input_prefix, file_input_extension):
    # Get a list of files matching the criteria
    files = [
        f for f in os.listdir(data_folder_path)
        if f.startswith(file_input_prefix) and f.endswith(file_input_extension)
    ]

    # Sort files by modification time in descending order
    files.sort(key=lambda f: os.path.getmtime(os.path.join(data_folder_path, f)), reverse=True)

    #prCyan(f"Newest files matching criteria: \n")
    #print(f"{files}\n")
    # Get the newest file
    newest_file = files[0]

    # Construct the full file path
    source_file = os.path.join(data_folder_path, newest_file)
    prCyan(f"Newest one:")
    print(f"{source_file}\n")

    return source_file


#  automatically adjusts the skiprows parameter based on the first occurrence of the keyword "Ticker" in the CSV file
def convert_csv_to_excel(csv_file, excel_file, sheet_name='Sheet1'):
    # Read the CSV file to determine the skiprows value and the skipped rows
    skipped_rows = []
    with open(csv_file, 'r') as file:
        reader = csv.reader(file)
        skiprows = 0
        for row in reader:
            if 'Ticker' in row:
                break
            skipped_rows.append(row)
            skiprows += 1
    
    # Read the CSV file into a pandas DataFrame without skipped rows
    df = pd.read_csv(csv_file,  delim_whitespace=True)

    # Create an Excel writer object
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')

    # Write the DataFrame to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the skipped rows to an "Info" sheet
    info_sheet_name = 'Info'
    skipped_df = pd.DataFrame(skipped_rows)
    skipped_df.to_excel(writer, sheet_name=info_sheet_name, index=False)

    # Save the changes and close the writer
    writer._save()
