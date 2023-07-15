folder_path = "./Data"  # Use the current folder as the source folder

ishares_russell_midcap_etf_url = 'https://www.ishares.com/us/products/239718/ishares-russell-midcap-etf/1521942788811.ajax?fileType=xls&fileName=iShares-Russell-Mid-Cap-ETF_fund&dataType=fund'

file_ishares_init_prefix = "iShares-Russell-Mid-Cap-ETF_fund"
file_ishares_out_prefix = "iShares-Russell-Mid-Cap-ETF_Out"

file_gurufocus_prefix = "iShares-Russell-Mid-Cap-ETF_Guru"
file_gurufocus_attributes = 'GuruFocus_attributes.xlsx'

file_ishares_init_extension = ".xls"
file_ishares_out_extension = ".xlsx"

ishares_holdings_sheet_name = "Holdings"
ishares_out_sheet_name = "iShares-Russell-Mid-Cap-ETF"
gf_sheet_name = "Sheet1"  # Name of the sheet to import
ishares_gurufocus_sheet_name = "GuruFocus_Data"  # New name for the imported sheet
analysis_sheet_name = "Analysis"

six_month_price_column_name = "6M Price"
price_column_name = "Price"
six_month_rps_column_name = '6M RPS'
upside_potencial_column_name = 'Upside Potencial'


import os
from datetime import datetime, date
import pandas as pd

def find_newest_file(folder_path, file_input_prefix, file_input_extension, file_output_prefix, file_out_extension, input_sheet_name):
    # Get a list of files matching the criteria
    files = [
        f for f in os.listdir(folder_path)
        if f.startswith(file_input_prefix) and f.endswith(file_input_extension)
    ]

    # Sort files by modification time in descending order
    files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)

    print(f"iShares files matching criteria: \n{files}\n")
    # Get the newest file
    newest_file = files[0]

    # Construct the full file path
    source_file = os.path.join(folder_path, newest_file)
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

def find_newest_file_simple(folder_path, file_input_prefix, file_input_extension):
    # Get a list of files matching the criteria
    files = [
        f for f in os.listdir(folder_path)
        if f.startswith(file_input_prefix) and f.endswith(file_input_extension)
    ]

    # Sort files by modification time in descending order
    files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)

    print(f"Newest files matching criteria: \n{files}\n")
    # Get the newest file
    newest_file = files[0]

    # Construct the full file path
    source_file = os.path.join(folder_path, newest_file)
    print(f"Newest one:\n{source_file}\n")

    return source_file


import pandas as pd
import csv

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
    df = pd.read_csv(csv_file, skiprows=skiprows)

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
