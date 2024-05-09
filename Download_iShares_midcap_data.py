import pandas as pd
import os
from openpyxl import load_workbook
import config as c
from simplified_scrapy import SimplifiedDoc, utils, req

# Download excel data iShares Russell Mid-Cap Value ETF to csv
# https://www.blackrock.com/us/individual/products/239719/ishares-russell-midcap-value-etf
# Convert csv to excel 
# Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out sheet name iShares-Russell-Mid-Cap-ETF

c.prCyan(f"******** Download {c.ishares_fund_prefix} Data from blackrock website ********")
# Generate filenames with current date in the format "yyyymmdd"
fund_csv_path, output_file_path, guru_csv_path, guru_xlsx_path = c.create_file_paths()

if not os.path.exists(c.data_folder_path):
    # Create the directory
    os.makedirs(c.data_folder_path)
    c.prCyan(f"Folder {c.data_folder_path} created successfully!")

if not c.manual_download:
    c.prCyan(f"Downloading {c.ishares_fund_prefix} Holdings Data:\n")
    xml = req.get(
        c.ishares_russell_midcap_value_etf_download_url
    )

    xml = xml.read().decode('utf-8')
    doc = SimplifiedDoc(xml)
    worksheets = doc.selects('ss:Worksheet') # Get all Worksheets
    for worksheet in worksheets:
        if worksheet['ss:Name'] == 'Holdings': 
            rows = worksheet.selects('ss:Row').selects('ss:Cell>text()') # Get all rows
            utils.save2csv(fund_csv_path, rows) # Save Holdings sheet data to csv
    print (f"{fund_csv_path} has been saved successfully!\n")
else:
     # Join the folder path with the file names
    ishares_fund_csv_path = os.path.join(c.data_folder_path, c.ishares_manual_fund_csv)
    try:
        with open(ishares_fund_csv_path, 'r') as f:
            fund_csv_path = ishares_fund_csv_path
            print (f"{ishares_fund_csv_path} has been loaded successfully!\n")

    except FileNotFoundError:
        c.prRed(f"{ishares_fund_csv_path} does not exist.")
        c.prRed("Make sure the file exists and is defined correctly in config file")
        exit
# Convert Holding sheet csv to excel
c.prCyan ("Converted csv to excel:\n")
read_file = pd.read_csv(fund_csv_path, delimiter=',', encoding = 'utf-8', on_bad_lines = 'skip', low_memory = False, skiprows=13)
#assign dataframe
read_file.head()
read_file.to_excel(output_file_path,sheet_name=c.ishares_out_sheet_name ,index=None, header=True)

# c.convert_csv_to_excel(fund_csv_path, output_file_path, c.ishares_out_sheet_name)
# c.prCyan(f"Downloaded {fund_csv} has been successfully converted to {output_xlsx}\n")
print(output_file_path)
exit