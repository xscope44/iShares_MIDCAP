import pandas as pd
import os
from openpyxl import load_workbook
import config
from simplified_scrapy import SimplifiedDoc, utils, req

# Download excel data iShares Russell Mid-Cap Value ETF to csv
# https://www.blackrock.com/us/individual/products/239719/ishares-russell-midcap-value-etf
# Convert csv to excel 
# Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out sheet name iShares-Russell-Mid-Cap-ETF

# Generate filenames with current date in the format "yyyymmdd"
fund_csv_path, output_file_path,c,d = config.create_file_paths()

if not os.path.exists(config.data_folder_path):
    # Create the directory
    os.makedirs(config.data_folder_path)
    print(f"Folder {config.data_folder_path} created successfully!")

if not config.manual_download:
    print(f"Downloading {config.ishares_fund_prefix} Holdings Data.\nPlease wait, it may take a while...\n")
    xml = req.get(
        config.ishares_russell_midcap_value_etf_download_url
    )

    xml = xml.read().decode('utf-8')
    doc = SimplifiedDoc(xml)
    worksheets = doc.selects('ss:Worksheet') # Get all Worksheets
    for worksheet in worksheets:
        if worksheet['ss:Name'] == 'Holdings': 
            rows = worksheet.selects('ss:Row').selects('ss:Cell>text()') # Get all rows
            utils.save2csv(fund_csv_path, rows) # Save Holdings sheet data to csv
else: 
    fund_csv_path = config.ishares_manual_fund_csv
print(f"{fund_csv_path} has been saved successfully!\n")

# Convert Holding sheet csv to excel
read_file = pd.read_csv(fund_csv_path, delimiter=',', encoding = 'utf-8', on_bad_lines = 'skip', low_memory = False, skiprows=13)
#assign dataframe
read_file.head()
read_file.to_excel(output_file_path,sheet_name=config.ishares_out_sheet_name ,index=None, header=True)

# config.convert_csv_to_excel(fund_csv_path, output_file_path, config.ishares_out_sheet_name)
# print(f"Downloaded {fund_csv} has been successfully converted to {output_xlsx}\n")
print(fund_csv_path)
print(output_file_path)
