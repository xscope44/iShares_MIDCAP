import pandas as pd
import os
from datetime import date, datetime
from openpyxl import load_workbook
import config
from simplified_scrapy import SimplifiedDoc, utils, req

# Download excel data of iShares-Russell-Mid-Cap-ETF_fund to csv
# # https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf
# convert csv to excel 
# Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out sheet name iShares-Russell-Mid-Cap-ETF

folder_path = config.folder_path
file_input_prefix = config.file_ishares_init_prefix
file_output_prefix = config.file_ishares_out_prefix
file_input_extension = config.file_ishares_init_extension
file_out_extension = config.file_ishares_out_extension
input_sheet_name = config.ishares_holdings_sheet_name
output_sheet_name = config.ishares_out_sheet_name
ishares_russell_midcap_etf_url = config.ishares_russell_midcap_etf_url

# Generate the current date in the format "yyyymmdd"
current_date = date.today().strftime("%Y%m%d")

# source_file = "iShares-Russell-Mid-Cap-ETF_fund_20230703.xlsx"
output_file = f"{file_output_prefix}_{current_date}.xlsx"
csv_file = f"{file_input_prefix}_{current_date}.csv"

# Join the folder path with the file names
csv_file_path = os.path.join(folder_path, csv_file)
output_file_path = os.path.join(folder_path, output_file)

if not os.path.exists(folder_path):
    # Create the directory
    os.makedirs(folder_path)
    print(f"Folder {folder_path} created successfully!")
else:
    print(f"Folder {folder_path} already exists!")


print(f"Downloading {file_input_prefix} Holdings Data from:\n{ishares_russell_midcap_etf_url}\nPlease wait, it may take a while...")
xml = req.get(
    # 'https://www.ishares.com/us/products/239722/ishares-russell-top-200-value-etf/1521942788811.ajax?fileType=xls&fileName=iShares-Russell-Top-200-Value-ETF_fund&dataType=fund'
    ishares_russell_midcap_etf_url
)   
xml = xml.read().decode('utf-8')
doc = SimplifiedDoc(xml)
worksheets = doc.selects('ss:Worksheet') # Get all Worksheets
for worksheet in worksheets:
    if worksheet['ss:Name'] == 'Holdings': 
        rows = worksheet.selects('ss:Row').selects('ss:Cell>text()') # Get all rows
        utils.save2csv(csv_file_path, rows) # Save Holdings sheet data to csv


# Convert Holding sheet csv to excel
config.convert_csv_to_excel(csv_file_path, output_file_path, output_sheet_name)

print(f"\nDownloaded {csv_file} has been successfully converted to {output_file}\n")
print(csv_file)
print(output_file)
