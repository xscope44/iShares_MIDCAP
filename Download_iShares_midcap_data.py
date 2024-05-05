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

data_folder = config.folder_path
file_input_prefix = config.file_ishares_init_prefix
file_output_prefix = config.file_ishares_out_prefix
file_input_extension = config.file_ishares_init_extension
file_out_extension = config.file_ishares_out_extension
input_sheet_name = config.ishares_holdings_sheet_name
output_sheet_name = config.ishares_out_sheet_name
ishares_russell_midcap_etf_url = config.ishares_russell_midcap_value_etf_url
csv_file_manual = 'iShares-Russell-Mid-Cap-ETF_fund_20240428.csv'
manual_download = True

# Generate the current date in the format "yyyymmdd"
current_date = date.today().strftime("%Y%m%d")

# source_file = "iShares-Russell-Mid-Cap-ETF_fund_20230703.xlsx"
output_file = f"{file_output_prefix}_{current_date}.xlsx"
csv_file = f"{file_input_prefix}_{current_date}.csv"

# Join the folder path with the file names
csv_file_path = os.path.join(data_folder, csv_file)
output_file_path = os.path.join(data_folder, output_file)

if not os.path.exists(data_folder):
    # Create the directory
    os.makedirs(data_folder)
    print(f"Folder {data_folder} created successfully!")
else:
    print(f"Folder {data_folder} already exists!")

print(f"csv_file_path: {csv_file_path}\noutput_file_path:{output_file_path}\n")

if not manual_download:
    print(f"Downloading {file_input_prefix} Holdings Data from:\n{ishares_russell_midcap_etf_url}\nPlease wait, it may take a while...\n")
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
else: 
    csv_file_path = csv_file_manual
print(f"csv file: {csv_file_path} has been saved successfully!\n")

print(f"Converting {csv_file_path} to {output_file_path} \n")

# Convert Holding sheet csv to excel
read_file = pd.read_csv(csv_file_path, delimiter=',', encoding = 'utf-8', on_bad_lines = 'skip', low_memory = False, skiprows=13)
#assign dataframe
read_file.head()
read_file.to_excel(output_file_path,sheet_name=output_sheet_name ,index=None, header=True)

# config.convert_csv_to_excel(csv_file_path, output_file_path, output_sheet_name)
print(f"\nDownloaded {csv_file} has been successfully converted to {output_file}\n")
print(csv_file)
print(output_file)
