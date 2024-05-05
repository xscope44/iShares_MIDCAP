# pip3 install progressbar --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org 
# pip3 install requests
# pip3 install bs4
# pip3 install pandas openpyxl
# pip3 install progress progressbar2 alive-progress tqdm --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org

# attributes Excel example:
# +--------+-------------------+----------------+-------------------+
# | Ticker | Piotroski F-Score | Altman Z-Score | Beneish M-Score   |
# +--------+-------------------+----------------+-------------------+

# Download Holdings Data by downloading excel from:
# https://www.blackrock.com/us/individual/products/239719/ishares-russell-midcap-value-etf
# open in excel and save as excel format
# Run read_iShares_excel to generate clean Ticker excel
# Run this GuruFocus_parser to parse data from GuruFocus

# Google does not like to be scraped directly. Instead of a simple requests.get, use a session and a post request to create initial cookies. Then, proceed with scraping.
# Here's an example code snippet:

# import requests
# from bs4 import BeautifulSoup

# with requests.Session() as s:
    # url = "https://www.google.com/search?q=fitness+wear"
    # headers = {
    #     "referer": "referer: https://www.google.com/",
    #     "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
    # }
    # s.post(url, headers=headers)
    # response = s.get(url, headers=headers)
    # soup = BeautifulSoup(response.text, 'html.parser')
    # print(soup)
# This approach allows you to scrape Google search results after handling the consent form1.

import os
import sys
import pandas as pd
import csv
from datetime import date
import config
import requests
import re
from bs4 import BeautifulSoup
from alive_progress import alive_bar

data_folder = config.folder_path  # Use the current folder as the source folder
file_source_prefix = config.file_ishares_out_prefix
file_extension = config.file_ishares_out_extension
file_output_prefix = config.file_gurufocus_prefix
attributes_file = config.file_gurufocus_attributes
ishares_out_sheet_name = config.ishares_out_sheet_name
ticker = config.ticker_column_name

# Generate the current date in the format "yyyymmdd"
current_date = date.today().strftime("%Y%m%d")

# File names
output_file = f"{file_output_prefix}_{current_date}.xlsx"
csv_file = f"{file_output_prefix}_{current_date}.csv"

csv_file_path = os.path.join(data_folder, csv_file)
output_file_path = os.path.join(data_folder, output_file)

def sanitize(s):
    out = s
    # Fill this up with whatever additional meta characters you need to escape
    for meta_char in ['(', ')']:
        out = out.replace(meta_char, '\\'+meta_char)
    return out
source_file = config.find_newest_file_simple(data_folder, file_source_prefix, file_extension)

# Read the symbols and ls from Excel file
df_input = pd.read_excel(attributes_file)
df_symbols = pd.read_excel(source_file)
symbols = df_symbols[ticker].tolist()
ls = df_input.columns.tolist()
print(f"{ls[:5]}...")
print(f"{symbols[:13]}...")

# Check if CSV file exists
try:
    df_output = pd.read_csv(csv_file_path)
except FileNotFoundError:
    # Create an empty DataFrame if the CSV file doesn't exist
    df_output = pd.DataFrame(columns=ls)
    df_output.to_csv(csv_file_path, mode='a', index=False)


# Determine the start index for parsing
start_index = len(df_output)
if start_index > 1:
    print(f"CONTINUE Parsing from index: {start_index}\n")
i = 0
for t in symbols[start_index:]:
    i += 1
print(f"Amount of Stocks to parse: {i}\n\nPlease wait, this may take a while...")

with alive_bar(i, force_tty=True) as bar:
    for t in symbols[start_index:]:
        with requests.Session() as s:
            url = "https://gurufocus.com/stock/" + t + "/summary"
            headers = {
                "referer": "referer: https://www.google.com/",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
            }
            s.post(url, headers=headers)
            response = s.get(url, headers=headers)
            soup = BeautifulSoup(response.content, 'html.parser')
            scores = [t]
        
        for val in ls[1:]:
            val = sanitize(val)
            score1 = soup.find('a', string=re.compile(val))
            if score1 is not None:
                score = soup.find('a', string=re.compile(val)).find_next('td').text.strip()
            else: 
                score = "N/A"
            
            # GF Value
            if val == 'GF Value':
                score2 = soup.select('h2 > a', class_="t-h6", string=re.compile(val))
                i=0
                for xt in score2:
                    gf_value = xt.text.strip()
                    # print(f"GF Value gf_value: {gf_value}!")
                    if gf_value.find(val) != -1:
                        gf_value2 = gf_value.split('\n')
                        # print(f"GF Value gf_value2: {gf_value2}!")
                        i+=1
                if i>0:
                    gf_value = [x.replace(' ','') for x in gf_value2]
                    score = gf_value[1].replace('$','')
                    score = float(score)
                    
                else:
                    score = 0
                # print(f"GF Value score: {score}!")     
              
            scores.append(score)
            
        # print(len(df_output))
        df_len = len(df_output)
        # print(df_output['GF Value'].loc[df_len-1])
        # gfvalue1 = df_output['GF Value'].loc[df_len-1]
        
        # df_output.loc[len(df_output)] = scores
        
        # df_len = len(df_output)
        # gfvalue2 = df_output['GF Value'].loc[df_len-1]
      
        # if gfvalue1 == gf_value2:
        #     print('Warning!')
        #     print(df_output['Ticker'].loc[df_len-1])
        #     print(df_output['GF Value'].loc[df_len-1]) 
        df_output['GF Value'].diff().eq(0)
        with open(csv_file_path, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(scores)
        bar()

print("\nProcessing is complete.")

# Export the scraped data to Excel file
df_output.to_excel(output_file_path, index=False)
print(f"Data were saved to file: {output_file_path}\n")