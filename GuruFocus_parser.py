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
# https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf
# open in excel and save as excel format
# Run read_iShares_excel to generate clean Ticker excel
# Run this GuruFocus_parser to parse data from GuruFocus

# import pandas as pd
# import requests
# from bs4 import BeautifulSoup
# import re
# import os
# import csv
# from datetime import date, datetime
# from alive_progress import alive_bar
# import time
# import config

import sys
import pandas as pd
import csv
from datetime import date
import config
import requests
import re
from bs4 import BeautifulSoup
from alive_progress import alive_bar

folder_path = config.folder_path  # Use the current folder as the source folder
file_source_prefix = config.file_ishares_out_prefix
file_extension = config.file_ishares_out_extension
file_output_prefix = config.file_gurufocus_prefix
attributes_file = config.file_gurufocus_attributes
ishares_out_sheet_name = config.ishares_out_sheet_name

# Generate the current date in the format "yyyymmdd"
current_date = date.today().strftime("%Y%m%d")

# File names
output_file = f"{file_output_prefix}_{current_date}.xlsx"
csv_file = f"{file_output_prefix}_{current_date}.csv"

def sanitize(s):
    out = s
    # Fill this up with whatever additional meta characters you need to escape
    for meta_char in ['(', ')']:
        out = out.replace(meta_char, '\\'+meta_char)
    return out
source_file = config.find_newest_file_simple(folder_path, file_source_prefix, file_extension)

# Read the symbols and ls from Excel file
df_input = pd.read_excel(attributes_file)
df_symbols = pd.read_excel(source_file)
symbols = df_symbols['Ticker'].tolist()
ls = df_input.columns.tolist()
print(f"{ls[:5]}...")
print(f"{symbols[:13]}...")

# Check if CSV file exists
try:
    df_output = pd.read_csv(csv_file)
except FileNotFoundError:
    # Create an empty DataFrame if the CSV file doesn't exist
    df_output = pd.DataFrame(columns=ls)
    df_output.to_csv(csv_file, mode='a', index=False)


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
        req = requests.get("https://www.gurufocus.com/stock/" + t)
        if req.status_code != 200:
            continue
        soup = BeautifulSoup(req.content, 'html.parser')
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
                for xt in score2:
                    gf_value = xt.text.strip()
                    if gf_value.find(val) != -1:
                        gf_value2 = gf_value.split('\n')
                gf_value = [x.replace(' ','') for x in gf_value2]
                score = gf_value[1].replace('$','')

            scores.append(score)
        
        df_output.loc[len(df_output)] = scores
        with open(csv_file, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(scores)
        bar()

print("\nProcessing is complete.")

# Export the scraped data to Excel file
df_output.to_excel(output_file, index=False)
print(f"Data were saved to file: {output_file}\n")
