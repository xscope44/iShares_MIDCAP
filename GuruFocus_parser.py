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

# with requests.Session() as s:
    # url = "https://www.google.com/search?q=fitness+wear"
    # headers = {
    #     "referer": "referer: https://www.google.com/",
    #     "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
    # }
    # s.post(url, headers=headers)
    # response = s.get(url, headers=headers)
    # soup = BeautifulSoup(response.text, 'html.parser')
    # c.prCyan(soup)
# This approach allows you to scrape Google search results after handling the consent form1.
import pandas as pd
import csv
import config as c
import requests
import re
from bs4 import BeautifulSoup
from alive_progress import alive_bar

c.prCyan(f"******** Gurufocus parser of Tickers from: {c.ishares_fund_prefix} ********")
# ls_special=['Debt-to-EBITDA','Gross Margin %','PE Ratio','Financial Strength','Profitability Rank']
ls_special=['Cash-To-Debt', 'Equity-to-Asset', 'Debt-to-Equity', 'Debt-to-EBITDA', 'Interest Coverage', 'Piotroski F-Score', 'Altman Z-Score', 'Beneish M-Score', 'WACC vs ROIC', '3-Year Revenue Growth Rate', '3-Year Book Growth Rate', 'Future 3-5Y EPS without NRI Growth Rate', 'Future 3-5Y Total Revenue Growth Rate', '5-Day RSI', '9-Day RSI', '14-Day RSI', '6-1 Month Momentum %', '12-1 Month Momentum %', 'Current Ratio', 'Quick Ratio', 'Cash Ratio', 'Days Inventory', 'Days Sales Outstanding', 'Days Payable', 'Dividend Yield %', 'Dividend Payout Ratio', '3-Year Dividend Growth Rate', 'Forward Dividend Yield %', '5-Year Yield-on-Cost %', '3-Year Average Share Buyback Ratio', 'Shareholder Yield %', 'Gross Margin %', 'Operating Margin %', 'Net Margin %', 'FCF Margin %', 'ROE %', 'ROA %', 'ROIC %', 'ROC (Joel Greenblatt) %', 'ROCE %', 'Years of Profitability over Past 10-Year', 'PE Ratio', 'Forward PE Ratio', 'PE Ratio without NRI', 'Shiller PE Ratio', 'Price-to-Owner-Earnings', 'PS Ratio', 'PB Ratio', 'Price-to-Tangible-Book', 'Price-to-Free-Cash-Flow', 'Price-to-Operating-Cash-Flow', 'EV-to-EBIT', 'EV-to-Forward-EBIT', 'EV-to-EBITDA', 'EV-to-Forward-EBITDA', 'EV-to-Revenue', 'EV-to-Forward-Revenue', 'EV-to-FCF', 'Price-to-Projected-FCF', 'Price-to-Median-PS-Value', 'Price-to-Graham-Number', 'Earnings Yield (Greenblatt) %', 'FCF Yield %', 'Forward Rate of Return (Yacktman) %', 'Financial Strength', 'Growth Rank', 'Momentum Rank', 'Profitability Rank', 'GF Value Rank']

# Create file paths
ishares_fund_csv_path, ishares_out_xlsx_path, guru_csv_path, guru_xlsx_path = c.create_file_paths()

def sanitize(s):
    # Escape meta characters
    out = s
    for meta_char in ['(', ')']:
        out = out.replace(meta_char, '\\'+meta_char)
    return out

# Get newest ishares out excel file
ishares_out_excel_file = c.find_newest_file_simple(c.data_folder_path, c.ishares_out_prefix, c.file_ishares_out_extension)

# Read symbols and columns from Excel files
df_input = pd.read_excel(c.gurufocus_attributes_xlsx)
df_symbols = pd.read_excel(ishares_out_excel_file)
symbols = df_symbols[c.ticker_column_name].tolist()
ls = df_input.columns.tolist()
print(f"{ls[:5]}...")
print(f"{symbols[:13]}...")

# Check if CSV file exists
try:
    df_output = pd.read_csv(guru_csv_path)
except FileNotFoundError:
    # Create an empty DataFrame if the CSV file doesn't exist
    df_output = pd.DataFrame(columns=ls)
    df_output.to_csv(guru_csv_path, mode='a', index=False)

# Determine the start index for parsing
start_index = len(df_output)
if start_index > 1:
    print(f"Parsing from index: {start_index}")

number_of_symbols = c.character_count(symbols, start_index)

# Print a message indicating the number of stocks to parse
c.prCyan(f"Amount of Stocks to parse: {number_of_symbols}\n\n")
print("Please wait, this may take a while...")

# Create a progress bar using the 'alive_bar' library
with alive_bar(number_of_symbols, force_tty=True) as bar:
    # Iterate over each stock symbol in the 'symbols' list
    for t in symbols[start_index:]:
        # Set up an HTTP session using the 'requests' library
        with requests.Session() as s:
            # Construct the URL for the stock summary page
            url = "https://gurufocus.com/stock/" + t + "/summary"
            headers = {
                "referer": "referer: https://www.google.com/",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
            }
            # Send a POST request to the URL
            s.post(url, headers=headers)
            # Send a GET request to retrieve the stock summary page
            response = s.get(url, headers=headers)
            # Parse the HTML content using BeautifulSoup
            soup = BeautifulSoup(response.content, 'html.parser')
            # Initialize a list to store scores related to the stock
            scores = [t]
            # Iterate over each value in the 'ls' list (assuming 'ls' is defined elsewhere)
            for val in ls[1:]:
                # Sanitize the value (e.g., remove spaces or special characters)
                val = sanitize(val)

                if val in ls_special:
                        # Find the score associated with the value on the stock summary page using css
                        measures = [mea.text.strip() for mea in soup.select('td.t-caption > a')]
                        vals = [va.text.strip() for va in soup.select('td.t-caption span.p-l-sm') ]
                        measures2 = [mea.text.strip() for mea in soup.select('h2.t-h6 >a')]
                        vals2 = [va.text.strip() for va in soup.select('div.flex.flex-center span.t-body-sm.m-l-md')]

                        all_meas = measures + measures2
                        all_vals = vals + vals2
                        
                        for m,v in zip(all_meas,all_vals):
                            if m == val:
                                score = v
                                # print(f"{m}:{score}")
                else:        
                    # Find the score associated with the value on the stock summary page
                    score1 = soup.find('a', string=re.compile(val))
                    if score1 is not None:
                        # Extract the score from the next table cell
                        score = soup.find('a', string=re.compile(val)).find_next('td').text.strip()
                    else:
                        # If the score is not found, set it to "N/A"
                        score = "N/A"

                # Process specific score types (e.g., 'GF Value')
                if val == 'GF Value':
                    # Find relevant elements for 'GF Value'
                    score2 = soup.select('h2 > a', class_="t-h6", string=re.compile(val))
                    i = 0
                    for xt in score2:
                        gf_value = xt.text.strip()
                        if gf_value.find(val) != -1:
                            gf_value2 = gf_value.split('\n')
                            i += 1
                    if i > 0:
                        # Extract the numeric value and convert it to a float
                        gf_value = [x.replace(' ', '') for x in gf_value2]
                        score = gf_value[1].replace('$', '')
                        score = float(score)
                    else:
                        # Set the score to 0 if not found
                        score = 0
                        
                # Append the score to the 'scores' list
                scores.append(score)

        # Update the length of the output DataFrame (assuming 'df_output' is defined elsewhere)
        df_len = len(df_output)
        # Calculate the difference in 'GF Value' and check if it equals 0
        df_output['GF Value'].diff().eq(0)
        # Write the scores to a CSV file (assuming 'guru_csv_path' is defined elsewhere)
        with open(guru_csv_path, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(scores)
        # Update the progress bar
        bar()


c.prCyan("\nProcessing is complete.")

# Export the scraped data to Excel file
df_output.to_excel(guru_xlsx_path, index=False)
c.prCyan(f"Data were saved to file: ")
print(f"{guru_xlsx_path}\n")