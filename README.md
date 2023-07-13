# iShares_MIDCAP
GuruFocus parsing of iShares-Russell-Mid-Cap-ETF Holdings Tickers Data

Download excel data of iShares-Russell-Mid-Cap-ETF_fund to csv
https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf
Converting downloaded Mid-Cap-ETF csv to excel
Creating new excel file _Out to merge all gathered data and to make Analysis later.
Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out with sheet name iShares-Russell-Mid-Cap-ETF
Parsing GF Data to new excel sheet from GuruFocus
Merging all data in new excel sheet for manual filter Analysis
Adding googlefinance functions, used in google sheets, to get 6 month rps momentum and current price for comparison

Required modules:
 pip3 install pandas openpyxl bs4 simplified_scrapy
 pip3 install progress progressbar2 alive-progress tqdm

Tested on Windows and Mac
