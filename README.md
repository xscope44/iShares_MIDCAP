# iShares_MIDCAP<br>
GuruFocus parsing of iShares-Russell-Mid-Cap-ETF Holdings Tickers Data<br><br>

Download excel data of iShares-Russell-Mid-Cap-ETF_fund to csv<br>
https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf<br>
Converting downloaded Mid-Cap-ETF csv to excel<br>
Creating new excel file _Out to merge all gathered data and to make Analysis later.<br>
Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out with sheet name iShares-Russell-Mid-Cap-ETF<br>
Parsing GF Data to new excel sheet from GuruFocus<br>
Merging all data in new excel sheet for manual filter Analysis<br>
Adding googlefinance functions, used in google sheets, to get 6 month rps momentum and current price for comparison<br>

Required modules:<br>
 pip3 install pandas openpyxl bs4 simplified_scrapy<br>
 pip3 install progress progressbar2 alive-progress tqdm<br>

Tested on Windows and Mac<br>
