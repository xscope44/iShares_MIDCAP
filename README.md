# GuruFocus Parsing of iShares-Russell-Mid-Cap-ETF Holdings Tickers Data
This repository contains a script that facilitates parsing and analyzing the holdings data of iShares Russell Mid-Cap ETF using GuruFocus. The script allows you to download the Excel data of the fund, convert it to CSV format, and perform various data manipulation tasks.

# Instructions
Download the Excel data of iShares Russell Mid-Cap ETF from the following link:
iShares Russell Mid-Cap ETF

Convert the downloaded Mid-Cap-ETF file to CSV format.

Create a new Excel file named _Out to merge all the gathered data and enable further analysis.

Copy the "Holdings" data sheet from the original Excel file to the newly created Excel file (iShares-Russell-Mid-Cap-ETF_Out), ensuring that the sheet name is set as "iShares-Russell-Mid-Cap-ETF".

Parse the GuruFocus data and populate it into a new sheet within the Excel file.

Merge all the data in the new Excel sheet for manual filter analysis.

Optionally, add Google Finance functions to your Google Sheets using the googlefinance library. These functions can help you retrieve 6-month RPS momentum and current price data for comparison.

Required Modules
Ensure that the following Python modules are installed:


```python
pip3 install pandas openpyxl bs4 simplified_scrapy
pip3 install progress progressbar2 alive-progress tqdm
```
Compatibility
This script has been tested on both Windows and Mac operating systems.

Feel free to explore and modify the code to suit your specific needs. Happy analyzing!