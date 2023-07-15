# GuruFocus Parsing of iShares-Russell-Mid-Cap-ETF Holdings Tickers Data
This repository contains a script that facilitates parsing and analyzing the holdings data of iShares Russell Mid-Cap ETF using GuruFocus. The script allows you to download the Excel data of the fund, convert it to CSV format, and perform various data manipulation tasks.

This repository contains a script that enables parsing and analysis of the holdings tickers data for iShares Russell Mid-Cap ETF using GuruFocus. The script automates the following tasks:

Downloading the Excel data of iShares Russell Mid-Cap ETF from the official BlackRock website.
Converting the downloaded Excel data of the Mid-Cap ETF to CSV format.
Creating a new Excel file named "_Out" to merge all the gathered data and facilitate further analysis.
Copying the "Holdings" data sheet from the original Excel file to the newly created Excel file, with the sheet named "iShares-Russell-Mid-Cap-ETF".
Parsing the GuruFocus data and populating it into a new sheet within the Excel file.
Merging all the data in the new Excel sheet for manual filter analysis.
Adding Google Finance functions, used in Google Sheets, to retrieve 6-month RPS momentum and current price data for comparison.
# Required Modules
Ensure that the following Python modules are installed:

pandas
openpyxl
bs4
simplified_scrapy
progress
progressbar2
alive-progress
tqdm

```python
pip3 install pandas openpyxl bs4 simplified_scrapy
pip3 install progress progressbar2 alive-progress tqdm
```

# Compatibility
The script has been tested on both Windows and Mac operating systems.

Feel free to explore and modify the code to suit your specific needs. Happy analyzing!





