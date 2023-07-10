# # # # # # # # read_iShares_excel:
# Loads Data from newest excel with file_prefix eg. iShares-Russell-Mid-Cap-ETF_fund 
# # https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf
# Copy Holdings data sheet to new excel iShares-Russell-Mid-Cap-ETF_Out sheet name iShares-Russell-Mid-Cap-ETF

# # # # # # # # GuruFocus_parser:
# 
# pip3 install pandas openpyxl bs4 simplified_scrapy
# pip3 install progress progressbar2 alive-progress tqdm --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org
# --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org

# # # # # # # # attributes Excel example:
# +--------+-------------------+----------------+-------------------+
# | Ticker | Piotroski F-Score | Altman Z-Score | Beneish M-Score   |
# +--------+-------------------+----------------+-------------------+

# Download Holdings Data by downloading excel from:
# https://www.blackrock.com/us/individual/products/239718/ishares-russell-midcap-etf
# open in excel and save as excel format
# Run read_iShares_excel to generate clean Ticker excel
# Run this GuruFocus_parser to parse data from GuruFocus


# # # # # # # # Add_GF_DataSheet:
# Copy GuruFocus sheet to iShares Out excel

# # # # # # # # Add_Analysis_DataSheet:
# Creates Analysis sheet in iShares Out excel, 
# mergin available data created previously into Analysis sheet
import sys

# List of script file names to include
script_files = ["read_iShares_excel.py", "GuruFocus_parser.py", "Add_GF_DataSheet.py", "Add_Analysis_DataSheet.py"]

# Configuration variable to skip prompt questions
skip_prompt = False

# Function to include and execute a script
def include_script(file_name):
    with open(file_name, "rb") as script_file:
        script_code = compile(script_file.read(), file_name, 'exec')
        exec(script_code, globals())

# Function to display the option menu
def display_menu():
    print("Select an option:")
    print("1. Run all scripts")
    for i, script_file in enumerate(script_files):
        print(f"{i+2}. Run {script_file}")
    print("0. Exit")

# Include and execute the selected script(s)
def run_scripts(selection):
    if selection == 0:
        print("Exiting...")
        sys.exit()

    if selection == 1:
        for script_file in script_files:
            include_script(script_file)
    elif 1 < selection <= len(script_files) + 1:
        script_file = script_files[selection - 2]
        include_script(script_file)
    else:
        print("Invalid selection. Please try again.")

# Main program loop
while True:
    display_menu()

    if skip_prompt:
        selection = 1  # Run all scripts by default
    else:
        user_input = input("Enter your selection: ")
        try:
            selection = int(user_input)
        except ValueError:
            print("Invalid selection. Please try again.")
            continue

    run_scripts(selection)
