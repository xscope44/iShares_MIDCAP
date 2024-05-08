# run.py: 
# Description: show menu selection to run scripts in order or give user choice to run one

# # # # # # # # Download Holdings Data by downloading excel from blackrock web

# # # # # # # # GuruFocus_parser

# # # # # # # # Add_GF_DataSheet

# # # # # # # # Add_Analysis_DataSheet

import sys
import config as c
# List of script file names to include
script_files = ["Download_iShares_midcap_data.py", "GuruFocus_parser.py", "Add_GF_DataSheet.py", "Add_Analysis_DataSheet.py"]

# Configuration variable to skip prompt questions
skip_prompt = False

# Function to include and execute a script
def include_script(file_name):
    with open(file_name, "rb") as script_file:
        script_code = compile(script_file.read(), file_name, 'exec')
        exec(script_code, globals())

# Function to display the option menu
def display_menu():
    c.prCyan ("Select an option:")
    print("1. Run all scripts")
    for i, script_file in enumerate(script_files):
        print(f"{i+2}. Run {script_file}")
    print("0. Exit")

# Include and execute the selected script(s)
def run_scripts(selection):
    if selection == 0:
        c.prCyan("Exiting...")
        sys.exit()

    if selection == 1:
        for script_file in script_files:
            c.prCyan(f"Running script file: {script_file}")
            include_script(script_file)
    elif 1 < selection <= len(script_files) + 1:
        script_file = script_files[selection - 2]
        include_script(script_file)
    else:
        c.prRed("Invalid selection. Please try again.")

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
            c.prRed("Invalid selection. Please try again.")
            continue

    run_scripts(selection)
