from openpyxl import load_workbook
import config

file_gurufocus = config.find_newest_file_simple(config.data_folder_path, config.gurufocus_prefix, config.file_ishares_out_extension)
file_ishares_out = config.find_newest_file_simple(config.data_folder_path, config.ishares_out_prefix, config.file_ishares_out_extension)

# Load the source file
source_workbook = load_workbook(file_gurufocus, read_only=True)
source_sheet = source_workbook[config.gf_sheet_name]

# Load the existing file
existing_workbook = load_workbook(file_ishares_out)
existing_sheet_names = existing_workbook.sheetnames

# Check if the new sheet name already exists in the existing file
if config.ishares_gurufocus_sheet_name in existing_sheet_names:
    # Remove the existing sheet before renaming
    existing_workbook.remove(existing_workbook[config.ishares_gurufocus_sheet_name])

# Create a new sheet in the existing file
existing_workbook.create_sheet(config.ishares_gurufocus_sheet_name)
existing_sheet = existing_workbook[config.ishares_gurufocus_sheet_name]

# Copy the source sheet data to the existing sheet
for row in source_sheet.iter_rows(values_only=True):
    existing_sheet.append(row)

# Save the changes to the existing file
existing_workbook.save(file_ishares_out)
existing_workbook.close()