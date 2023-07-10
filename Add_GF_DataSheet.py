from openpyxl import load_workbook
import config

folder_path = config.folder_path
file_output_prefix = config.file_ishares_out_prefix
file_gurufocus_prefix = config.file_gurufocus_prefix
file_ishares_out_extension = config.file_ishares_out_extension
new_sheet_name = config.ishares_gurufocus_sheet_name  # New name for the imported sheet
gf_sheet_name = config.gf_sheet_name  # Name of the sheet to import

file_gurufocus = config.find_newest_file_simple(folder_path, file_gurufocus_prefix, file_ishares_out_extension)
file_ishares_out = config.find_newest_file_simple(folder_path, file_output_prefix, file_ishares_out_extension)

# Load the source file
source_workbook = load_workbook(file_gurufocus, read_only=True)
source_sheet = source_workbook[gf_sheet_name]

# Load the existing file
existing_workbook = load_workbook(file_ishares_out)
existing_sheet_names = existing_workbook.sheetnames

# Check if the new sheet name already exists in the existing file
if new_sheet_name in existing_sheet_names:
    # Remove the existing sheet before renaming
    existing_workbook.remove(existing_workbook[new_sheet_name])

# Create a new sheet in the existing file
existing_workbook.create_sheet(new_sheet_name)
existing_sheet = existing_workbook[new_sheet_name]

# Copy the source sheet data to the existing sheet
for row in source_sheet.iter_rows(values_only=True):
    existing_sheet.append(row)

# Save the changes to the existing file
existing_workbook.save(file_ishares_out)