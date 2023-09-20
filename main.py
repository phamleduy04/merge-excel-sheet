import os
import re
import xlwings as xw

folder_path = 'need_to_merge'
output_file = 'merged_output.xlsx'

# Create a new workbook
wb_out = xw.Book()

# Define the regex pattern
pattern = re.compile('(0[1-9]|1[012])-[12][0-9]{3}')

for file in os.listdir(folder_path):
    if file.endswith('.xlsx'):
        # Open the source workbook
        wb_in = xw.Book(os.path.join(folder_path, file))

        for sheet in wb_in.sheets:
            # Copy each sheet to the new workbook
            sheet.api.Copy(Before=wb_out.sheets[0].api)
            
            # Get the newly created sheet (the first one in the list)
            copied_sheet = wb_out.sheets[0]
            
            # If filename matches the pattern, rename the sheet to the matched data
            match = pattern.search(file)
            if match:
                copied_sheet.name = match.group()
            else :
                copied_sheet.name = file

        # Close the source workbook without saving
        wb_in.close()

# Delete the default sheet created and save the new workbook
del wb_out.sheets['Sheet1']
wb_out.save(output_file)
wb_out.close()

print(f"Merged Excel file '{output_file}' has been created.")
