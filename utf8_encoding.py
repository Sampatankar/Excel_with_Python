import os
import csv
import openpyxl

"""
1. Loop through every file that is a .csv in a folder.
2. Open each .csv
3. Change the text encoding to utf-8
4. Save the changes as an excel file - naming to be decided.
"""


# Place the full path of the folder containing the .csv's in place of 'full_folder_path':
directory = r'full_folder_path'

# Loop through every .csv file in the folder above:
for filename in os.scandir(directory):
    # Instantiate an OpenPYXl workbook:
    wb = openpyxl.Workbook()
    # Make the opened sheet the active sheet:
    ws = wb.active
    
    # Check for appropriate files:
    if (filename.path.endswith(".csv")) and filename.is_file():
        print(filename.path)

        # Open the .csv:
        with open(filename.path, encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                ws.append(row)

            # Remove the .csv:
            split_string = filename.path.split(".", 1)
            substring = split_string[0]
            print(substring)

            # Save the Excel file with .csv filename:
            wb.save(substring + '.xlsx')
