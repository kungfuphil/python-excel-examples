# This module is to read files with the .xls extension
from xlrd import open_workbook

source_spreadsheet_path = "source_spreadsheet.xls"
source_workbook = open_workbook(source_spreadsheet_path)

# This is the first worksheet available
source_worksheet = source_workbook.sheet_by_index(0)

# Columns in xlrd start at index 0
source_max_columns = source_worksheet.ncols

# Rows in xlrd start at index 1
source_max_rows = source_worksheet.nrows

# Iterate through the rows of the source file
# for i in range(1, source_max_rows): # If you want to skip the header, use this line instead of the next
for i in range(source_max_rows):
    row = source_worksheet.row_values(i)
    for j, source_cell_value in enumerate(row):
        print(f"[{i},{j}]: {source_cell_value}")