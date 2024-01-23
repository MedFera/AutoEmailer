import pandas as pd

file_path = 'Assigned eCards GSPGuard.xlsx'

# Get the list of sheet names without reading any data
sheet_names = pd.read_excel(file_path, sheet_name=None).keys()
sheets = []
# Print all the sheet names
for sheet_name in sheet_names:
    print(sheet_name)
    sheets.append(sheet_name)