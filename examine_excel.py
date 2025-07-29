import pandas as pd

# Read the Excel file
excel_file = 'SampleData.xlsx'

# Get all sheet names
xls = pd.ExcelFile(excel_file)
print(f"Sheets in the Excel file: {xls.sheet_names}")

# Read the first sheet as a sample
first_sheet = xls.sheet_names[0]
print(f"\nReading first sheet: {first_sheet}")
df = pd.read_excel(excel_file, sheet_name=first_sheet)

# Display basic info
print("\nFirst 5 rows of the data:")
print(df.head())

print("\nColumn names:")
print(df.columns.tolist())

print("\nData types:")
print(df.dtypes)
