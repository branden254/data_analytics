import pandas as pd
import openpyxl
import os
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def fuzzy_match(x, choices, cutoff=80):
    match = process.extractOne(x, choices, scorer=fuzz.token_sort_ratio)
    return match[0] if match and match[1] >= cutoff else None

# Set the file path directly
file_path = r"C:\Users\carso\Documents\october\matched camers.xlsx"

# Check if the file exists
if not os.path.exists(file_path):
    print(f"Error: The file '{file_path}' does not exist.")
    exit(1)

# Read the Excel sheets
try:
    with pd.ExcelFile(file_path) as xls:
        sheet1 = pd.read_excel(xls, 'Sheet1')
        sheet2 = pd.read_excel(xls, 'Sheet2')
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Create 'Matched Products' sheet
sheet1['Product_Lower'] = sheet1['Product'].str.lower()
sheet2['Product_Lower'] = sheet2['Product'].str.lower()

matched_products = pd.merge(sheet2, sheet1, on='Product_Lower', how='inner', suffixes=('_Sheet2', '_Sheet1'))
matched_products = matched_products.drop(['Product_Lower'], axis=1)

# Create 'Matched 2' sheet
matched2 = sheet1.copy()
matched2.insert(0, 'Product from Sheet2', '')
matched2.insert(1, 'Match Type', '')

# Perform fuzzy matching
sheet2_products = sheet2['Product'].tolist()
for index, row in matched2.iterrows():
    exact_match = row['Product'] in sheet2_products
    if exact_match:
        matched2.at[index, 'Product from Sheet2'] = row['Product']
        matched2.at[index, 'Match Type'] = 'Exact'
    else:
        fuzzy_match_result = fuzzy_match(row['Product'], sheet2_products)
        if fuzzy_match_result:
            matched2.at[index, 'Product from Sheet2'] = fuzzy_match_result
            matched2.at[index, 'Match Type'] = 'Fuzzy'
        else:
            matched2.at[index, 'Match Type'] = 'No Match'

# Drop temporary columns
matched2 = matched2.drop(['Product_Lower'], axis=1)

# Write results back to the same Excel file
try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        matched_products.to_excel(writer, sheet_name='Matched Products', index=False)
        matched2.to_excel(writer, sheet_name='Matched 2', index=False)
    print(f"Processing complete. Results added to '{file_path}'.")
except Exception as e:
    print(f"Error writing to Excel file: {e}")
    print("This error often occurs if the file is open in Excel. Please close the file and run the script again.")
    