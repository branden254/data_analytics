import pandas as pd
import openpyxl
import os
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def fuzzy_match(x, choices, cutoff=80):
    # Ensure that 'x' is a string
    if isinstance(x, str):
        # Focus on important product details like model numbers
        match = process.extractOne(x, choices, scorer=fuzz.token_sort_ratio)
        if match and match[1] >= cutoff:
            return match[0], choices.index(match[0]) + 1  # Return match and its position (row number)
    return None, None

# Set the file path directly
file_path = r"C:\Users\carso\Documents\september\Level Up Tech Store\price lists\Level Up Comparisons.xlsx"

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

# Ensure 'Product' columns are strings and handle missing values
sheet1['Product'] = sheet1['Product'].fillna('').astype(str)
sheet2['Product'] = sheet2['Product'].fillna('').astype(str)

# Create 'Matched Products' sheet
sheet1['Product_Lower'] = sheet1['Product'].str.lower()
sheet2['Product_Lower'] = sheet2['Product'].str.lower()

matched_products = pd.merge(sheet2, sheet1, on='Product_Lower', how='inner', suffixes=('_Sheet2', '_Sheet1'))
matched_products = matched_products.drop(['Product_Lower'], axis=1)

# Create 'Matched 2' sheet
matched2 = sheet1.copy()
matched2.insert(0, 'Product from Sheet2', '')
matched2.insert(1, 'Match Type', '')
matched2.insert(2, 'Matched Sheet2 Row', '')  # Add column for matched row in Sheet2

# Perform fuzzy matching
sheet2_products = sheet2['Product'].tolist()
for index, row in matched2.iterrows():
    exact_match = row['Product'] in sheet2_products
    if exact_match:
        matched2.at[index, 'Product from Sheet2'] = row['Product']
        matched2.at[index, 'Match Type'] = 'Exact'
        matched2.at[index, 'Matched Sheet2 Row'] = sheet2[sheet2['Product'] == row['Product']].index[0] + 1  # Row number in Sheet2
    else:
        fuzzy_match_result, matched_row = fuzzy_match(row['Product'], sheet2_products)
        if fuzzy_match_result:
            matched2.at[index, 'Product from Sheet2'] = fuzzy_match_result
            matched2.at[index, 'Match Type'] = 'Fuzzy'
            matched2.at[index, 'Matched Sheet2 Row'] = matched_row  # Store the matched row number
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
