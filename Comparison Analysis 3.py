import pandas as pd
from fuzzywuzzy import process, fuzz

# Load the Excel file
file_path = r"C:\Users\carso\Documents\July\jbl maping.xlsx"

# Read the two sheets into dataframes
sheet1_df = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2_df = pd.read_excel(file_path, sheet_name='Sheet2')

# Assuming both sheets have columns 'Product' and 'Price'
# Convert product names to lowercase for case-insensitive comparison
sheet1_df['PRODUCT'] = sheet1_df['PRODUCT'].str.lower()
sheet2_df['PRODUCT'] = sheet2_df['PRODUCT'].str.lower()

# Define a function to perform fuzzy matching with a lower threshold
def fuzzy_match(product, choices, threshold=80):
    match, score = process.extractOne(product, choices, scorer=fuzz.partial_ratio)
    return match if score >= threshold else None

# Initialize the new columns
sheet2_df['Available'] = 'Not Available'
sheet2_df['PRICE'] = None
sheet2_df['Page'] = None

# Apply fuzzy matching to determine availability and fetch price and page number
for idx, row in sheet2_df.iterrows():
    matched_product = fuzzy_match(row['PRODUCT'], sheet1_df['PRODUCT'].tolist())
    if matched_product:
        sheet2_df.at[idx, 'Available'] = 'Available'
        # Get the matched row in Sheet1
        matched_row = sheet1_df[sheet1_df['PRODUCT'] == matched_product]
        sheet2_df.at[idx, 'PRICE'] = matched_row['PRICE'].values[0]
        sheet2_df.at[idx, 'Page'] = matched_row.index[0] + 2  # Adding 2 to adjust for header and 1-based index

# Save the result to a new sheet in the same Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    sheet2_df.to_excel(writer, sheet_name='Comparison_Result2', index=False)

print("Comparison completed and saved in 'Comparison_Result' sheet.")
