import pandas as pd
from fuzzywuzzy import fuzz
import numpy as np

def fuzzy_match(row, sheet1_products, threshold=80):
    matches = sheet1_products.apply(lambda x: fuzz.token_set_ratio(row['Product'], x))
    best_match_index = matches.idxmax()
    if matches[best_match_index] >= threshold:
        return sheet1_products[best_match_index]
    return np.nan

# Read the Excel file
excel_file = r"C:\Users\carso\Documents\october\WEEK 5\comparisons\Networking Comparisons.xlsx"
sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1')
sheet2 = pd.read_excel(excel_file, sheet_name='Sheet2')

# Perform fuzzy matching
sheet2['matched_product'] = sheet2.apply(fuzzy_match, args=(sheet1['Product'],), axis=1)

# Identify products in both sheets (including fuzzy matches)
products_in_both = sheet2[sheet2['matched_product'].notna()].merge(
    sheet1, left_on='matched_product', right_on='Product', suffixes=('_sheet2', '_sheet1')
)

# Identify products not in Sheet1
products_not_in_sheet1 = sheet2[sheet2['matched_product'].isna()]

# Create a new Excel writer object
with pd.ExcelWriter('comparison_analysis_genspace.xlsx') as writer:
    # Write original sheets
    sheet1.to_excel(writer, sheet_name='Sheet1', index=False)
    sheet2.to_excel(writer, sheet_name='Sheet2', index=False)
    
    # Write comparison sheets
    products_in_both.to_excel(writer, sheet_name='Products in Both', index=False)
    products_not_in_sheet1.to_excel(writer, sheet_name='Products Not in Sheet1', index=False)

print("Comparison analysis completed. Check 'comparison_analysis_Genspace.xlsx'")

