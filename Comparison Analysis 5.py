import pandas as pd

# Load the provided Excel file
file_path = r"C:\Users\carso\Documents\July\analysis phones mobile.xlsx" # Update the path to your file
excel_data = pd.ExcelFile(file_path)

# Load data from Sheet1 and Sheet2
sheet1_data = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2_data = pd.read_excel(file_path, sheet_name='Sheet2')

# Define a function to normalize product names
def normalize_product_name(name):
    name = name.lower()
    replacements = {
        '2yr': '2 years',
        '1yr': '1 year',
        '3yr': '3 years',
        # Add more replacements as needed
    }
    for key, value in replacements.items():
        name = name.replace(key, value)
    return name

# Normalize product names in both sheets
sheet1_data['Normalized_PRODUCT'] = sheet1_data['PRODUCT'].apply(normalize_product_name)
sheet2_data['Normalized_PRODUCT'] = sheet2_data['PRODUCT'].apply(normalize_product_name)

# Create the comparison DataFrame
comparison_data = pd.DataFrame()
comparison_data['Product'] = sheet2_data['PRODUCT']  # Column 1: All products in Sheet2

# Initialize other columns
comparison_data['Price_Sheet1'] = None
comparison_data['Availability_Sheet1'] = 'No'
comparison_data['Page_Number_Sheet1'] = None

# Iterate through products in Sheet2 and find matches in Sheet1
for index, product in sheet2_data.iterrows():
    normalized_product_name = product['Normalized_PRODUCT']
    match = sheet1_data[sheet1_data['Normalized_PRODUCT'] == normalized_product_name]
    
    if not match.empty:
        comparison_data.at[index, 'Price_Sheet1'] = match['PRICE'].values[0]
        comparison_data.at[index, 'Availability_Sheet1'] = 'Yes'
        comparison_data.at[index, 'Page_Number_Sheet1'] = match.index[0] + 1

# Save the comparison DataFrame to a new sheet in the Excel file
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    comparison_data.to_excel(writer, sheet_name='Sheet4_Comparison', index=False)
