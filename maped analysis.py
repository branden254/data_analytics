import pandas as pd

# Function to make column names unique
def make_unique(column_names):
    counts = {}
    new_column_names = []
    for col in column_names:
        if col not in counts:
            counts[col] = 1
            new_column_names.append(col)
        else:
            counts[col] += 1
            new_column_names.append(f"{col}_{counts[col]}")
    return new_column_names

# Load original, scrambled lists, and analysis from Excel sheets
file_path = r"C:\Users\carso\Documents\July\maping gaming.xlsx"
original_df = pd.read_excel(file_path, sheet_name='original_list')
scrambled_df = pd.read_excel(file_path, sheet_name='scrambled_list')
analysis_df = pd.read_excel(file_path, sheet_name='analysis')

# Ensure column names are unique
analysis_df.columns = make_unique(analysis_df.columns)

# Create a dictionary from the analysis data for quick lookup
analysis_dict = analysis_df.set_index('PRODUCT').T.to_dict('list')

# Map analysis data to scrambled list based on product name
mapped_analysis = []
unmatched_products = []
for product in scrambled_df['PRODUCT']:
    if product in analysis_dict:
        mapped_analysis.append([product] + analysis_dict[product])
    else:
        unmatched_products.append([product, scrambled_df[scrambled_df['PRODUCT'] == product]['PRICE'].values[0]])

# Convert mapped analysis to DataFrame
mapped_analysis_df = pd.DataFrame(mapped_analysis, columns=['PRODUCT'] + analysis_df.columns[1:].tolist())

# Convert unmatched products to DataFrame
unmatched_products_df = pd.DataFrame(unmatched_products, columns=['PRODUCT', 'PRICE'])

# Save the mapped analysis and unmatched products to new Excel sheets using openpyxl engine
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
    mapped_analysis_df.to_excel(writer, sheet_name='mapped_analysis', index=False)
    unmatched_products_df.to_excel(writer, sheet_name='unmatched_products', index=False)

print('Mapped analysis and unmatched products saved to the Excel file.')
