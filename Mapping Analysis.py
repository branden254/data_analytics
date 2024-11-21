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
file_path = r"C:\Users\carso\Documents\July\jbl maping.xlsx"
original_df = pd.read_excel(file_path, sheet_name='original_list')
scrambled_df = pd.read_excel(file_path, sheet_name='scrambled_list')
analysis_df = pd.read_excel(file_path, sheet_name='analysis')

# Ensure column names are unique
analysis_df.columns = make_unique(analysis_df.columns)

# Create a dictionary from the analysis data for quick lookup
analysis_dict = analysis_df.set_index('PRODUCT').T.to_dict('list')

# Map analysis data to scrambled list based on product name
mapped_analysis = []
for product in scrambled_df['PRODUCT']:
    if product in analysis_dict:
        mapped_analysis.append([product] + analysis_dict[product])

# Convert mapped analysis to DataFrame
mapped_analysis_df = pd.DataFrame(mapped_analysis, columns=['PRODUCT'] + analysis_df.columns[1:].tolist())

# Save the mapped analysis to a new Excel sheet using openpyxl engine
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
    mapped_analysis_df.to_excel(writer, sheet_name='mapped_analysis', index=False)

print('Mapped analysis saved to the Excel file.')
