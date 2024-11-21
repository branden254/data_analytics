import pandas as pd

def find_duplicates(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Find duplicates across all categories
    cross_category_dupes = df[df.duplicated(subset=['Product'], keep=False)]
    
    # Find duplicates within categories
    within_category_dupes = df[df.duplicated(subset=['Product', 'Category'], keep=False)]
    
    # Create a new Excel writer object
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Write cross-category duplicates
        cross_category_dupes.to_excel(writer, sheet_name='Cross Category Dupes', index=False)
        
        # Write within-category duplicates
        within_category_dupes.to_excel(writer, sheet_name='Within Category Dupes', index=False)

    print("Duplicate analysis complete. Check the new sheets in your Excel file.")

# Usage
file_path = r"C:\Users\carso\Documents\september\Flex\PRICE LISTS\All Products Flex.xlsx"
find_duplicates(file_path)
