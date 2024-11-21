import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import re

# Read the Excel file
df = pd.read_excel(r"C:\Users\carso\Documents\september\January to April nyambura.xlsx")

# Debugging: Print column names and first few rows
print("Columns:", df.columns)
print("First few rows:")
print(df.head())

# Ensure 'Date' is datetime
if 'Date' in df.columns:
    df['Date'] = pd.to_datetime(df['Date'])
else:
    raise KeyError("The 'Date' column is missing from the DataFrame.")

# Calculate Total Value
df['Total Value'] = df['Price'] * df['Qty Ordered']

# Create a new Excel workbook
wb = openpyxl.Workbook()

# Remove the default sheet
wb.remove(wb.active)

def sanitize_sheet_name(name):
    """Sanitize sheet name to ensure it is valid for Excel."""
    name = name[:31]  # Truncate to 31 characters
    name = re.sub(r'[\\/*?:[\]]', '', name)  # Remove invalid characters
    return name

# Function to add a sheet with data
def add_sheet(wb, name, data):
    sanitized_name = sanitize_sheet_name(name)
    ws = wb.create_sheet(sanitized_name)
    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

# 1. Group by Category
category_groups = df.groupby('Category')
for category, group in category_groups:
    add_sheet(wb, f"Category - {category}", group[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value']])

# 2. Purchase Numbers
purchases = df[df['Qty Ordered'] > 0]
add_sheet(wb, "Purchases", purchases[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value']])

# 3. Inquiry Numbers
inquiries = df[df['Qty Ordered'] == 0]
add_sheet(wb, "Inquiries", inquiries[['No', 'Customer Name', 'Product']])

# 4. Group by Call Outcome
outcome_groups = df.groupby('Call Outcome')
for outcome, group in outcome_groups:
    add_sheet(wb, f"Outcome - {outcome}", group[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value']])

# 5. High Value Customers
high_value = df[df['Total Value'] > df['Total Value'].median()]
add_sheet(wb, "High Value Customers", high_value[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value']])

# 6. Recent Customers (last 30 days)
recent_date = df['Date'].max() - pd.Timedelta(days=30)
recent = df[df['Date'] > recent_date]
add_sheet(wb, "Recent Customers", recent[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value', 'Date']])

# 7. Customers with Feedback
feedback = df[df['Remarks'].notna()]
add_sheet(wb, "Customers with Feedback", feedback[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value', 'Remarks']])

# 8. Group by Media
media_groups = df.groupby('Media')
for media, group in media_groups:
    add_sheet(wb, f"Media - {media}", group[['No', 'Customer Name', 'Product', 'Qty Ordered', 'Total Value']])

# Save the workbook
wb.save('number_groupings_January to April.xlsx')

print("Number groupings have been saved to 'number_groupings_may.xlsx'")
