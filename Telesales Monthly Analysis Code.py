import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# Read the Excel file
df = pd.read_excel(r"C:\Users\carso\Documents\November\Telesales October.xlsx", sheet_name="Sheet1")

# Convert Date column to datetime
df['Date'] = pd.to_datetime(df['Date'])

# 1. Top 5 most products ordered
top_products = df.groupby('Product')['Qty Ordered'].sum().sort_values(ascending=False).head(5)
plt.figure(figsize=(10, 6))
top_products.plot(kind='bar')
plt.title('Top 5 Most Ordered Products')
plt.xlabel('Product')
plt.ylabel('Quantity Ordered')
plt.tight_layout()
plt.savefig('top_5_products.png')
plt.show()
plt.close()

# 2. Top 5 products by margin
top_margin_products = df.groupby('Product')['Margin'].sum().sort_values(ascending=False).head(5)
plt.figure(figsize=(10, 6))
top_margin_products.plot(kind='bar')
plt.title('Top 5 Products by Margin')
plt.xlabel('Product')
plt.ylabel('Total Margin')
plt.tight_layout()
plt.savefig('top_5_margin_products.png')
plt.show()
plt.close()

# 3. Top 5 repetitive numbers
top_numbers = df['No'].value_counts().head(5)
plt.figure(figsize=(10, 6))
top_numbers.plot(kind='bar')
plt.title('Top 5 Repetitive Numbers')
plt.xlabel('Phone Number')
plt.ylabel('Frequency')
plt.tight_layout()
plt.savefig('top_5_repetitive_numbers.png')
plt.show()
plt.close()

# 4. Top 5 Selling Products by Revenue
df['Revenue'] = df['Price'] 
top_revenue_products = df.groupby('Product')['Revenue'].sum().sort_values(ascending=False).head(5)
plt.figure(figsize=(10, 6))
top_revenue_products.plot(kind='bar')
plt.title('Top 5 Selling Products by Revenue')
plt.xlabel('Product')
plt.ylabel('Total Revenue')
plt.tight_layout()
plt.savefig('top_5_revenue_products.png')
plt.show()
plt.close()

# Rest of the code remains the same...
# 5. Day with most calls, day with least calls, week with most profit
calls_per_day = df['Date'].dt.date.value_counts().sort_index()
most_calls_day = calls_per_day.idxmax()
least_calls_day = calls_per_day.idxmin()

df['Week'] = df['Date'].dt.to_period('W')
profit_per_week = df.groupby('Week')['Margin'].sum()
most_profit_week = profit_per_week.idxmax()

print(f"Day with most calls: {most_calls_day}")
print(f"Day with least calls: {least_calls_day}")
print(f"Week with most profit: {most_profit_week}")

# 6. Total inquiries
total_inquiries = len(df)
print(f"Total inquiries: {total_inquiries}")

# 7. Total revenue vs total margin comparison
total_revenue = df['Price'].sum()
total_margin = df['Margin'].sum()

plt.figure(figsize=(10, 6))
plt.bar(['Price', 'Margin'], [total_revenue, total_margin])
plt.title('Total Revenue vs Total Margin')
plt.ylabel('Amount')
plt.tight_layout()
plt.savefig('revenue_vs_margin.png')
plt.show()
plt.close()

print(f"Total Revenue: {total_revenue}")
print(f"Total Margin: {total_margin}")
print(f"Profit Margin: {(total_margin / total_revenue) * 100:.2f}%")

print("Visualizations have been saved as PNG files.")

# Recommendations based on the analysis
recommendations = [
    "1. Focus on promoting and stocking the top 5 most ordered products to meet demand.",
    "2. Emphasize sales of high-margin products to boost overall profitability.",
    "3. Implement a customer loyalty program for frequent callers (top 5 repetitive numbers).",
    "4. Allocate marketing resources to the top 5 revenue-generating products.",
    "5. Analyze factors contributing to high-call and low-call days to optimize staffing.",
    "6. Investigate the most profitable week to replicate successful strategies.",
    "7. Develop strategies to increase the overall profit margin, currently at 18.14%.",
    "8. Improve inventory management to reduce 'Out of Stock' and 'Availability not confirmed' responses.",
    "9. Implement a robust follow-up system for the high number of 'Follow Up-Needed' outcomes.",
    "10. Conduct regular price and margin analysis to ensure competitive pricing while maintaining profitability."
]

print("\nRecommendations:")
for rec in recommendations:
    print(rec)
    import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from wordcloud import WordCloud

# Read the Excel file
df = pd.read_excel(r"C:\Users\carso\Documents\November\Telesales October.xlsx", sheet_name="Sheet1")
# Convert Date column to datetime
df['Date'] = pd.to_datetime(df['Date'])

# 1. Pie chart: Distribution of inquiries between website and call channels
plt.figure(figsize=(10, 6))
df['Media'] = df['Media'].str.lower().replace({'website ': 'website'})  # Standardize 'website' entries
channel_counts = df['Media'].value_counts()
channel_counts = channel_counts.reindex(['website', 'call'])  # Ensure only these two categories
channel_counts.plot(kind='pie', autopct='%1.1f%%')
plt.title('Distribution of Inquiries by Channel')
plt.savefig('channel_distribution.png')
plt.show()
plt.close()

# 2. Bar chart: Distribution of call outcomes
plt.figure(figsize=(12, 6))
df['Call Outcome'].value_counts().plot(kind='bar')
plt.title('Distribution of Call Outcomes')
plt.xlabel('Outcome')
plt.ylabel('Count')
plt.savefig('call_outcomes.png')
plt.show()
plt.close()
call_outcome_counts = df['Call Outcome'].value_counts()
print(f"call outcomes: {call_outcome_counts}")

# 3. Scatter plot: Price vs. Margin
plt.figure(figsize=(10, 6))
plt.scatter(df['Price'], df['Margin'])
plt.title('Price vs. Margin')
plt.xlabel('Price')
plt.ylabel('Margin')
plt.savefig('price_vs_margin.png')
plt.show()
plt.close()

# 4. Stacked bar chart: Product availability status across categories
availability_by_category = df.groupby('Category')['Call Outcome'].value_counts().unstack()
availability_by_category.plot(kind='bar', stacked=True, figsize=(12, 6))
plt.title('Product Availability by Category')
plt.xlabel('Category')
plt.ylabel('Count')
plt.legend(title='Availability Status', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig('availability_by_category.png')
plt.show()
plt.close()

# 5. Funnel chart: Sales process from inquiry to closed sale
stages = ['Total Inquiries', 'Follow Up-Needed', 'Closed Sale']
values = [len(df), len(df[df['Call Outcome'] == 'Follow Up-Needed']), len(df[df['Call Outcome'] == 'Closed Sale'])]

plt.figure(figsize=(10, 6))
plt.bar(stages, values)
plt.title('Sales Funnel')
plt.xlabel('Stage')
plt.ylabel('Count')
plt.savefig('sales_funnel.png')
plt.show()
plt.close()

# 6. Heat map: Inquiry volume by date and product category
inquiry_volume = df.groupby(['Date', 'Category']).size().unstack()
plt.figure(figsize=(12, 8))
sns.heatmap(inquiry_volume, cmap='YlOrRd')
plt.title('Inquiry Volume by Date and Category')
plt.savefig('inquiry_heatmap.png')
plt.show()
plt.close()

# 7. Word cloud of most frequently mentioned products or brands
text = ' '.join(df['Product'].dropna())
wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')
plt.title('Most Frequently Mentioned Products')
plt.savefig('product_wordcloud.png')
plt.show()
plt.close()

print("Visualizations have been saved as PNG files.")

# Recommendations based on the analysis
recommendations = [
    "1. Improve inventory management to reduce 'Out of Stock' and 'Availability not confirmed' responses.",
    "2. Implement a robust follow-up system for the high number of 'Follow Up-Needed' outcomes.",
    "3. Focus marketing efforts on popular product categories like tools, electronics, and laboratory equipment.",
    "4. Consider bundling high-margin products with related items to increase overall profitability.",
    "5. Standardize data entry processes to ensure consistency in phone number formatting and product naming.",
    "6. Analyze seasonal trends to optimize inventory for school supplies and sports equipment before the academic year.",
    "7. Develop a customer loyalty program to encourage repeat business from frequent inquirers.",
    "8. Invest in real-time inventory tracking to improve accuracy of product availability information.",
    "9. Train sales staff on effective upselling and cross-selling techniques for high-value items.",
    "10. Conduct regular price and margin analysis to ensure competitive pricing while maintaining profitability."
]

print("\nRecommendations:")
for rec in recommendations:
    print(rec)
    
    # 15. Category Sales Analysis
plt.figure(figsize=(12, 6))
category_sales = df.groupby('Category')['Price'].sum().sort_values(ascending=False)
category_sales.plot(kind='bar', color='skyblue')
plt.title('Total Sales by Category (Monthly)')
plt.xlabel('Category')
plt.ylabel('Sales (Price)')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.grid(axis='y')
plt.savefig('monthly_category_sales_analysis.png')
plt.show()
plt.close()

print("Sales by Category:")
for category, sales in category_sales.items():
    print(f'{category}: {sales:.2f}')

# 16. Inquiries per Category
plt.figure(figsize=(12, 6))
inquiries_per_category = df['Category'].value_counts()
inquiries_per_category.plot(kind='bar', color='lightgreen')
plt.title('Inquiries per Category (Monthly)')
plt.xlabel('Category')
plt.ylabel('Number of Inquiries')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.grid(axis='y')
plt.savefig('monthly_inquiries_per_category.png')
plt.show()
plt.close()

print("\nInquiries per Category:")
for category, count in inquiries_per_category.items():
    print(f'{category}: {count} inquiries')

# 17. Inquiries vs Number of Sales per Category
category_inquiries = df['Category'].value_counts()
category_sales_count = df[df['Price'] > 0]['Category'].value_counts()

category_comparison = pd.DataFrame({'Inquiries': category_inquiries, 'Sales': category_sales_count}).fillna(0)
category_comparison = category_comparison.sort_values('Inquiries', ascending=False)

fig, ax1 = plt.subplots(figsize=(14, 7))

ax1.bar(category_comparison.index, category_comparison['Inquiries'], alpha=0.7, color='lightblue', label='Inquiries')
ax1.set_xlabel('Category')
ax1.set_ylabel('Number of Inquiries', color='blue')
ax1.tick_params(axis='y', labelcolor='blue')
ax1.set_xticklabels(category_comparison.index, rotation=45, ha='right')

ax2 = ax1.twinx()
ax2.plot(category_comparison.index, category_comparison['Sales'], color='red', marker='o', label='Sales')
ax2.set_ylabel('Number of Sales', color='red')
ax2.tick_params(axis='y', labelcolor='red')

plt.title('Inquiries vs Number of Sales per Category (Monthly)')
fig.legend(loc='upper right', bbox_to_anchor=(1,1), bbox_transform=ax1.transAxes)

plt.tight_layout()
plt.savefig('monthly_inquiries_vs_sales_count_per_category.png')
plt.show()
plt.close()

print("\nInquiries vs Number of Sales per Category:")
for category in category_comparison.index:
    inquiries = int(category_comparison.loc[category, 'Inquiries'])
    sales = int(category_comparison.loc[category, 'Sales'])
    print(f'{category}: {inquiries} inquiries, {sales} sales')

print("\nConversion Rates:")
for category in category_comparison.index:
    inquiries = category_comparison.loc[category, 'Inquiries']
    sales = category_comparison.loc[category, 'Sales']
    if inquiries > 0:
        conversion_rate = (sales / inquiries) * 100
        print(f'{category}: {conversion_rate:.2f}%')
    else:
        print(f'{category}: N/A (no inquiries)')
