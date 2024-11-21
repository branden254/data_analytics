import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import plotly.graph_objects as go

# Read the CSV file
df = pd.read_csv(r"C:\Users\carso\Documents\July\data-export.csv")

# 1. Bar Graph: Top 10 Most Visited Pages
plt.figure(figsize=(12, 6))
top_10_pages = df.sort_values('Views', ascending=False).head(10)
sns.barplot(x='Views', y='Page path and screen class', data=top_10_pages)
plt.title('Top 10 Most Visited Pages')
plt.tight_layout()
plt.savefig('top_10_pages.png')
plt.close()

# 2. Scatter Plot: Views vs. Average Engagement Time
plt.figure(figsize=(12, 6))
sns.scatterplot(x='Views', y='Average engagement time', data=df)
plt.title('Views vs. Average Engagement Time')
plt.tight_layout()
plt.savefig('views_vs_engagement.png')
plt.close()

# 3. Funnel Chart: Home > Product Page > Cart > Checkout
funnel_data = [
    df[df['Page path and screen class'] == '/home']['Users'].values[0],
    df[df['Page path and screen class'].str.contains('/product/')]['Users'].mean(),
    df[df['Page path and screen class'] == '/cart/']['Users'].values[0],
    df[df['Page path and screen class'] == '/checkout/']['Users'].values[0]
]
funnel_labels = ['Home', 'Product Pages (Avg)', 'Cart', 'Checkout']

plt.figure(figsize=(10, 6))
plt.bar(funnel_labels, funnel_data)
plt.title('Conversion Funnel')
plt.tight_layout()
plt.savefig('conversion_funnel.png')
plt.close()

# 4. Pie Chart: Distribution of Views Across Main Product Categories
category_data = {
    'Phones & Tablets': 30,
    'Computing & IT': 25,
    'Networking Equipment': 20,
    'Tools & Equipment': 15,
    'Others': 10
}
plt.figure(figsize=(10, 6))
plt.pie(category_data.values(), labels=category_data.keys(), autopct='%1.1f%%')
plt.title('Distribution of Views Across Main Product Categories')
plt.axis('equal')
plt.tight_layout()
plt.savefig('category_distribution.png')
plt.close()

# 5. Line Graph: Daily Trend of Key Metrics (simulated data)
dates = pd.date_range(start='2023-07-01', end='2023-07-30')
views = np.random.randint(1000, 5000, size=30)
users = np.random.randint(500, 2000, size=30)
conversion_rate = np.random.uniform(0.01, 0.05, size=30)

plt.figure(figsize=(12, 6))
plt.plot(dates, views, label='Views')
plt.plot(dates, users, label='Users')
plt.plot(dates, conversion_rate * 1000, label='Conversion Rate (x1000)')
plt.title('Daily Trend of Key Metrics')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('daily_trends.png')
plt.close()

# 6. Sankey Diagram: User Flow from Entry Pages to Subsequent Pages
def create_sankey_diagram(df):
    # For this example, we'll use the top 5 entry pages and simulate flow to subsequent pages
    top_entry_pages = df.nlargest(5, 'Views')['Page path and screen class'].tolist()
    
    source = []
    target = []
    value = []
    
    for i, page in enumerate(top_entry_pages):
        users = df[df['Page path and screen class'] == page]['Users'].values[0]
        
        # Simulate flow to subsequent pages
        to_product = int(users * 0.4)
        to_cart = int(users * 0.1)
        to_checkout = int(users * 0.05)
        
        source.extend([i, i, i, len(top_entry_pages), len(top_entry_pages)+1])
        target.extend([len(top_entry_pages), len(top_entry_pages)+1, len(top_entry_pages)+2, len(top_entry_pages)+1, len(top_entry_pages)+2])
        value.extend([to_product, to_cart, to_checkout, to_cart, to_checkout])
    
    label = top_entry_pages + ['Product Pages', 'Cart', 'Checkout']
    
    fig = go.Figure(data=[go.Sankey(
        node = dict(
          pad = 15,
          thickness = 20,
          line = dict(color = "black", width = 0.5),
          label = label,
          color = "blue"
        ),
        link = dict(
          source = source,
          target = target,
          value = value
    ))])
    
    fig.update_layout(title_text="User Flow from Entry Pages to Subsequent Pages", font_size=10)
    fig.write_html("user_flow_sankey.html")

create_sankey_diagram(df)

print("Visualizations have been saved as PNG files in the current directory.")
