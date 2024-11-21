import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

# File paths configuration remains the same
FILE_PATHS = {
    "July": r"C:\Users\carso\Documents\November\comparisons telesales\telecalls July.xlsx",
    "August": r"C:\Users\carso\Documents\November\comparisons telesales\telecalls august.xlsx",
    "September": r"C:\Users\carso\Documents\November\comparisons telesales\telecalls september.xlsx",
    "October": r"C:\Users\carso\Documents\November\comparisons telesales\Telesales October.xlsx",
}

def load_multiple_months_data(file_paths):
    """
    Load data for multiple months from Excel files using specific file paths
    """
    all_data = []
    
    for month, file_path in file_paths.items():
        try:
            df = pd.read_excel(file_path, sheet_name="Sheet1")
            df['Month'] = month
            df['Date'] = pd.to_datetime(df['Date'])
            # Add week number for weekly analysis
            df['Week'] = df['Date'].dt.isocalendar().week
            df['WeekYear'] = df['Date'].dt.strftime('%Y-%V')
            all_data.append(df)
            print(f"Successfully loaded data for {month}")
        except Exception as e:
            print(f"Error loading data for {month}: {e}")
    
    return pd.concat(all_data) if all_data else None

def create_enhanced_visualizations(df, output_dir="output"):
    """
    Create and save enhanced analysis visualizations
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # 1. Monthly Growth Trend Analysis
    plt.figure(figsize=(15, 8))
    monthly_metrics = df.groupby('Month').agg({
        'Price': 'sum',
        'Margin': 'sum',
        'No': 'count'
    }).reset_index()
    
    # Sort months chronologically
    month_order = ['July', 'August', 'September', 'October']
    monthly_metrics['Month'] = pd.Categorical(monthly_metrics['Month'], categories=month_order, ordered=True)
    monthly_metrics = monthly_metrics.sort_values('Month')
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(15, 12))
    
    # Revenue and Margin Trends
    ax1.plot(monthly_metrics['Month'], monthly_metrics['Price'], marker='o', label='Revenue', linewidth=2)
    ax1.plot(monthly_metrics['Month'], monthly_metrics['Margin'], marker='s', label='Margin', linewidth=2)
    ax1.set_title('Monthly Revenue and Margin Growth Trend')
    ax1.set_ylabel('Amount')
    ax1.legend()
    ax1.grid(True)
    
    # Inquiries Trend
    ax2.plot(monthly_metrics['Month'], monthly_metrics['No'], marker='o', color='green', linewidth=2)
    ax2.set_title('Monthly Inquiries Trend')
    ax2.set_ylabel('Number of Inquiries')
    ax2.grid(True)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/monthly_growth_trends.png')
    plt.show()
    plt.close()

    # 2. Top 3 Days by Inquiries for Each Month
    top_inquiry_days = []
    for month in df['Month'].unique():
        month_data = df[df['Month'] == month]
        daily_inquiries = month_data.groupby(month_data['Date'].dt.date).size().reset_index()
        daily_inquiries.columns = ['Date', 'Inquiries']
        top_3 = daily_inquiries.nlargest(3, 'Inquiries')
        top_3['Month'] = month
        top_inquiry_days.append(top_3)
    
    top_inquiry_days_df = pd.concat(top_inquiry_days)
    print("\nTop 3 Days with Most Inquiries by Month:")
    print(top_inquiry_days_df.to_string(index=False))
    
    # Save to Excel
    top_inquiry_days_df.to_excel(f'{output_dir}/top_inquiry_days.xlsx', index=False)

    # 3. Weekly Performance Analysis
    weekly_metrics = df.groupby(['WeekYear', 'Month']).agg({
        'Price': 'sum',
        'Margin': 'sum',
        'No': 'count'
    }).reset_index()
    
    weekly_metrics['Conversion_Rate'] = weekly_metrics.apply(
        lambda x: (len(df[(df['WeekYear'] == x['WeekYear']) & (df['Price'] > 0)]) / x['No'] * 100)
        if x['No'] > 0 else 0, axis=1
    )
    
    # Plot weekly trends
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 15))
    
    # Weekly Revenue
    sns.barplot(data=weekly_metrics, x='WeekYear', y='Price', hue='Month', ax=ax1)
    ax1.set_title('Weekly Revenue')
    ax1.tick_params(axis='x', rotation=45)
    
    # Weekly Margin
    sns.barplot(data=weekly_metrics, x='WeekYear', y='Margin', hue='Month', ax=ax2)
    ax2.set_title('Weekly Margin')
    ax2.tick_params(axis='x', rotation=45)
    
    # Weekly Inquiries
    sns.barplot(data=weekly_metrics, x='WeekYear', y='No', hue='Month', ax=ax3)
    ax3.set_title('Weekly Inquiries')
    ax3.tick_params(axis='x', rotation=45)
    
    # Weekly Conversion Rate
    sns.barplot(data=weekly_metrics, x='WeekYear', y='Conversion_Rate', hue='Month', ax=ax4)
    ax4.set_title('Weekly Conversion Rate')
    ax4.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/weekly_performance_analysis.png')
    plt.show()
    plt.close()

    # 4. Category Performance Analysis
    category_metrics = df.groupby(['Month', 'Category']).agg({
        'Price': 'sum',
        'Margin': 'sum',
        'No': 'count'
    }).reset_index()
    
    category_metrics['Conversion_Rate'] = category_metrics.apply(
        lambda x: (len(df[(df['Month'] == x['Month']) & 
                         (df['Category'] == x['Category']) & 
                         (df['Price'] > 0)]) / x['No'] * 100)
        if x['No'] > 0 else 0, axis=1
    )
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 15))
    
    # Category Revenue
    sns.barplot(data=category_metrics, x='Category', y='Price', hue='Month', ax=ax1)
    ax1.set_title('Category Revenue by Month')
    ax1.tick_params(axis='x', rotation=45)
    
    # Category Margin
    sns.barplot(data=category_metrics, x='Category', y='Margin', hue='Month', ax=ax2)
    ax2.set_title('Category Margin by Month')
    ax2.tick_params(axis='x', rotation=45)
    
    # Category Inquiries
    sns.barplot(data=category_metrics, x='Category', y='No', hue='Month', ax=ax3)
    ax3.set_title('Category Inquiries by Month')
    ax3.tick_params(axis='x', rotation=45)
    
    # Category Conversion Rate
    sns.barplot(data=category_metrics, x='Category', y='Conversion_Rate', hue='Month', ax=ax4)
    ax4.set_title('Category Conversion Rate by Month')
    ax4.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/category_performance_analysis.png')
    plt.show()
    plt.close()


    # === NEW ANALYSIS 1: Advanced Category Performance Matrix ===
    plt.figure(figsize=(15, 10))
    
    # Calculate key category metrics
    category_matrix = df.groupby('Category').agg({
        'Price': ['sum', 'mean', 'count'],
        'Margin': ['sum', 'mean'],
        'Qty Ordered': 'sum'
    }).reset_index()
    
    # Calculate average margin percentage
    category_matrix['Margin_Percentage'] = (
        category_matrix[('Margin', 'sum')] / category_matrix[('Price', 'sum')] * 100
    )
    
    # Create bubble chart
    plt.figure(figsize=(12, 8))
    plt.scatter(
        category_matrix[('Price', 'mean')],
        category_matrix['Margin_Percentage'],
        s=category_matrix[('Price', 'count')] * 50,  # Size based on number of sales
        alpha=0.6
    )
    
    # Add category labels
    for i, category in enumerate(category_matrix['Category']):
        plt.annotate(
            category,
            (category_matrix[('Price', 'mean')][i], category_matrix['Margin_Percentage'][i])
        )
    
    plt.title('Category Performance Matrix')
    plt.xlabel('Average Sale Value')
    plt.ylabel('Margin Percentage')
    plt.savefig(f'{output_dir}/category_performance_matrix.png')
    plt.show()
    plt.close()

    # === NEW ANALYSIS 2: Sales Velocity Analysis ===
    # Calculate sales velocity (sales per day) over time
    df['Week_Start'] = df['Date'].dt.to_period('W').astype(str)
    
    velocity_metrics = df.groupby(['Week_Start', 'Month']).agg({
        'Price': ['count', 'sum'],
        'Date': 'nunique'
    }).reset_index()
    
    velocity_metrics['Sales_Velocity'] = (
        velocity_metrics[('Price', 'count')] / velocity_metrics[('Date', 'nunique')]
    )
    velocity_metrics['Revenue_Velocity'] = (
        velocity_metrics[('Price', 'sum')] / velocity_metrics[('Date', 'nunique')]
    )
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(15, 12))
    
    sns.lineplot(
        data=velocity_metrics,
        x='Week_Start',
        y='Sales_Velocity',
        hue='Month',
        marker='o',
        ax=ax1
    )
    ax1.set_title('Sales Velocity Trend (Orders per Day)')
    ax1.tick_params(axis='x', rotation=45)
    
    sns.lineplot(
        data=velocity_metrics,
        x='Week_Start',
        y='Revenue_Velocity',
        hue='Month',
        marker='o',
        ax=ax2
    )
    ax2.set_title('Revenue Velocity Trend (Revenue per Day)')
    ax2.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/sales_velocity_analysis.png')
    plt.close()

    # === NEW ANALYSIS 3: Product Mix Evolution ===
    # Analyze how product mix changes over time
    product_mix = df.pivot_table(
        index='Month',
        columns='Category',
        values='Price',
        aggfunc='sum',
        fill_value=0
    )
    
    # Calculate percentage contribution
    product_mix_pct = product_mix.div(product_mix.sum(axis=1), axis=0) * 100
    
    plt.figure(figsize=(12, 6))
    product_mix_pct.plot(kind='area', stacked=True)
    plt.title('Product Category Mix Evolution')
    plt.xlabel('Month')
    plt.ylabel('Percentage of Total Revenue')
    plt.legend(title='Category', bbox_to_anchor=(1.05, 1))
    plt.tight_layout()
    plt.savefig(f'{output_dir}/product_mix_evolution.png')
    plt.show()
    plt.close()

    # === NEW ANALYSIS 4: Customer Response Time Impact ===
    # Analyze impact of day of week and time of day on sales success
    df['DayOfWeek'] = df['Date'].dt.day_name()
    
    # Day of week analysis
    day_metrics = df.groupby('DayOfWeek').agg({
        'Price': ['count', 'sum'],
        'No': 'count'
    }).reset_index()
    
    day_metrics['Conversion_Rate'] = (
        day_metrics[('Price', 'count')] / day_metrics[('No', 'count')] * 100
    )
    
    # Sort by days of week
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    day_metrics['DayOfWeek'] = pd.Categorical(day_metrics['DayOfWeek'], categories=day_order, ordered=True)
    day_metrics = day_metrics.sort_values('DayOfWeek')
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
    sns.barplot(
        data=day_metrics,
        x='DayOfWeek',
        y=('Price', 'sum'),
        ax=ax1
    )
    ax1.set_title('Revenue by Day of Week')
    ax1.tick_params(axis='x', rotation=45)
    
    sns.barplot(
        data=day_metrics,
        x='DayOfWeek',
        y='Conversion_Rate',
        ax=ax2,
        color='green'
    )
    ax2.set_title('Conversion Rate by Day of Week')
    ax2.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(f'{output_dir}/day_of_week_analysis.png')
    plt.show()
    plt.close()

    # === NEW ANALYSIS 5: Cross-Category Purchase Analysis ===
    # Analyze relationships between category purchases
    customer_categories = df[df['Price'] > 0].groupby(['Date', 'Category'])['Price'].sum().unstack(fill_value=0)
    category_correlations = customer_categories.corr()
    
    plt.figure(figsize=(12, 10))
    sns.heatmap(
        category_correlations,
        annot=True,
        cmap='RdYlBu',
        center=0,
        fmt='.2f'
    )
    plt.title('Category Purchase Correlation Matrix')
    plt.tight_layout()
    plt.savefig(f'{output_dir}/category_correlations.png')
    plt.show()
    plt.close()

    # Save additional analysis to Excel
    with pd.ExcelWriter(f'{output_dir}/advanced_analysis_summary.xlsx') as writer:
        category_matrix.to_excel(writer, sheet_name='Category_Matrix')
        velocity_metrics.to_excel(writer, sheet_name='Sales_Velocity')
        product_mix_pct.to_excel(writer, sheet_name='Product_Mix')
        day_metrics.to_excel(writer, sheet_name='Day_Analysis')
        category_correlations.to_excel(writer, sheet_name='Category_Correlations')

    return {
        'category_matrix': category_matrix,
        'velocity_metrics': velocity_metrics,
        'product_mix': product_mix_pct,
        'day_metrics': day_metrics,
        'category_correlations': category_correlations
    }

# Main execution
if __name__ == "__main__":
    # Load data using file paths
    combined_data = load_multiple_months_data(FILE_PATHS)
    
    if combined_data is not None:
        # Create visualizations
        analysis_results = create_enhanced_visualizations(combined_data)
        print("Analysis completed. Check the 'output' directory for results.")
    else:
        print("No data was loaded. Please check the file paths.")
