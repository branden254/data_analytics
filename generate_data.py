import matplotlib.pyplot as plt
import pandas as pd

# Data for visualizations
competitive_analysis_data = {
    'Date': [
        '1st July', '2nd July', '3rd July', '4th July', '5th July',
        '7th-8th July', '9th July', '10th July', '11th July', '12th July',
        '14th-15th July', '16th July', '17th July', '18th July', '19th July',
        '21st-22nd July', '23rd July', '24th July', '25th July', '26th July',
        '28th-29th July', '30th July', '31st July'
    ],
    'Products Analyzed': [
        40, 264, 390, 372, 387, 244, 518, 351, 275, 275, 562, 0, 4, 3, 3, 0, 300, 300, 2, 3, 0, 0, 0
    ]
}

telecalls_data = {
    'Date': [
        '7th-8th July', '9th July', '10th July', '11th July', '12th July', 
        '14th-15th July', '23rd July', '24th July', '31st July'
    ],
    'Updates': [
        1, 1, 1, 1, 1, 1, 1, 1, 1
    ]
}

google_analytics_data = {
    'Date': [
        '16th July', '23rd July', '24th July', '25th July', '26th July', '30th July', '31st July'
    ],
    'Insights': [
        1, 1, 1, 1, 1, 1, 1
    ]
}

# Creating dataframes
df_competitive_analysis = pd.DataFrame(competitive_analysis_data)
df_telecalls = pd.DataFrame(telecalls_data)
df_google_analytics = pd.DataFrame(google_analytics_data)

# Plotting
fig, axs = plt.subplots(3, 1, figsize=(10, 15))

# Competitive Analysis Breakdown
axs[0].bar(df_competitive_analysis['Date'], df_competitive_analysis['Products Analyzed'], color='skyblue')
axs[0].set_title('Monthly Competitive Analysis Breakdown')
axs[0].set_xlabel('Date')
axs[0].set_ylabel('Products Analyzed')
axs[0].tick_params(axis='x', rotation=90)

# Telecalls Analysis Summary
axs[1].bar(df_telecalls['Date'], df_telecalls['Updates'], color='lightgreen')
axs[1].set_title('Telecalls Analysis Summary')
axs[1].set_xlabel('Date')
axs[1].set_ylabel('Updates')
axs[1].tick_params(axis='x', rotation=90)

# Google Analytics Insights
axs[2].bar(df_google_analytics['Date'], df_google_analytics['Insights'], color='salmon')
axs[2].set_title('Google Analytics Insights')
axs[2].set_xlabel('Date')
axs[2].set_ylabel('Insights')
axs[2].tick_params(axis='x', rotation=90)

plt.tight_layout()
plt.savefig("July_Work_Summary_Visualizations.png")
plt.show()
