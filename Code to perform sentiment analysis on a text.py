import pandas as pd
from textblob import TextBlob
from urllib.parse import urlparse
import re

# Function to perform sentiment analysis on a text
def analyze_sentiment(text):
    blob = TextBlob(str(text))  # Ensure text is converted to string
    return blob.sentiment.polarity

# Function to extract company names from titles
def extract_company_names(title):
    company_names = ["Jumia", "M-Pesa", "Huwawei", "Copia"]  # Add more company names as needed
    matches = [company for company in company_names if re.search(r'\b' + re.escape(company) + r'\b', str(title), re.IGNORECASE)]
    return ', '.join(matches) if matches else None

# Function to extract referral sites from URLs
def extract_referral_site(url):
    parsed_url = urlparse(url)
    domain_parts = parsed_url.netloc.split(".")
    if domain_parts[0] == "www" and len(domain_parts) > 1:
        return domain_parts[1]
    elif len(domain_parts) > 0:
        return domain_parts[0]
    else:
        return None

# Read Excel file into a pandas DataFrame
file_path = r"C:\Users\carso\Documents\My Web Sites\Go Pro - Generation Space\genspace full list.xlsx"
df = pd.read_excel(file_path)

# Apply sentiment analysis to Description column
df['Sentiment'] = df['Description'].apply(analyze_sentiment)

# Extract company names from Title column
df['Company_Names'] = df['Title'].apply(extract_company_names)

# Extract referral sites from Detail_URL column
df['Referral_Site'] = df['Detail_URL'].apply(extract_referral_site)

# Convert decimal sentiment scores to categories
def sentiment_category(score):
    if score > 0:
        return "Positive"
    elif score < 0:
        return "Negative"
    else:
        return "Neutral"

df['Sentiment_Category'] = df['Sentiment'].apply(sentiment_category)

# Save the updated DataFrame back to the Excel file
output_file_path = r"C:\Users\carso\Documents\My Web Sites\Go Pro - Generation Space\genspace with analysis.xlsx"
df.to_excel(output_file_path, index=False)
