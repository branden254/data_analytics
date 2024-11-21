import requests
from bs4 import BeautifulSoup
import re
import time
import random
import pandas as pd
from urllib.parse import quote

def google_search(query, num_results=500):
    results = []
    start = 0
    while len(results) < num_results:
        url = f"https://www.google.com/search?q={quote(query)}&num=100&start={start}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        new_results = [div.find('a')['href'] for div in soup.find_all('div', class_='yuRUbf')]
        results.extend(new_results)
        start += 100
        if len(new_results) == 0:
            break
        time.sleep(random.uniform(1, 3))
    return results[:num_results]

def scrape_instagram_profile(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    data = {}
    data['url'] = url
    data['name'] = soup.find('h1', class_='_aacl _aacs _aact _aacx _aada').text if soup.find('h1', class_='_aacl _aacs _aact _aacx _aada') else "N/A"
    
    followers_span = soup.find('span', string=re.compile(r'followers', re.I))
    data['followers'] = followers_span.find_previous('span').text if followers_span else "N/A"
    
    description_div = soup.find('div', class_='-vDIg')
    data['description'] = description_div.text if description_div else "N/A"
    
    return data

def main():
    query = 'site:instagram.com "digital marketing" "gmail"'
    urls = google_search(query, num_results=500)
    
    results = []
    for url in urls:
        if 'instagram.com' in url:
            try:
                data = scrape_instagram_profile(url)
                results.append(data)
                print(f"Scraped: {url}")
            except Exception as e:
                print(f"Error scraping {url}: {str(e)}")
            
            # Add a delay to avoid rate limiting
            time.sleep(random.uniform(1, 3))
    
    # Save results to Excel file
    df = pd.DataFrame(results)
    df.to_excel("instagram_digital_marketers.xlsx", index=False)
    print(f"Results saved to instagram_beauty_influencers.xlsx")

if __name__ == "__main__":
    main()
    