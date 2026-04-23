# Import required libraries
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

# Loop through all 10 pages and collect quotes and authors
data = []
for page in range(1, 11):
    url = f"https://quotes.toscrape.com/page/{page}/"
    response = requests.get(url)
    
    # Parse the HTML and extract quotes and authors from current page
    soup = BeautifulSoup(response.text, "html.parser")
    quotes = soup.find_all("span", class_="text")
    authors = soup.find_all("small", class_="author")
    
    # Store each quote and author as a dictionary in the data list
    for quote, author in zip(quotes, authors):
        data.append({"Author": author.text, "Quote": quote.text})
    
    print(f"Scraped page {page}")
    
    # Wait 1 second between requests to avoid being blocked
    time.sleep(1)

# Convert collected data to DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel("quotes_report.xlsx", index=False)
print(f"Total quotes scraped: {len(df)}")

