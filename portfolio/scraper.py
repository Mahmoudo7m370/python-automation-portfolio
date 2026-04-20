import requests
import pandas as pd
from bs4 import BeautifulSoup

# Fetch the webpage
response = requests.get("https://quotes.toscrape.com")

# Parse the HTML content
soup = BeautifulSoup(response.text, "html.parser")

# Extract all quotes and authors from the page
quotes = soup.find_all("span", class_="text")
authors = soup.find_all("small", class_="author")

# Combine quotes and authors into a list of dictionaries
data = []
for quote, author in zip(quotes, authors):
    data.append({"Author": author.text, "Quote": quote.text})

# Convert to DataFrame and save to CSV
df = pd.DataFrame(data)
df.to_csv("quotes.csv", index=False)
print(f"Scraped {len(df)} quotes.")
  