# Python Automation Portfolio

Python scripts for automating Excel, CSV, and web data tasks — cleaning, merging, analyzing, and reporting.
Built as part of a structured learning plan to offer automation services on Fiverr.

## Portfolio Scripts

### sales_report_generator.py
Takes a folder of messy raw CSV sales files and outputs a clean professional Excel report.
- Reads all CSV files in the folder automatically
- Cleans data — removes duplicates, fixes casing, strips spaces, fills empty cells
- Calculates total sales, average sale, and highest sale per salesperson
- Highlights the top salesperson row in green
- Outputs final_report.xlsx with bold headers and frozen top row
- Never modifies the original CSV files

### pandas_cleaner.py
Cleans any messy CSV file using pandas in under 20 lines.
- Removes duplicate rows
- Fixes inconsistent text casing and strips extra spaces across all text columns
- Fills empty cells with "N/A"
- Reports how many duplicates were removed and how many empty cells were filled
- Saves cleaned data to Excel without modifying the original file

### pandas_report.py
Merges multiple CSV files and generates a sales summary report using pandas.
- Reads and merges all CSV files in the folder automatically
- Cleans and validates data before analysis
- Calculates total, average, and highest sale per salesperson using groupby
- Sorts results by total sales and highlights the top salesperson in green
- Outputs a professional Excel report with bold headers

### price_tracker.py
Scrapes all 1000 books from books.toscrape.com and outputs a professional price analysis report.
- Scrapes 50 pages automatically with respectful request delays
- Cleans and converts price data for analysis
- Sorts all books by price
- Outputs price_report.xlsx with bold headers and a Summary sheet showing total books, average price, cheapest and most expensive book

### cleaner.py
Takes a messy Excel file and returns a clean version automatically.
- Removes duplicate rows
- Strips extra spaces from text
- Fixes inconsistent text casing
- Fills empty cells with "N/A"
- Never overwrites the original file

### merger.py
Merges multiple Excel files from a folder into one master report.
- Automatically finds all .xlsx files in the folder
- Validates that each file has an "Amount" column
- Dynamically detects column positions — works with any column order
- Sorts all rows by Amount (highest first)
- Outputs master_report.xlsx with bold headers

## Practice Scripts

### scraper.py
First web scraping script — scrapes quotes and authors from quotes.toscrape.com and saves to CSV.

### multi_scraper.py
Scrapes all 100 quotes from 10 pages of quotes.toscrape.com and saves to Excel.

### excel_create.py
Creates a new Excel file from scratch with styled headers.

### excel_read.py
Opens an existing Excel file and filters rows based on a condition.

### file_handling_practice.py
Practice script for os and glob — navigating folders and finding files.

### day1.py
First Python script — variables, loops, and functions practice.

## How to Run

1. Install required libraries:
pip install openpyxl pandas requests beautifulsoup4

2. Run any script from its folder:
python sales_report_generator.py
python pandas_report.py
python price_tracker.py
python cleaner.py
python merger.py

## Requirements
- Python 3.x
- openpyxl
- pandas
- requests
- beautifulsoup4

## Progress
- Week 1 complete — Excel automation with openpyxl
- Week 2 complete — pandas data analysis and web scraping
- Week 3 coming — Flask web tools