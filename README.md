# Python Automation Portfolio

Python scripts for automating Excel and CSV tasks — cleaning, merging, and reporting.
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

### excel_create.py
Creates a new Excel file from scratch with styled headers.

### excel_read.py
Opens an existing Excel file and filters rows based on a condition.

### file_handling_practice.py
Practice script for os and glob — navigating folders and finding files.

### day1.py
First Python script — variables, loops, and functions practice.

## How to Run

1. Install the required library:
pip install openpyxl

2. Run any script from its folder:
python sales_report_generator.py
python cleaner.py
python merger.py

## Requirements
- Python 3.x
- openpyxl

## Progress
- Week 1 complete — Excel automation with openpyxl
- Week 2 coming — pandas data analysis and web scraping
