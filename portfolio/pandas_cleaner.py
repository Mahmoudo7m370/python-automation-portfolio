import pandas as pd

# Read the CSV file into a DataFrame
df = pd.read_csv("sales_q1.csv")

# Remove duplicate rows
df = df.drop_duplicates()

# Clean all text columns — strip spaces and fix casing
for col in df.columns:
    try:
        df[col] = df[col].str.strip().str.title()
    except:
        pass

# Count empty cells before filling
empty_cells = df.isnull().sum().sum()

# Fill empty cells with "N/A"
df = df.fillna("N/A")

# Count removed duplicates
duplicates = len(pd.read_csv("sales_q1.csv")) - len(df)

# Save cleaned data to a new Excel file without row numbers
df.to_excel("cleaned_output.xlsx", index=False)

# Print summary
print(f"Removed {duplicates} duplicates. Fixed {empty_cells} empty cells.")
print("Done!")