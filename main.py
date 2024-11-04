import pandas as pd
import glob

# Use glob to access the multiple filepaths
filepaths = glob.glob("invoices/*.xlsx")

# For loop to extract data from the different filepaths
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)