import pandas as pd
import glob

# Glob allows us to access all the Excel files in the folder
filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

# Get the input data and load into python
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)

