import pandas as pd
import glob   # we use this library when we hava multiple file, and we save those path inside a list


filepaths = glob.glob('invoices/*.xlsx')   # *.xlsx mean that we import every file in the folder that is xlsx type
print(filepaths)

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")   # when we use multiple files we need to define the key sheet file
    print(df)