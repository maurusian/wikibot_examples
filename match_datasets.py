import pandas as pd

#can be either csv or xlsx
df1 = pd.read_csv("dataset1.csv")
df2 = pd.read_excel("dataset2.xlsx")

# Merge dataframes on a common column or column columns
result = pd.merge(df1, df2, on=['col1', 'col2','col3'])

result.to_excel("fully_matched_datasets.xlsx", index=False)
