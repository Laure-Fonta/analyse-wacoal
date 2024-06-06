import pandas as pd

file_path = 'REWORK STOCK.xlsx'
df = pd.read_excel(file_path)

print(df.head())
