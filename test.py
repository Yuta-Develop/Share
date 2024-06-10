import pandas as pd

df1 = pd.read_excel('Data.xlsx',sheet_name ='Data')

df2 = df1.values.tolist()
print(df2)
print('')
print(df2[0])
print('')
print(df2[0][1])





