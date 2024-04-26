import pandas as pd

xls = pd.ExcelFile("/home/hydraangel/Downloads/ERF TESTE PEDRO/01 - Fachada Frontal/Fachada Frontal rev.Kelly.xls")

df1 = pd.read_excel(xls, 'Sheet1')
df2 = pd.read_excel(xls, 'Sheet2')

print(df1)
print(df2)
