import pandas as pd
from pandas import ExcelWriter
df = pd.read_excel(r'C:\Users\MALIKA\Desktop\mapping file.xlsx')
result = df[df['Prod.ERP.Name'].isin(list(df['STATUS'])) & df['Prod.GOV.Name'].isin(list(df['STATUS']))]
print(result)
writer = ExcelWriter('result.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()