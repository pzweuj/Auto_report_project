# coding:utf-8
import pandas as pd

df = pd.read_excel('temp.xlsx', sheetname = 0, header = 0)

df.columns = ['med', 'gene', 'rsid', 'genetype', 'sen', 'pmid', 'level', 'cancer']
df2 = pd.DataFrame(columns=['gene', 'rsid', 'alle', 'genetype'])
df2['gene'] = df['gene']
df2['rsid'] = df['rsid']
df2['genetype'] = df['genetype']

for i in range(df2['genetype'].count()):
    ele = df2['genetype'][i].split(';')
    if ele[0] == ele[1]:
        df2['alle'][i] = '纯合子'
    else:
        df2['alle'][i] = '杂合子'

df3 = df2.sort_values(by='gene').drop_duplicates()

df3.to_excel('out2.xlsx', index=False, sheet_name='rs_check')