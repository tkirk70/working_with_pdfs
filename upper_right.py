# Trying Tabula-Py

# Creates an excel file for sql inserts into quickbooks invoice

import tabula

# Read pdf into a list of DataFrame
dfs = tabula.read_pdf("upper_right.pdf", pages='all')

import pandas as pd

nme = []
for item in dfs[1]['Vendor']:
    nme.append(item)

qty = []
for item in dfs[1]['Ship To']:
    qty.append(item)

rte = []
for item in dfs[1]['Unnamed: 3']:
    rte.append(item)

po = []
for item in dfs[0]["P.O. No."]:
    po.append(item)

ntz = []
for item in dfs[1]['Account #']:
    ntz.append(item)

dict1 = {'Item': nme, 'Quantity': qty, 'Rate': rte}

df = pd.DataFrame(dict1)
df.dropna(inplace=True)

# Add PO and Delivery Notes to df.
df['PO'] = dfs[0]["P.O. No."][1]
df['Notes'] = dfs[1]['Account #'][9]

df.to_excel(f'current_upperright_po3.xlsx', index=False)