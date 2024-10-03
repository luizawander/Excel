import pandas as pd

df = pd.read_excel('limpezaconsolidado.xlsx', sheet_name='Limpeza')

print("Antes da limpeza:")
print(df[['RG', 'CPF']].head(10)) 

df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

df['CPF'] = df['CPF'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['RG'] = df['RG'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)

print(df[['RG', 'CPF']].head(10)) 

df.to_excel('planilhalimpa.xlsx', index=False, engine='openpyxl')
