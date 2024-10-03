import pandas as pd

df = pd.read_excel('limpezaconsolidado.xlsx', sheet_name='Limpeza')

df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)

df['CPF'] = df['CPF'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['RG'] = df['RG'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['Contato'] = df['Contato'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce').dt.strftime('%d/%m/%Y')

cpf_sem_vinculo = df[df['Vínculo'].isna() | (df['Vínculo'] == '') | (df['Vínculo'] == '#N/D')]

print(cpf_sem_vinculo[['CPF', 'Vínculo']].to_string(index=False))

df.to_excel('planilhalimpa.xlsx', index=False, engine='openpyxl')
