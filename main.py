import pandas as pd

df = pd.read_excel('limpezaconsolidado.xlsx', sheet_name=0)

df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

df['CPF'] = df['CPF'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['RG'] = df['RG'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
df['Contato'] = df['Contato'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)

df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce').dt.strftime('%d/%m/%Y')

cpf_sem_vinculo = df[df['Vínculo'].isna() | (df['Vínculo'] == '') | (df['Vínculo'] == '#N/D')]

cpf_fora_padrao = df[(df['CPF'].map(lambda x: len(x) != 11 if isinstance(x, str) and x != '' else False)) & df['CPF'].notna()]

print("CPFs sem vínculo:")
print(cpf_sem_vinculo[['CPF', 'Vínculo']].to_string(index=False))

print("\nCPFs fora da padronização (não possuem 11 dígitos):")
print(cpf_fora_padrao[['CPF']].to_string(index=False))

df.to_excel('dadoslimpos.xlsx', index=False, engine='openpyxl')


