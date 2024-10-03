import pandas as pd

df_existing = pd.read_excel('funcionarios_existentes.xlsx')  # Planilha existente
df_new = pd.read_excel('funcionarios_novos.xlsx')  # Nova planilha

duplicates = df_new[df_new['CPF'].isin(df_existing['CPF']) | df_new['RG'].isin(df_existing['RG'])]

print("Dados duplicados encontrados:")
print(duplicates)

duplicates.to_excel('duplicados_encontrados.xlsx', index=False)