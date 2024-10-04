import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# Função para carregar o arquivo, processar os dados e exibir CPFs sem vínculo
def processar_dados():
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel('limpezaconsolidado.xlsx', sheet_name='Limpeza')

        # Remover espaços em branco no início e no fim de todas as células que são strings
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        # Limpar CPF, RG e Contato (remover qualquer caractere que não seja dígito)
        df['CPF'] = df['CPF'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
        df['RG'] = df['RG'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
        df['Contato'] = df['Contato'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)

        # Converter Nascimento para formato de data
        df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce').dt.strftime('%d/%m/%Y')

        # Filtrar CPFs sem vínculo
        cpf_sem_vinculo = df[df['Vínculo'].isna() | (df['Vínculo'] == '') | (df['Vínculo'] == '#N/D')]

        # Adicionar impressão para depuração
        print(cpf_sem_vinculo)

        # Limpar a lista antes de exibir os resultados
        lista_cpf.delete(*lista_cpf.get_children())

        # Exibir os CPFs sem vínculo na lista
        for index, row in cpf_sem_vinculo.iterrows():
            lista_cpf.insert('', 'end', values=(row['CPF'], row['Vínculo']))

        # Mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Processado com sucesso. {len(cpf_sem_vinculo)} CPF(s) sem vínculo encontrados.")
    
    except Exception as e:
        # Exibir mensagem de erro caso ocorra algum problema
        messagebox.showerror("Erro", f"Erro ao processar os dados: {str(e)}")

# Criar a interface gráfica
root = tk.Tk()
root.title("CPFs Sem Vínculo")
root.geometry("400x300")

# Label para o título
label_titulo = tk.Label(root, text="Lista de CPFs sem Vínculo", font=("Arial", 14))
label_titulo.pack(pady=10)

# Tabela para exibir os CPFs
colunas = ('CPF', 'Vínculo')
lista_cpf = ttk.Treeview(root, columns=colunas, show='headings')
lista_cpf.heading('CPF', text='CPF')
lista_cpf.heading('Vínculo', text='Vínculo')
lista_cpf.pack(pady=20, fill='x')

# Botão para processar os dados
btn_processar = tk.Button(root, text="Processar Dados", command=processar_dados)
btn_processar.pack(pady=10)

# Iniciar a interface
root.mainloop()
