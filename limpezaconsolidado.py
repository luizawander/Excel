import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

def processar_dados():
    try:
        df = pd.read_excel('limpezaconsolidado.xlsx', sheet_name='Limpeza')

        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        df['CPF'] = df['CPF'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
        df['RG'] = df['RG'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)
        df['Contato'] = df['Contato'].map(lambda x: ''.join(filter(str.isdigit, x)) if isinstance(x, str) else x)

        df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce').dt.strftime('%d/%m/%Y')

        cpf_sem_vinculo = df[df['Vínculo'].isna() | (df['Vínculo'] == '') | (df['Vínculo'] == '#N/D')]

        print(cpf_sem_vinculo)

        lista_cpf.delete(*lista_cpf.get_children())

        for index, row in cpf_sem_vinculo.iterrows():
            lista_cpf.insert('', 'end', values=(row['CPF'], row['Vínculo']))

        messagebox.showinfo("Sucesso", f"Processado com sucesso. {len(cpf_sem_vinculo)} CPF(s) sem vínculo encontrados.")
    
    except Exception as e:

        messagebox.showerror("Erro", f"Erro ao processar os dados: {str(e)}")

root = tk.Tk()
root.title("CPFs Sem Vínculo")
root.geometry("400x300")

label_titulo = tk.Label(root, text="Lista de CPFs sem Vínculo", font=("Arial", 14))
label_titulo.pack(pady=10)

colunas = ('CPF', 'Vínculo')
lista_cpf = ttk.Treeview(root, columns=colunas, show='headings')
lista_cpf.heading('CPF', text='CPF')
lista_cpf.heading('Vínculo', text='Vínculo')
lista_cpf.pack(pady=20, fill='x')

btn_processar = tk.Button(root, text="Processar Dados", command=processar_dados)
btn_processar.pack(pady=10)

root.mainloop()
