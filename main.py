from datetime import datetime
from tkinter import ttk, messagebox
import tkinter as tk
import pandas as pd
import os
import openpyxl
from dateutil.relativedelta import relativedelta

servicos = 'Vencimentos.xlsx'

def calcular_vencimento(data_servico_str):
   try:
    data_servico = datetime.strptime(data_servico_str, '%d/%m/%Y')
    data_vencimento = data_servico + relativedelta(months=3)
    return data_vencimento.strftime('%d/%m/%Y')
   except ValueError:
       return None

def salvar_servico():
    nome_condominio = entry_name.get().strip()
    data_realizacao_str = entry_data.get().strip()

    if not nome_condominio or not data_realizacao_str:
        messagebox.showwarning("Campos Vazios", "Por favor, preencha Nome e Data da Realização.")
        return


    data_vencimento_str = calcular_vencimento(data_realizacao_str)


    if data_vencimento_str is None:
        messagebox.showerror("Erro de data" , "Formato de 'Data realização' inválido. Use DD/MM/AAAA.")
        return


    novo_servico_df = pd.DataFrame({
        'Nome': [nome_condominio],
        'Data realização': [data_realizacao_str],
        'Data vencimento': [data_vencimento_str]
    })

    try:
        df_atualizado = None
        if os.path.exists(servicos):
            df_existente = pd.read_excel(servicos)
            df_atualizado = pd.concat([df_existente, novo_servico_df], ignore_index=True)
        else:
            df_atualizado = novo_servico_df


        with pd.ExcelWriter(servicos, engine='openpyxl') as writer:
            # Salva o DataFrame no writer
            df_atualizado.to_excel(writer, sheet_name='Sheet1', index=False)


            workbook = writer.book
            worksheet = writer.sheets['Sheet1'] 


            worksheet.column_dimensions['A'].width = 30
            worksheet.column_dimensions['B'].width = 18
            worksheet.column_dimensions['C'].width = 18



        messagebox.showinfo("Cadastro realizado", f"Serviço para '{nome_condominio}' cadastrado com sucesso!")
        entry_name.delete(0, tk.END)
        entry_data.delete(0, tk.END)


    except Exception as e:
        messagebox.showerror("Erro ao salvar", f"Ocorreu um erro ao salvar o cadastro: {e}")


root = tk.Tk()
root.title("Cadastro serviços")
root.resizable(False, False)

frm = ttk.Frame(root, padding=10)
frm.grid(sticky=(tk.N, tk.W, tk.E, tk.S))


ttk.Label(frm, text="Cadastro de Serviços e calcular prazos").grid(column=0, row=0, columnspan=2, pady=10)

nome_label = ttk.Label(frm, text="Nome condomínio:")
nome_label.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
entry_name = ttk.Entry(frm, width=50)
entry_name.grid(column=1, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)

data_label = ttk.Label(frm, text="Data Realização (DD/MM/AAAA):")
data_label.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
entry_data = ttk.Entry(frm, width=20)
entry_data.grid(column=1, row=2, sticky=(tk.W, tk.E), padx=5, pady=5)


button_click = ttk.Button(frm, text="Cadastrar Serviço", command=salvar_servico)
button_click.grid(column=0, row=3, columnspan=2, pady=15)


quit_button = ttk.Button(frm, text="Sair", command=root.destroy)
quit_button.grid(column=0, row=4, columnspan=2, pady=5)


root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
frm.columnconfigure(0, weight=0)
frm.columnconfigure(1, weight=1)
frm.rowconfigure(0, weight=0)
frm.rowconfigure(1, weight=0)
frm.rowconfigure(2, weight=0)
frm.rowconfigure(3, weight=0)
frm.rowconfigure(4, weight=0)


root.mainloop()
