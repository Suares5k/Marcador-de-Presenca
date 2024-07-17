import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import os

# Função para marcar a presença
def marcar_presenca():
    nome = entry_nome.get()
    if not nome:
        messagebox.showwarning("Atenção", "Por favor, insira o nome completo.")
        return
    
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    lista_presenca.insert(tk.END, f"{nome} - {data_hora}")
    presencas.append({"Nome": nome, "Data e Hora": data_hora})
    entry_nome.delete(0, tk.END)

# Função para salvar a lista de presença em um arquivo Excel
def salvar_presenca():
    if not presencas:
        messagebox.showwarning("Atenção", "Não há presenças para salvar.")
        return
    
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    file_path = os.path.join(desktop_path, "lista_presenca.xlsx")
    
    df = pd.DataFrame(presencas)
    df.to_excel(file_path, index=False)
    
    messagebox.showinfo("Sucesso", f"Lista de presença salva em {file_path}")

# Criação da janela principal
root = tk.Tk()
root.title("Registro de Presença")

presencas = []

# Label e Entry para o nome completo
label_nome = tk.Label(root, text="Nome Completo:")
label_nome.pack(pady=5)
entry_nome = tk.Entry(root, width=50)
entry_nome.pack(pady=5)

# Botão para marcar presença
botao_marcar = tk.Button(root, text="Marcar Presença", command=marcar_presenca)
botao_marcar.pack(pady=10)

# Lista para exibir a presença dos alunos
lista_presenca = tk.Listbox(root, width=75, height=15)
lista_presenca.pack(pady=10)

# Botão para salvar a lista de presença
botao_salvar = tk.Button(root, text="Baixar Lista de Presença", command=salvar_presenca)
botao_salvar.pack(pady=10)

# Inicia o loop principal da interface
root.mainloop()
