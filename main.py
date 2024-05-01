import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import filedialog
import openpyxl

# Valores pré-definidos para os pesos
pesos_predefinidos = {
    "Administracao:": "1.0",
    "Creche:": "1.0",
    "Pré escola:": "1.0",
    "1 ano:": "1.0",
    "2 ano:": "1.0",
    "3 ano:": "1.0",
    "4 ano:": "1.0",
    "5 ano:": "1.0",
    "6 ano:": "1.0",
    "7 ano:": "1.0",
    "8 ano:": "1.0",
    "9 ano:": "1.0"
}

# Variáveis globais para armazenar os valores dos pesos
pesos_values = pesos_predefinidos.copy()
pesos_entry = {}
pesos_window = None  # Definindo pesos_window globalmente

def definir_pesos():
    global pesos_entry
    global pesos_values
    global pesos_window  # Atribuindo a pesos_window globalmente
    pesos_window = tk.Toplevel(root)
    pesos_window.title("Definir Pesos")

    campos = [
        "Administracao:", "Creche:", "Pré escola:", "1 ano:",
        "2 ano:", "3 ano:", "4 ano:", "5 ano:",
        "6 ano:", "7 ano:", "8 ano:", "9 ano:"
    ]

    pesos_entry = {}

    for i, campo in enumerate(campos):
        label = tk.Label(pesos_window, text=campo)
        label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

        entry = tk.Entry(pesos_window)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entry.insert(0, pesos_values.get(campo, "1.0"))
        pesos_entry[campo] = entry

    confirmar_button = tk.Button(pesos_window, text="Confirmar", command=confirmar_pesos)
    confirmar_button.grid(row=len(campos), columnspan=2, pady=10)

def confirmar_pesos():
    global pesos_values
    global pesos_entry
    global pesos_window  # Atribuindo a pesos_window globalmente

    for campo, entry in pesos_entry.items():
        pesos_values[campo] = entry.get()

    print("Pesos definidos:", pesos_values)

    pesos_window.destroy()

def inserir_caso():
    global entrada_entries
    entrada_window = tk.Toplevel(root)
    entrada_window.title("Inserir Caso de Entrada")
    entrada_entries = []

    campos = [
        "Administracao:", "Creche:", "Pré escola:", "1 ano:",
        "2 ano:", "3 ano:", "4 ano:", "5 ano:",
        "6 ano:", "7 ano:", "8 ano:", "9 ano:"
    ]

    for i, campo in enumerate(campos):
        label = tk.Label(entrada_window, text=campo)
        label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

        if campo == "Administracao:":
            admin_combobox = ttk.Combobox(entrada_window, values=["Federal", "Estadual", "Municipal", "Particular"])
            admin_combobox.grid(row=i, column=1, padx=5, pady=5)
            entrada_entries.append(admin_combobox)
        else:
            entry = tk.Entry(entrada_window)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entrada_entries.append(entry)

    gerar_similaridade_button = tk.Button(entrada_window, text="Gerar Similaridade", command=lambda: gerar_similaridade(entrada_window))
    gerar_similaridade_button.grid(row=len(campos), columnspan=2, pady=10)

def gerar_similaridade(window):
    global pesos_values
    global entrada_entries

    # Lê o arquivo XLSX
    file_path = "Base_de_dados.xlsx"
    if file_path:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        count = 0

        # Aqui você pode fazer os cálculos com todas as linhas e colunas do arquivo XLSX
        # Por exemplo, calcular a similaridade entre os pesos definidos e os dados do arquivo

        # Imprimir as 5 primeiras linhas
        for row in sheet.iter_rows(values_only=True):
            print(row)
            count += 1
            if count > 10:
                break

        workbook.close()
        window.destroy()

root = tk.Tk()
root.title("Programa de Definição de Pesos e Inserção de Caso de Entrada")

botao_definir_pesos = tk.Button(root, text="Definir Pesos", command=definir_pesos)
botao_definir_pesos.pack(pady=10)

botao_inserir_caso = tk.Button(root, text="Inserir Caso de Entrada", command=inserir_caso)
botao_inserir_caso.pack(pady=10)

root.mainloop()
