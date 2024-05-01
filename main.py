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

        # Criar Treeview para exibir as linhas da planilha em formato de tabela
        tree = ttk.Treeview(window)
        tree.grid(row=len(entrada_entries) + 1, columnspan=2, padx=5, pady=5, sticky="nsew")

        # Configurar as colunas
        tree["columns"] = sheet[1]
        tree.heading("#0", text="ID")
        tree.column("#0", width=50, stretch=False)
        tree.heading("#1", text="Ano")
        tree.heading("#2", text="Cód. Mun.")
        tree.heading("#3", text="Município")
        tree.heading("#4", text="Cód. INEP")
        tree.heading("#5", text="Nome da Escola")
        tree.heading("#6", text="Dep. Adm.")
        tree.heading("#7", text="Classe de Alfabetização")
        tree.heading("#8", text="Creche")
        tree.heading("#9", text="Pré escola")
        tree.heading("#10", text="1 ano")
        tree.heading("#11", text="2 ano")
        tree.heading("#12", text="3 ano")
        tree.heading("#13", text="4 ano")
        tree.heading("#14", text="5 ano")
        tree.heading("#15", text="6 ano")
        tree.heading("#16", text="7 ano")
        tree.heading("#17", text="8 ano")
        tree.heading("#18", text="9 ano")

        for col, title in enumerate(sheet[1], start=1):
            #tree.heading(f"#{col}", text=title)
            tree.column(f"#{col}", width=100)  # Definindo largura padrão para as colunas

        # Adicionar linhas
        for row_data in sheet.iter_rows(min_row=2, values_only=True):
            tree.insert("", "end", text=count, values=row_data)
            count += 1

        # Adicionar barra de rolagem horizontal
        hscroll = ttk.Scrollbar(window, orient="horizontal", command=tree.xview)
        hscroll.grid(row=len(entrada_entries) + 2, column=0, columnspan=2, sticky="ew")
        tree.configure(xscrollcommand=hscroll.set)

        # Definir largura máxima da janela da Treeview
        tree_width = min(800, sum([100 for _ in sheet[1]]))  # Defina a largura máxima aqui
        tree_width += 50  # Adicionar espaço extra para a barra de rolagem
        window.geometry(f"{tree_width}x400")

        workbook.close()

root = tk.Tk()
root.title("Programa de Definição de Pesos e Inserção de Caso de Entrada")

botao_definir_pesos = tk.Button(root, text="Definir Pesos", command=definir_pesos)
botao_definir_pesos.pack(pady=10)

botao_inserir_caso = tk.Button(root, text="Inserir Caso de Entrada", command=inserir_caso)
botao_inserir_caso.pack(pady=10)

root.mainloop()
