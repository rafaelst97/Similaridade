import tkinter as tk
from tkinter import ttk
from tkinter import font


def definir_pesos():
    # Função para definir pesos
    global pesos_entry
    pesos_window = tk.Toplevel(root)
    pesos_window.title("Definir Pesos")

    campos = [
        "Administracao:", "Número de Matrículas", "Creche:", "Pre_escola:", "Primeiro_ano:",
        "Segundo_ano:", "Terceiro_ano:", "Quarto_ano:", "Quinto_ano:",
        "Sexto_ano:", "Setimo_ano:", "Oitavo_ano:", "Nono_ano:"
    ]

    pesos_entry = {}

    for i, campo in enumerate(campos):
        label = tk.Label(pesos_window, text=campo)
        label.grid(row=i, column=0, padx=5, pady=5, sticky="w")  # Ajustando para sticky="w" para alinhar à esquerda

        # Para o campo "Administracao:", usaremos um Combobox
        if campo == "Administracao:":
            admin_combobox = ttk.Combobox(pesos_window, values=["Federal", "Estadual", "Municipal", "Particular"])
            admin_combobox.grid(row=i, column=1, padx=5, pady=5)
            pesos_entry[campo] = admin_combobox
        elif campo == "Número de Matrículas":
            label = tk.Label(pesos_window)
            label.config(font=font.Font(weight="bold", size = 16))  # Aplicando a configuração da fonte ao Label
            label.grid(row=i, column = 1, padx = 5, pady = 5, sticky = "w")
        else:
            entry = tk.Entry(pesos_window)
            entry.grid(row=i, column=1, padx=5, pady=5)
            pesos_entry[campo] = entry


def inserir_caso():
    # Função para inserir caso de entrada
    global entrada_entries
    entrada_window = tk.Toplevel(root)
    entrada_window.title("Inserir Caso de Entrada")
    entrada_entries = []
    for i in range(5):
        label = tk.Label(entrada_window, text=f"Campo {i + 1}:")
        label.grid(row=i, column=0, padx=5, pady=5)
        entrada_entry = tk.Entry(entrada_window)
        entrada_entry.grid(row=i, column=1, padx=5, pady=5)
        entrada_entries.append(entrada_entry)


def gerar_similaridade():
    # Função para gerar similaridade
    pesos = {campo: entry.get() if isinstance(entry, tk.Entry) else entry.get() for campo, entry in pesos_entry.items()}
    caso_entrada = [entry.get() for entry in entrada_entries]

    # Aqui você pode implementar a lógica para calcular a similaridade entre os pesos e o caso de entrada
    # Por enquanto, apenas exibiremos os valores coletados
    print("Pesos:", pesos)
    print("Caso de Entrada:", caso_entrada)


root = tk.Tk()
root.title("Programa de Definição de Pesos e Inserção de Caso de Entrada")

# Botões para definir pesos, inserir caso de entrada e gerar similaridade
botao_definir_pesos = tk.Button(root, text="Definir Pesos", command=definir_pesos)
botao_definir_pesos.pack(pady=10)
botao_inserir_caso = tk.Button(root, text="Inserir Caso de Entrada", command=inserir_caso)
botao_inserir_caso.pack(pady=10)
botao_gerar_similaridade = tk.Button(root, text="Gerar Similaridade", command=gerar_similaridade)
botao_gerar_similaridade.pack(pady=10)

root.mainloop()
