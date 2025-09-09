import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os


alunos = []


def tirarMedia():
    # Pega nome e nota do aluno
    nome_aluno = entry_nome.get()
    nota_aluno = entry_nota.get()

    #Checa que as caixas não estão vazias
    if not nome_aluno and not nota_aluno:
        messagebox.showwarning("Atenção", "Preencha Nome e Nota do Aluno")
        return

    try:
        nota_aluno = float(nota_aluno)
    except ValueError:
        messagebox.showerror("Erro", "A nota deve ser um número")
        return
    
    #Adiciona a nota do aluno a lista "alunos"
    alunos.append({"Nome": nome_aluno, "Nota": nota_aluno})

    nome_text.insert(tk.END, f"{nome_aluno} - Nota: {nota_aluno}\n")

    # Limpa a caixa de entries
    entry_nome.delete(0, tk.END)
    entry_nota.delete(0, tk.END)

def gerarPlanilha():
    # Caso os entries estejam vazios
    if not alunos:
        messagebox.showwarning("Atenção", "Nenhum aluno cadastrado")

    aprovados, exame, reprovados = [], [], []

    # Checa a nota do aluno e faz a média pra saber se está de exame/aprovado/reprovado
    for aluno in alunos:
        nota = aluno["Nota"]
        if nota >= 5:
            aprovados.append(aluno)
        elif 4 <= nota < 5:
            exame.append(aluno)
        else:
            reprovados.append(aluno)

    #Cria as tabelas Excel e organiza em ordem alfabética
    if aprovados:
        df = pd.DataFrame(aprovados).sort_values(by="Nome")
        df.to_excel("aprovados.xlsx", index=False)

    if exame:
        df = pd.DataFrame(exame).sort_values(by="Nome")
        df.to_excel("exame.xlsx", index=False)

    if reprovados:
        df = pd.DataFrame(reprovados).sort_values(by="Nome")
        df.to_excel("reprovados.xlsx", index=False)

    messagebox.showinfo("Sucesso", "As planilhas foram criadas com sucesso!")


def adicionarPlanilha():
    nome_aluno = entry_nome.get().strip()
    nota_aluno = entry_nota.get().strip()

    if not nome_aluno or not nota_aluno:
        messagebox.showwarning("Atenção", "Preencha nota e nome do aluno")
        return

    try:
        nota_aluno = float(nota_aluno)
    except ValueError:
        messagebox.showerror("Erro", "A nota deve ser um número")
        return

    aluno = {"Nome": nome_aluno, "Nota": nota_aluno}

    if nota_aluno >= 5:
        filename = "aprovados.xlsx"
    elif 4 < nota_aluno < 5:
        filename = "exame.xlsx"
    else:
        filename = "reprovados.xlsx"

    # Faz o check se já existem as planilhas, e caso já existam, concatena elas com os novos dados adicionados
    if os.path.exists(filename):
        df = pd.read_excel(filename)
        df = pd.concat([df, pd.DataFrame([aluno])], ignore_index=True)
    else:
        df = pd.DataFrame([aluno])

    # Organização em ordem alfabética!
    df = df.sort_values(by="Nome")
    df.to_excel(filename, index=False)

    nome_text.insert(tk.END, f"{nome_aluno} - Nota: {nota_aluno} (adicionado direto em {filename})\n")

    entry_nome.delete(0, tk.END)
    entry_nota.delete(0, tk.END)

    messagebox.showinfo("Sucesso", f"Aluno adicionado na planilha {filename}")


root = tk.Tk()
root.title("Média dos alunos")

# Nome e nota dos alunos
label_nome = tk.Label(root, text="Nome do aluno:")
label_nome.grid(row=0, column=0)

label_nota = tk.Label(root, text="Nota do aluno:")
label_nota.grid(row=1, column=0)

# Entrada de dados
entry_nome = tk.Entry(root)
entry_nome.grid(row=0, column=1, padx=5, pady=5)

entry_nota = tk.Entry(root)
entry_nota.grid(row=1, column=1, padx=5, pady=5)

# Apresentação dos dados inseridos
nome_label = tk.Label(root, text="Notas lançadas:")
nome_label.grid(row=2, column=0, columnspan=2, padx=5, pady=(10, 0))

nome_text = tk.Text(root, height=10, width=60)
nome_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

# Botões com as funções
btn_cadastrar_nota = tk.Button(root, text="Cadastrar Nota", command=tirarMedia)
btn_cadastrar_nota.grid(row=4, column=0, padx=5, pady=10, sticky="ew")

btn_gerar_planilhas = tk.Button(root, text="Gerar planilhas", command=gerarPlanilha)
btn_gerar_planilhas.grid(row=4, column=1, padx=5, pady=10, sticky="ew")

btn_adicionar_planilha = tk.Button(root, text="Adicionar nas planilhas", command=adicionarPlanilha)
btn_adicionar_planilha.grid(row=5, column=0, columnspan=2, padx=5, pady=10, sticky="ew")

root.mainloop()
