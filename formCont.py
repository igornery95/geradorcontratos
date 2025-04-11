# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog
from docx import Document
import os

def gerar_contrato(nome, empresa, cpf, endereco, pasta_selecionada):
    """Cria um contrato com base nas informações fornecidas e salva em .docx."""
    doc = Document()
    
    texto_contrato = f"""
    CONTRATO DE PRESTAÇÃO DE SERVIÇOS

    Pelo presente instrumento, {nome}, {empresa} portador do CPF {cpf}, residente no endereço {endereco},
    firma o seguinte contrato nos termos descritos a seguir...
    """
    
    doc.add_paragraph(texto_contrato.strip())
    nome_arquivo = f"contrato_{nome.replace(' ', '_')}.docx"
    caminho_completo = os.path.join(pasta_selecionada, nome_arquivo)
    doc.save(caminho_completo)
    print(f"Contrato salvo em: {caminho_completo}")

def abrir_formulario():
    """Cria e exibe a interface gráfica para preenchimento dos dados."""
    root = tk.Tk()
    root.title("Gerador de Contratos")
    # root.iconbitmap("") #icone

    # Configurar as colunas que devem crescer
    root.columnconfigure(0, weight=1)  
    root.columnconfigure(1, weight=1)  
    root.columnconfigure(2, weight=1)
    root.columnconfigure(3, weight=1)
    root.columnconfigure(4, weight=1)

    # Configurando linhas (vertical)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)
    root.rowconfigure(2, weight=1)
    root.rowconfigure(3, weight=1)
    root.rowconfigure(4, weight=1)

    # Não maximizar
    root.resizable(False, False)

    # Variável para armazenar o caminho da pasta
    pasta_var = tk.StringVar()

    # Labels e campos
    tk.Label(root, text="Nome Completo:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    nome_entry = tk.Entry(root, width=40)
    nome_entry.grid(row=0, column=1, columnspan=3, sticky="nsew", padx=5, pady=5)

    tk.Label(root, text="Empresa:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    emp_entry = tk.Entry(root, width=40)
    emp_entry.grid(row=1, column=1, columnspan=3, sticky="nsew", padx=5, pady=5)

    tk.Label(root, text="CPF/CNPJ:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    cpf_entry = tk.Entry(root, width=40)
    cpf_entry.grid(row=2, column=1, columnspan=3, sticky="nsew", padx=5, pady=5)

    tk.Label(root, text="Endereço:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    endereco_entry = tk.Entry(root, width=40)
    endereco_entry.grid(row=3, column=1, columnspan=3, sticky="nsew", padx=5, pady=5)

    # Botão e label de seleção de pasta
    def selecionar_pasta():
        pasta = filedialog.askdirectory(title="Selecione uma pasta")
        if pasta:
            pasta_var.set(pasta)
            print(f"Pasta selecionada: {pasta}")

    tk.Label(root, text="Pasta destino:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
    tk.Label(root, textvariable=pasta_var, bg="#ffffff", anchor="w", width=30).grid(row=4, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)
    tk.Button(root, text="Selecionar", command=selecionar_pasta).grid(row=4, column=3, sticky="nsew", padx=5, pady=5)

    # Botão salvar
    def processar_formulario():
        nome = nome_entry.get()
        empresa = emp_entry.get()
        cpf = cpf_entry.get()
        endereco = endereco_entry.get()
        pasta = pasta_var.get()

        if nome and empresa and cpf and endereco and pasta:
            gerar_contrato(nome, empresa, cpf, endereco, pasta)
            root.destroy()
        else:
            print("Preencha todos os campos e selecione a pasta.")

    tk.Button(root, text="Gerar Contrato", command=processar_formulario).grid(row=6, column=0, columnspan=4, sticky="nsew")#, pady=10)

    root.mainloop()

if __name__ == "__main__":
    abrir_formulario()
