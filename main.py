import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import pandas as pd

# Função para processar os arquivos
def processar_xmls():
    pasta_entrada = entrada_pasta.get()
    pasta_processados = entrada_processados.get()
    pasta_nao_processados = entrada_nao_processados.get()

    if not (pasta_entrada and pasta_processados and pasta_nao_processados):
        messagebox.showerror("Erro", "Por favor, selecione todos os diretórios.")
        return

    dados = []
    ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}

    for arquivo in os.listdir(pasta_entrada):
        if arquivo.endswith('.xml'):
            caminho_arquivo = os.path.join(pasta_entrada, arquivo)
            try:
                tree = ET.parse(caminho_arquivo)
                root = tree.getroot()

                ide = root.find('.//ns:ide', ns)
                remetente = root.find('.//ns:rem/ns:infCteDest/ns:dest/ns:xNome', ns)
                valor_total = root.find('.//ns:vPrest', ns)
                chave = root.attrib.get('Id', '')[-44:] if root.attrib.get('Id') else ''

                dados.append({
                    'Arquivo': arquivo,
                    'Chave': chave,
                    'Data de Emissão': ide.findtext('ns:dEmi', default='', namespaces=ns) if ide is not None else '',
                    'Número do CTe': ide.findtext('ns:nCT', default='', namespaces=ns) if ide is not None else '',
                    'Nome do Destinatário': remetente.text if remetente is not None else '',
                    'Valor Total': valor_total.findtext('ns:vTPrest', default='', namespaces=ns) if valor_total is not None else ''
                })

                shutil.move(caminho_arquivo, os.path.join(pasta_processados, arquivo))
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")
                shutil.move(caminho_arquivo, os.path.join(pasta_nao_processados, arquivo))

    if dados:
        df = pd.DataFrame(dados)
        caminho_excel = os.path.join(pasta_processados, 'dados_cte.xlsx')
        df.to_excel(caminho_excel, index=False)
        messagebox.showinfo("Sucesso", f"Processamento concluído!\nArquivo salvo em:\n{caminho_excel}")
    else:
        messagebox.showwarning("Aviso", "Nenhum arquivo processado com sucesso.")

# Função para escolher diretórios
def escolher_diretorio(campo):
    caminho = filedialog.askdirectory()
    if caminho:
        campo.delete(0, tk.END)
        campo.insert(0, caminho)

# Interface gráfica
janela = tk.Tk()
janela.title("Processador de XML CTe")

tk.Label(janela, text="Pasta com XMLs:").grid(row=0, column=0, sticky="e")
entrada_pasta = tk.Entry(janela, width=50)
entrada_pasta.grid(row=0, column=1)
tk.Button(janela, text="Selecionar", command=lambda: escolher_diretorio(entrada_pasta)).grid(row=0, column=2)

tk.Label(janela, text="Pasta para processados:").grid(row=1, column=0, sticky="e")
entrada_processados = tk.Entry(janela, width=50)
entrada_processados.grid(row=1, column=1)
tk.Button(janela, text="Selecionar", command=lambda: escolher_diretorio(entrada_processados)).grid(row=1, column=2)

tk.Label(janela, text="Pasta para não processados:").grid(row=2, column=0, sticky="e")
entrada_nao_processados = tk.Entry(janela, width=50)
entrada_nao_processados.grid(row=2, column=1)
tk.Button(janela, text="Selecionar", command=lambda: escolher_diretorio(entrada_nao_processados)).grid(row=2, column=2)

tk.Button(janela, text="Processar XMLs", command=processar_xmls, bg="#4CAF50", fg="white", width=20).grid(row=3, column=1, pady=10)

janela.mainloop()
