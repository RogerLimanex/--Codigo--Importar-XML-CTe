import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import uuid
from datetime import datetime

class CTeProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de XML - CTe")

        # Variáveis para armazenar os diretórios selecionados
        self.input_dir = tk.StringVar()
        self.success_dir = tk.StringVar()
        self.error_dir = tk.StringVar()
        self.output_excel = tk.StringVar()

        self.setup_ui()  # Configura a interface gráfica

    def setup_ui(self):
        # Criação do layout da interface com labels, campos e botões
        frame = ttk.Frame(self.root, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        # Diretório de entrada
        ttk.Label(frame, text="Diretório de entrada:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.input_dir, width=50).grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.input_dir)).grid(row=0, column=2)

        # Diretório de sucesso
        ttk.Label(frame, text="Diretório de sucesso:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.success_dir, width=50).grid(row=1, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.success_dir)).grid(row=1, column=2)

        # Diretório de erro
        ttk.Label(frame, text="Diretório de erro:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.error_dir, width=50).grid(row=2, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.error_dir)).grid(row=2, column=2)

        # Caminho de saída do Excel
        ttk.Label(frame, text="Salvar planilha como:").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_excel, width=50).grid(row=3, column=1)
        ttk.Button(frame, text="Procurar", command=self.save_file).grid(row=3, column=2)

        # Botões de ação
        ttk.Button(frame, text="Processar", command=self.start_processing).grid(row=4, column=1, sticky="e", pady=10)
        ttk.Button(frame, text="Cancelar", command=self.root.quit).grid(row=4, column=2, sticky="w", pady=10)

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(frame, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=3, pady=5)

        self.progress = ttk.Label(frame, text="")
        self.progress.grid(row=6, column=0, columnspan=3, pady=5)

    def browse_dir(self, var):
        # Seleciona um diretório e armazena no campo correspondente
        selected = filedialog.askdirectory()
        if selected:
            var.set(selected)

    def save_file(self):
        # Seleciona caminho para salvar a planilha Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.output_excel.set(file_path)

    def start_processing(self):
        # Inicia o processamento em uma thread separada
        threading.Thread(target=self.process_xmls).start()

    def process_xmls(self):
        # Lógica principal para processar os arquivos XML
        input_path = self.input_dir.get()
        success_path = self.success_dir.get()
        error_path = self.error_dir.get()
        output_file = self.output_excel.get()

        if not all([input_path, success_path, error_path, output_file]):
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        # Garante que os diretórios de sucesso e erro existam
        os.makedirs(success_path, exist_ok=True)
        os.makedirs(error_path, exist_ok=True)

        files = [f for f in os.listdir(input_path) if f.lower().endswith(".xml")]
        total_files = len(files)
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0

        data = []
        processed = 0

        for idx, file in enumerate(files):
            input_file_path = os.path.join(input_path, file)
            try:
                tree = ET.parse(input_file_path)
                root = tree.getroot()

                # Trata namespaces se existirem
                ns = {'ns': root.tag[root.tag.find('{')+1:root.tag.find('}')]} if '}' in root.tag else {}

                # Função utilitária para extrair texto de um caminho XML
                def get_text_path(path):
                    el = root.find(path, ns) if ns else root.find(path)
                    return el.text.strip() if el is not None and el.text else ''

                # Extrai a chave do CTe
                chave = ''
                inf_cte_elem = root.find('.//ns:infCte', ns) if ns else root.find('.//infCte')
                if inf_cte_elem is not None:
                    chave = inf_cte_elem.attrib.get('Id', '').replace('CTe', '')

                # Converte a data para formato brasileiro
                data_emissao_raw = get_text_path('.//ns:ide/ns:dhEmi')
                data_emissao = ''
                if data_emissao_raw:
                    try:
                        data_emissao = datetime.fromisoformat(data_emissao_raw.replace('Z', '')).strftime('%d/%m/%Y')
                    except Exception:
                        data_emissao = data_emissao_raw

                # Tenta extrair o telefone de dois possíveis caminhos
                telefone = get_text_path('.//ns:dest/ns:enderDest/ns:fone')
                if not telefone:
                    telefone = get_text_path('.//ns:dest/ns:fone')

                # Monta o dicionário com os dados extraídos
                data.append({
                    'Nome do Arquivo': file,
                    'Chave': chave,
                    'Data de Emissão': data_emissao,
                    'Destinatário': get_text_path('.//ns:dest/ns:xNome'),
                    'CNPJ': get_text_path('.//ns:dest/ns:CNPJ'),
                    'IE': get_text_path('.//ns:dest/ns:IE'),
                    'Telefone': telefone,
                    'Cidade': get_text_path('.//ns:dest/ns:enderDest/ns:xMun'),
                    'CEP': get_text_path('.//ns:dest/ns:enderDest/ns:CEP'),
                    'UF/Estado': get_text_path('.//ns:dest/ns:enderDest/ns:UF'),
                    'País': get_text_path('.//ns:dest/ns:enderDest/ns:xPais')
                })

                # Move o arquivo para o diretório de sucesso, renomeando se necessário
                dest_path = os.path.join(success_path, file)
                if os.path.exists(dest_path):
                    base, ext = os.path.splitext(file)
                    dest_path = os.path.join(success_path, f"{base}_{uuid.uuid4().hex[:6]}{ext}")
                os.rename(input_file_path, dest_path)
                processed += 1
            except Exception as e:
                # Move arquivos com erro para diretório de erro
                dest_path = os.path.join(error_path, file)
                if os.path.exists(dest_path):
                    base, ext = os.path.splitext(file)
                    dest_path = os.path.join(error_path, f"{base}_{uuid.uuid4().hex[:6]}{ext}")
                os.rename(input_file_path, dest_path)
            finally:
                # Atualiza barra de progresso
                self.progress_bar["value"] = idx + 1
                self.root.update_idletasks()

        # Exporta para Excel se algum arquivo for processado
        if processed > 0:
            df = pd.DataFrame(data)
            df = df.sort_values(by='UF/Estado')
            df.to_excel(output_file, index=False)
            self.progress.config(text=f"{processed} arquivos processados com sucesso.")
        else:
            self.progress.config(text="Nenhum arquivo processado com sucesso.")
            messagebox.showinfo("Resultado", "Nenhum arquivo processado com sucesso.")

# Executa a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = CTeProcessor(root)
    root.mainloop()
