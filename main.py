import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import uuid

class CTeProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de XML - CTe")

        # Diretórios
        self.input_dir = tk.StringVar()
        self.success_dir = tk.StringVar()
        self.error_dir = tk.StringVar()
        self.output_excel = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        # Campo de seleção de diretório de entrada
        ttk.Label(frame, text="Diretório de entrada:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.input_dir, width=50).grid(row=0, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.input_dir)).grid(row=0, column=2)

        # Campo de seleção de diretório de sucesso
        ttk.Label(frame, text="Diretório de sucesso:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.success_dir, width=50).grid(row=1, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.success_dir)).grid(row=1, column=2)

        # Campo de seleção de diretório de erro
        ttk.Label(frame, text="Diretório de erro:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.error_dir, width=50).grid(row=2, column=1)
        ttk.Button(frame, text="Procurar", command=lambda: self.browse_dir(self.error_dir)).grid(row=2, column=2)

        # Campo de seleção de saída do Excel
        ttk.Label(frame, text="Salvar planilha como:").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_excel, width=50).grid(row=3, column=1)
        ttk.Button(frame, text="Procurar", command=self.save_file).grid(row=3, column=2)

        # Botões de ação
        ttk.Button(frame, text="Processar", command=self.start_processing).grid(row=4, column=1, sticky="e", pady=10)
        ttk.Button(frame, text="Cancelar", command=self.root.quit).grid(row=4, column=2, sticky="w", pady=10)

        self.progress = ttk.Label(frame, text="")
        self.progress.grid(row=5, column=0, columnspan=3, pady=5)

    def browse_dir(self, var):
        selected = filedialog.askdirectory()
        if selected:
            var.set(selected)

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.output_excel.set(file_path)

    def start_processing(self):
        threading.Thread(target=self.process_xmls).start()

    def process_xmls(self):
        input_path = self.input_dir.get()
        success_path = self.success_dir.get()
        error_path = self.error_dir.get()
        output_file = self.output_excel.get()

        if not all([input_path, success_path, error_path, output_file]):
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        os.makedirs(success_path, exist_ok=True)
        os.makedirs(error_path, exist_ok=True)

        data = []
        processed = 0

        for file in os.listdir(input_path):
            if file.lower().endswith(".xml"):
                input_file_path = os.path.join(input_path, file)
                try:
                    tree = ET.parse(input_file_path)
                    root = tree.getroot()

                    # Busca considerando namespace
                    ns = {'ns': root.tag[root.tag.find('{')+1:root.tag.find('}')]} if '}' in root.tag else {}

                    inf_cte_elem = root.find('.//ns:infCte', ns) if ns else root.find('.//infCte')
                    chave = inf_cte_elem.attrib.get('Id', '').replace('CTe', '') if inf_cte_elem is not None else ''

                    data_emissao_elem = root.find('.//ns:ide/ns:dhEmi', ns) if ns else root.find('.//ide/dhEmi')
                    data_emissao = data_emissao_elem.text[:10] if data_emissao_elem is not None and data_emissao_elem.text else ''

                    dest_elem = root.find('.//ns:dest/ns:xNome', ns) if ns else root.find('.//dest/xNome')
                    nome_dest = dest_elem.text if dest_elem is not None else ''

                    data.append({
                        'Nome do Arquivo': file,
                        'Chave': chave,
                        'Data de Emissão': data_emissao,
                        'Destinatário': nome_dest
                    })

                    dest_path = os.path.join(success_path, file)
                    if os.path.exists(dest_path):
                        base, ext = os.path.splitext(file)
                        dest_path = os.path.join(success_path, f"{base}_{uuid.uuid4().hex[:6]}{ext}")
                    os.rename(input_file_path, dest_path)

                    processed += 1
                except Exception as e:
                    dest_path = os.path.join(error_path, file)
                    if os.path.exists(dest_path):
                        base, ext = os.path.splitext(file)
                        dest_path = os.path.join(error_path, f"{base}_{uuid.uuid4().hex[:6]}{ext}")
                    os.rename(input_file_path, dest_path)

        if processed > 0:
            df = pd.DataFrame(data)
            df.to_excel(output_file, index=False)
            self.progress.config(text=f"{processed} arquivos processados com sucesso.")
        else:
            self.progress.config(text="Nenhum arquivo processado com sucesso.")
            messagebox.showinfo("Resultado", "Nenhum arquivo processado com sucesso.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CTeProcessor(root)
    root.mainloop()
