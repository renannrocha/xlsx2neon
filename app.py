import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
import os

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")

engine = create_engine(DATABASE_URL)

# Função para ler Excel
def read_excel_tables(file_path, tables_info):
    tables = []
    for info in tables_info:
        df = pd.read_excel(
            file_path,
            sheet_name=info['sheet'],
            header=info['start_row'] - 1
        )
        tables.append(df)
    return tables

# Função para inserir no bancos
def create_and_insert_tables(tables, table_names):
    for df, table_name in zip(tables, table_names):
        df.to_sql(table_name, engine, if_exists='replace', index=False)

# Interface principal
class ExcelToNeonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to NeonDB")
        self.center_window(500, 500) # largura, altura desejadas
        self.root.resizable(False, False)
        self.root.iconbitmap("icone.ico")

        self.file_path = tk.StringVar()
        self.num_tables = tk.IntVar()
        self.table_frames = []

        self.build_ui()

    def build_ui(self):
        # Seletor de arquivo
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10)
        tk.Label(file_frame, text="Arquivo Excel:").pack(side=tk.LEFT)
        tk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT)
        tk.Button(file_frame, text="Selecionar", command=self.select_file).pack(side=tk.LEFT)

        # Número de tabelas
        num_frame = tk.Frame(self.root)
        num_frame.pack(pady=5)
        tk.Label(num_frame, text="Número de tabelas:").pack(side=tk.LEFT)
        tk.Entry(num_frame, textvariable=self.num_tables, width=5).pack(side=tk.LEFT)
        tk.Button(num_frame, text="Confirmar", command=self.create_table_inputs).pack(side=tk.LEFT)

        # Scrollable container para os campos de tabelas
        container_frame = tk.Frame(self.root)
        container_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(container_frame)
        scrollbar = tk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        self.tables_container = tk.Frame(canvas)

        self.tables_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.tables_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Botão processar
        tk.Button(self.root, text="Processar e Enviar para o Banco", command=self.process).pack(pady=10)
    
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")


    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_path.set(path)

    def create_table_inputs(self):
        # Apaga entradas antigas
        for item in self.table_frames:
            item["frame"].destroy()
        self.table_frames.clear()

        for i in range(self.num_tables.get()):
            frame = tk.LabelFrame(self.tables_container, text=f"Tabela {i+1}")
            frame.pack(padx=10, pady=5, fill="x")

            sheet_var = tk.StringVar()
            start_row_var = tk.IntVar()
            table_name_var = tk.StringVar()

            tk.Label(frame, text="Nome/Índice da planilha:").grid(row=0, column=0, sticky="e")
            tk.Entry(frame, textvariable=sheet_var).grid(row=0, column=1)

            tk.Label(frame, text="Linha do cabeçalho:").grid(row=1, column=0, sticky="e")
            tk.Entry(frame, textvariable=start_row_var).grid(row=1, column=1)

            tk.Label(frame, text="Nome da tabela no banco:").grid(row=2, column=0, sticky="e")
            tk.Entry(frame, textvariable=table_name_var).grid(row=2, column=1)

            self.table_frames.append({
                "frame": frame,
                "sheet": sheet_var,
                "start_row": start_row_var,
                "table_name": table_name_var
            })


    def process(self):
        try:
            path = self.file_path.get()
            if not path:
                messagebox.showerror("Erro", "Selecione um arquivo Excel.")
                return

            tables_info = []
            table_names = []

            for item in self.table_frames:
                sheet = item["sheet"].get()
                start_row = item["start_row"].get()
                table_name = item["table_name"].get()

                if sheet == "" or start_row == 0 or table_name == "":
                    messagebox.showerror("Erro", "Preencha todos os campos das tabelas.")
                    return

                tables_info.append({
                    "sheet": int(sheet) if sheet.isdigit() else sheet,
                    "start_row": start_row
                })
                table_names.append(table_name)

            tables = read_excel_tables(path, tables_info)
            create_and_insert_tables(tables, table_names)
            messagebox.showinfo("Sucesso", "Dados enviados com sucesso ao NeonDB!")

        except Exception as e:
            messagebox.showerror("Erro ao processar", str(e))

# Executar app
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToNeonApp(root)
    root.mainloop()
