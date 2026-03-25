import os
import re
import pdfplumber
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =========================
# Funções de extração
# =========================

def extrair_crlv(texto):
    placa = None
    modelo = None

    # Placa (ex: ABC1D23 ou ABC1234)
    placa_match = re.search(r'\b[A-Z]{3}[0-9][A-Z0-9][0-9]{2}\b', texto)
    if placa_match:
        placa = placa_match.group()

    # Modelo (linha após "MARCA/MODELO/VERSAO")
    modelo_match = re.search(r'MARCA/MODELO/VERSAO\s*\n(.+)', texto)
    if modelo_match:
        modelo = modelo_match.group(1).strip()

    return placa, modelo


def extrair_cnh(texto):
    nome = None
    cpf = None

    # Nome (linha após título CNH geralmente)
    nome_match = re.search(r'NOME\s*\n(.+)', texto)
    if nome_match:
        nome = nome_match.group(1).strip()

    # CPF
    cpf_match = re.search(r'\d{3}\.\d{3}\.\d{3}-\d{2}', texto)
    if cpf_match:
        cpf = cpf_match.group()

    return nome, cpf


def ler_pdf(caminho):
    texto_total = ""
    try:
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages:
                texto_total += pagina.extract_text() or ""
    except:
        pass
    return texto_total


def processar_pdfs(pasta, tipo):
    resultados = []

    for raiz, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            if arquivo.lower().endswith(".pdf"):
                caminho = os.path.join(raiz, arquivo)
                texto = ler_pdf(caminho)

                if tipo == "crlv":
                    placa, modelo = extrair_crlv(texto)
                    resultados.append((arquivo, placa, modelo))

                elif tipo == "cnh":
                    nome, cpf = extrair_cnh(texto)
                    resultados.append((arquivo, nome, cpf))

    return resultados


# =========================
# Interface
# =========================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Leitor de Documentos")
        self.root.geometry("900x600")
        self.root.configure(bg="#121212")

        style = ttk.Style()
        style.theme_use("default")

        style.configure("TNotebook", background="#121212", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1e1e1e", foreground="white", padding=10)
        style.map("TNotebook.Tab", background=[("selected", "#333333")])

        style.configure("TButton", background="#333333", foreground="white", padding=10)
        style.map("TButton", background=[("active", "#555555")])

        # Notebook (abas)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        # Abas
        self.aba_crlv = tk.Frame(self.notebook, bg="#121212")
        self.aba_cnh = tk.Frame(self.notebook, bg="#121212")

        self.notebook.add(self.aba_crlv, text="CRLV (Veículos)")
        self.notebook.add(self.aba_cnh, text="CNH (Motoristas)")

        self.criar_aba_crlv()
        self.criar_aba_cnh()

    # =========================
    # ABA CRLV
    # =========================
    def criar_aba_crlv(self):
        self.pasta_crlv = tk.StringVar()

        frame = self.aba_crlv

        tk.Label(frame, text="Selecionar pasta:", bg="#121212", fg="white").pack(pady=10)

        tk.Entry(frame, textvariable=self.pasta_crlv, width=80).pack(pady=5)

        ttk.Button(frame, text="Buscar", command=self.buscar_pasta_crlv).pack(pady=5)
        ttk.Button(frame, text="Processar PDFs", command=self.executar_crlv).pack(pady=10)

        self.texto_crlv = tk.Text(frame, bg="#1e1e1e", fg="white")
        self.texto_crlv.pack(fill="both", expand=True, padx=10, pady=10)

    def buscar_pasta_crlv(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_crlv.set(pasta)

    def executar_crlv(self):
        pasta = self.pasta_crlv.get()
        if not pasta:
            messagebox.showerror("Erro", "Selecione uma pasta")
            return

        resultados = processar_pdfs(pasta, "crlv")

        self.texto_crlv.delete(1.0, tk.END)
        for arquivo, placa, modelo in resultados:
            self.texto_crlv.insert(tk.END, f"{arquivo}\nPlaca: {placa}\nModelo: {modelo}\n\n")

    # =========================
    # ABA CNH
    # =========================
    def criar_aba_cnh(self):
        self.pasta_cnh = tk.StringVar()

        frame = self.aba_cnh

        tk.Label(frame, text="Selecionar pasta:", bg="#121212", fg="white").pack(pady=10)

        tk.Entry(frame, textvariable=self.pasta_cnh, width=80).pack(pady=5)

        ttk.Button(frame, text="Buscar", command=self.buscar_pasta_cnh).pack(pady=5)
        ttk.Button(frame, text="Processar PDFs", command=self.executar_cnh).pack(pady=10)

        self.texto_cnh = tk.Text(frame, bg="#1e1e1e", fg="white")
        self.texto_cnh.pack(fill="both", expand=True, padx=10, pady=10)

    def buscar_pasta_cnh(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_cnh.set(pasta)

    def executar_cnh(self):
        pasta = self.pasta_cnh.get()
        if not pasta:
            messagebox.showerror("Erro", "Selecione uma pasta")
            return

        resultados = processar_pdfs(pasta, "cnh")

        self.texto_cnh.delete(1.0, tk.END)
        for arquivo, nome, cpf in resultados:
            self.texto_cnh.insert(tk.END, f"{arquivo}\nNome: {nome}\nCPF: {cpf}\n\n")


# =========================
# Rodar aplicação
# =========================

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()