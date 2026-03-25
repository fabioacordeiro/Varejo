import os
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


class CSVParaExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV para Excel")
        self.root.geometry("980x700")
        self.root.configure(bg="#0d0d0d")
        self.root.minsize(900, 620)

        self.pasta_selecionada = tk.StringVar()
        self.arquivo_saida = tk.StringVar()

        self.configurar_estilo()
        self.criar_interface()

    def configurar_estilo(self):
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "TFrame",
            background="#0d0d0d"
        )

        style.configure(
            "Card.TFrame",
            background="#141414",
            relief="flat"
        )

        style.configure(
            "Title.TLabel",
            background="#0d0d0d",
            foreground="#ffffff",
            font=("Segoe UI", 22, "bold")
        )

        style.configure(
            "SubTitle.TLabel",
            background="#0d0d0d",
            foreground="#a8a8a8",
            font=("Segoe UI", 11)
        )

        style.configure(
            "Label.TLabel",
            background="#141414",
            foreground="#f2f2f2",
            font=("Segoe UI", 10)
        )

        style.configure(
            "Info.TLabel",
            background="#141414",
            foreground="#bfbfbf",
            font=("Segoe UI", 10)
        )

        style.configure(
            "Status.TLabel",
            background="#0d0d0d",
            foreground="#7CFC98",
            font=("Segoe UI", 10, "bold")
        )

        style.configure(
            "Modern.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=10,
            borderwidth=0,
            relief="flat",
            background="#00b894",
            foreground="#ffffff"
        )

        style.map(
            "Modern.TButton",
            background=[
                ("active", "#00d1a7"),
                ("disabled", "#3a3a3a")
            ],
            foreground=[
                ("disabled", "#888888")
            ]
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=10,
            borderwidth=0,
            relief="flat",
            background="#1f1f1f",
            foreground="#ffffff"
        )

        style.map(
            "Secondary.TButton",
            background=[
                ("active", "#2b2b2b"),
                ("disabled", "#3a3a3a")
            ],
            foreground=[
                ("disabled", "#888888")
            ]
        )

        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#1e1e1e",
            background="#00b894",
            bordercolor="#1e1e1e",
            lightcolor="#00b894",
            darkcolor="#00b894"
        )

    def criar_interface(self):
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Conversor de CSV para Excel",
            style="Title.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            container,
            text="Selecione uma pasta, leia todos os CSVs das subpastas e gere um único arquivo Excel.",
            style="SubTitle.TLabel"
        ).pack(anchor="w", pady=(4, 18))

        card = ttk.Frame(container, style="Card.TFrame", padding=20)
        card.pack(fill="x", pady=(0, 15))

        ttk.Label(card, text="Pasta de origem", style="Label.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )

        self.entry_pasta = tk.Entry(
            card,
            textvariable=self.pasta_selecionada,
            bg="#1e1e1e",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief="flat",
            font=("Segoe UI", 10)
        )
        self.entry_pasta.grid(row=1, column=0, sticky="ew", padx=(0, 10), ipady=10)

        self.btn_pasta = ttk.Button(
            card,
            text="Selecionar Pasta",
            style="Secondary.TButton",
            command=self.selecionar_pasta
        )
        self.btn_pasta.grid(row=1, column=1, sticky="ew")

        ttk.Label(card, text="Arquivo Excel de saída", style="Label.TLabel").grid(
            row=2, column=0, sticky="w", pady=(18, 6)
        )

        self.entry_saida = tk.Entry(
            card,
            textvariable=self.arquivo_saida,
            bg="#1e1e1e",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief="flat",
            font=("Segoe UI", 10)
        )
        self.entry_saida.grid(row=3, column=0, sticky="ew", padx=(0, 10), ipady=10)

        self.btn_saida = ttk.Button(
            card,
            text="Salvar Como",
            style="Secondary.TButton",
            command=self.selecionar_saida
        )
        self.btn_saida.grid(row=3, column=1, sticky="ew")

        card.columnconfigure(0, weight=1)

        botoes_frame = ttk.Frame(container, style="TFrame")
        botoes_frame.pack(fill="x", pady=(0, 15))

        self.btn_processar = ttk.Button(
            botoes_frame,
            text="Processar Arquivos",
            style="Modern.TButton",
            command=self.iniciar_processamento
        )
        self.btn_processar.pack(side="left")

        self.btn_limpar = ttk.Button(
            botoes_frame,
            text="Limpar Log",
            style="Secondary.TButton",
            command=self.limpar_log
        )
        self.btn_limpar.pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(
            container,
            mode="determinate",
            style="Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(0, 8))

        self.status_label = ttk.Label(
            container,
            text="Aguardando seleção da pasta...",
            style="Status.TLabel"
        )
        self.status_label.pack(anchor="w", pady=(0, 12))

        log_card = ttk.Frame(container, style="Card.TFrame", padding=15)
        log_card.pack(fill="both", expand=True)

        ttk.Label(log_card, text="Log de processamento", style="Label.TLabel").pack(anchor="w", pady=(0, 8))

        self.txt_log = tk.Text(
            log_card,
            bg="#0f0f0f",
            fg="#f1f1f1",
            insertbackground="#ffffff",
            relief="flat",
            wrap="word",
            font=("Consolas", 10)
        )
        self.txt_log.pack(fill="both", expand=True, side="left")

        scrollbar = tk.Scrollbar(log_card, command=self.txt_log.yview, bg="#1e1e1e")
        scrollbar.pack(side="right", fill="y")
        self.txt_log.config(yscrollcommand=scrollbar.set)

        self.log("Programa iniciado com sucesso.")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos CSV")
        if pasta:
            self.pasta_selecionada.set(pasta)

            if not self.arquivo_saida.get().strip():
                nome_padrao = os.path.join(pasta, "resultado_consolidado.xlsx")
                self.arquivo_saida.set(nome_padrao)

            self.log(f"Pasta selecionada: {pasta}")
            self.status("Pasta selecionada com sucesso.")

    def selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel como",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if arquivo:
            self.arquivo_saida.set(arquivo)
            self.log(f"Arquivo de saída definido: {arquivo}")
            self.status("Arquivo de saída definido.")

    def limpar_log(self):
        self.txt_log.delete("1.0", tk.END)
        self.log("Log limpo.")

    def log(self, mensagem):
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)
        self.root.update_idletasks()

    def status(self, mensagem):
        self.status_label.config(text=mensagem)
        self.root.update_idletasks()

    def habilitar_controles(self, habilitar=True):
        estado = "normal" if habilitar else "disabled"
        self.btn_pasta.config(state=estado)
        self.btn_saida.config(state=estado)
        self.btn_processar.config(state=estado)
        self.btn_limpar.config(state=estado)

    def iniciar_processamento(self):
        pasta = self.pasta_selecionada.get().strip()
        saida = self.arquivo_saida.get().strip()

        if not pasta:
            messagebox.showwarning("Atenção", "Selecione uma pasta de origem.")
            return

        if not os.path.isdir(pasta):
            messagebox.showerror("Erro", "A pasta selecionada não existe.")
            return

        if not saida:
            messagebox.showwarning("Atenção", "Informe o arquivo Excel de saída.")
            return

        if not saida.lower().endswith(".xlsx"):
            saida += ".xlsx"
            self.arquivo_saida.set(saida)

        self.habilitar_controles(False)
        self.progress["value"] = 0
        self.status("Processando arquivos CSV...")

        thread = threading.Thread(target=self.processar_csvs, args=(pasta, saida), daemon=True)
        thread.start()

    def encontrar_arquivos_csv(self, pasta):
        arquivos_csv = []
        for raiz, _, arquivos in os.walk(pasta):
            for arquivo in arquivos:
                if arquivo.lower().endswith(".csv"):
                    arquivos_csv.append(os.path.join(raiz, arquivo))
        return arquivos_csv

    def ler_csv_flexivel(self, caminho):
        encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
        separadores = [None, ";", ",", "\t", "|"]

        ultimo_erro = None

        for encoding in encodings:
            for sep in separadores:
                try:
                    if sep is None:
                        df = pd.read_csv(
                            caminho,
                            encoding=encoding,
                            sep=None,
                            engine="python",
                            dtype=str
                        )
                    else:
                        df = pd.read_csv(
                            caminho,
                            encoding=encoding,
                            sep=sep,
                            dtype=str
                        )

                    if df is not None and len(df.columns) > 0:
                        return df, encoding, sep if sep is not None else "auto"
                except Exception as e:
                    ultimo_erro = e

        raise Exception(f"Falha ao ler o arquivo: {caminho}\nErro: {ultimo_erro}")

    def processar_csvs(self, pasta, saida):
        try:
            arquivos_csv = self.encontrar_arquivos_csv(pasta)

            if not arquivos_csv:
                self.root.after(0, lambda: self.finalizar_erro("Nenhum arquivo CSV foi encontrado na pasta e subpastas."))
                return

            self.root.after(0, lambda: self.log(f"Total de arquivos CSV encontrados: {len(arquivos_csv)}"))

            dataframes = []
            total = len(arquivos_csv)

            for i, arquivo in enumerate(arquivos_csv, start=1):
                try:
                    self.root.after(0, lambda a=arquivo: self.status(f"Lendo arquivo: {os.path.basename(a)}"))

                    df, encoding_usado, sep_usado = self.ler_csv_flexivel(arquivo)

                    df["ARQUIVO_ORIGEM"] = os.path.basename(arquivo)
                    df["CAMINHO_ORIGEM"] = arquivo

                    dataframes.append(df)

                    self.root.after(
                        0,
                        lambda a=arquivo, e=encoding_usado, s=sep_usado, linhas=len(df):
                        self.log(f"[OK] {a} | encoding={e} | separador={s} | linhas={linhas}")
                    )

                except Exception as e:
                    self.root.after(0, lambda a=arquivo, err=str(e): self.log(f"[ERRO] {a}\n{err}\n"))

                progresso = (i / total) * 100
                self.root.after(0, lambda p=progresso: self.progress.configure(value=p))

            if not dataframes:
                self.root.after(0, lambda: self.finalizar_erro("Os arquivos CSV foram encontrados, mas nenhum pôde ser lido com sucesso."))
                return

            self.root.after(0, lambda: self.status("Consolidando dados..."))
            df_final = pd.concat(dataframes, ignore_index=True, sort=False)

            self.root.after(0, lambda: self.log(f"Total de linhas consolidadas: {len(df_final)}"))
            self.root.after(0, lambda: self.log(f"Total de colunas no consolidado: {len(df_final.columns)}"))

            self.root.after(0, lambda: self.status("Gerando arquivo Excel..."))

            with pd.ExcelWriter(saida, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Consolidado")

            self.root.after(0, lambda: self.progress.configure(value=100))
            self.root.after(0, lambda: self.finalizar_sucesso(saida, len(arquivos_csv), len(df_final)))

        except Exception as e:
            erro = f"{str(e)}\n\n{traceback.format_exc()}"
            self.root.after(0, lambda: self.finalizar_erro(erro))

    def finalizar_sucesso(self, saida, total_arquivos, total_linhas):
        self.habilitar_controles(True)
        self.status("Processamento finalizado com sucesso.")
        self.log("")
        self.log("========================================")
        self.log("PROCESSAMENTO CONCLUÍDO COM SUCESSO")
        self.log(f"Arquivo Excel gerado: {saida}")
        self.log(f"Arquivos CSV processados: {total_arquivos}")
        self.log(f"Total de linhas consolidadas: {total_linhas}")
        self.log("========================================")
        messagebox.showinfo(
            "Sucesso",
            f"Processamento concluído com sucesso.\n\n"
            f"Arquivo gerado:\n{saida}\n\n"
            f"Arquivos CSV processados: {total_arquivos}\n"
            f"Total de linhas consolidadas: {total_linhas}"
        )

    def finalizar_erro(self, erro):
        self.habilitar_controles(True)
        self.status("Ocorreu um erro durante o processamento.")
        self.log("")
        self.log("========================================")
        self.log("ERRO NO PROCESSAMENTO")
        self.log(erro)
        self.log("========================================")
        messagebox.showerror("Erro", erro)


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVParaExcelApp(root)
    root.mainloop()