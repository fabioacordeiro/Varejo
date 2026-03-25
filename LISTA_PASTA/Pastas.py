# -*- coding: utf-8 -*-
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# =========================
# Estilo (dark / moderno)
# =========================
def apply_dark_style(root: tk.Tk) -> ttk.Style:
    root.configure(bg="#0b0b0b")

    style = ttk.Style(root)
    # Em alguns ambientes, "clam" dá mais controle de cor
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("TFrame", background="#0b0b0b")
    style.configure("TLabel", background="#0b0b0b", foreground="#eaeaea")
    style.configure("TNotebook", background="#0b0b0b", borderwidth=0)
    style.configure("TNotebook.Tab", background="#1a1a1a", foreground="#eaeaea", padding=(12, 8))
    style.map("TNotebook.Tab",
              background=[("selected", "#2a2a2a")],
              foreground=[("selected", "#ffffff")])

    style.configure("TButton",
                    background="#2a2a2a",
                    foreground="#ffffff",
                    borderwidth=0,
                    padding=(10, 8))
    style.map("TButton",
              background=[("active", "#3a3a3a")])

    style.configure("TEntry",
                    fieldbackground="#141414",
                    background="#141414",
                    foreground="#eaeaea")

    style.configure("TScrollbar", background="#0b0b0b")

    return style


def choose_directory(initial: str | None = None) -> str:
    return filedialog.askdirectory(initialdir=initial or os.getcwd(), title="Selecione uma pasta") or ""


def list_subfolders(folder: str) -> list[str]:
    p = Path(folder)
    if not p.exists() or not p.is_dir():
        return []
    return sorted([x.name for x in p.iterdir() if x.is_dir()], key=str.lower)


def list_files_grouped_by_subfolder(folder: str) -> list[str]:
    """
    Retorna linhas textuais no formato:
      [SUBPASTA] nome
         arquivo: C:\...\subpasta\arquivo.ext
    Apenas arquivos dentro de CADA subpasta imediata.
    """
    base = Path(folder)
    if not base.exists() or not base.is_dir():
        return []

    lines: list[str] = []
    subdirs = sorted([x for x in base.iterdir() if x.is_dir()], key=lambda x: x.name.lower())

    for sd in subdirs:
        lines.append(f"[SUBPASTA] {sd.name}")
        files = sorted([f for f in sd.rglob("*") if f.is_file()], key=lambda x: str(x).lower())
        if not files:
            lines.append("   (sem arquivos)")
        else:
            for f in files:
                lines.append(f"   {str(f)}")
        lines.append("")  # linha em branco separadora

    if not subdirs:
        lines.append("(não há subpastas nesta pasta)")

    return lines


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Listador de Subpastas e Arquivos")
        self.geometry("980x620")
        self.minsize(860, 520)
        apply_dark_style(self)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=14, pady=14)

        # Notebook (abas)
        nb = ttk.Notebook(container)
        nb.pack(fill="both", expand=True)

        self.tab1 = ttk.Frame(nb)
        self.tab2 = ttk.Frame(nb)
        nb.add(self.tab1, text="Subpastas")
        nb.add(self.tab2, text="Subpastas + Arquivos")

        self._build_tab1()
        self._build_tab2()

    # -------------------------
    # Aba 1
    # -------------------------
    def _build_tab1(self):
        top = ttk.Frame(self.tab1)
        top.pack(fill="x", padx=12, pady=(12, 8))

        ttk.Label(top, text="Pasta base:").pack(side="left")

        self.tab1_path = tk.StringVar(value="")
        path_entry = ttk.Entry(top, textvariable=self.tab1_path)
        path_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        ttk.Button(top, text="Selecionar...", command=self._tab1_pick_folder).pack(side="left")
        ttk.Button(top, text="Gerar lista", command=self._tab1_generate).pack(side="left", padx=(10, 0))

        # Área de saída
        out_frame = ttk.Frame(self.tab1)
        out_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab1_text = tk.Text(
            out_frame,
            bg="#0f0f0f",
            fg="#eaeaea",
            insertbackground="#ffffff",
            relief="flat",
            wrap="none"
        )
        self.tab1_text.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(out_frame, orient="vertical", command=self.tab1_text.yview)
        yscroll.pack(side="right", fill="y")
        self.tab1_text.configure(yscrollcommand=yscroll.set)

    def _tab1_pick_folder(self):
        folder = choose_directory()
        if folder:
            self.tab1_path.set(folder)

    def _tab1_generate(self):
        folder = self.tab1_path.get().strip()
        if not folder:
            messagebox.showwarning("Atenção", "Selecione uma pasta primeiro.")
            return

        subs = list_subfolders(folder)

        self.tab1_text.delete("1.0", "end")
        self.tab1_text.insert("end", f"Pasta: {folder}\n\n")
        if not subs:
            self.tab1_text.insert("end", "(nenhuma subpasta encontrada)\n")
            return

        for name in subs:
            self.tab1_text.insert("end", f"- {name}\n")

    # -------------------------
    # Aba 2
    # -------------------------
    def _build_tab2(self):
        top = ttk.Frame(self.tab2)
        top.pack(fill="x", padx=12, pady=(12, 8))

        ttk.Label(top, text="Pasta base:").pack(side="left")

        self.tab2_path = tk.StringVar(value="")
        path_entry = ttk.Entry(top, textvariable=self.tab2_path)
        path_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        ttk.Button(top, text="Selecionar...", command=self._tab2_pick_folder).pack(side="left")
        ttk.Button(top, text="Gerar lista", command=self._tab2_generate).pack(side="left", padx=(10, 0))

        out_frame = ttk.Frame(self.tab2)
        out_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab2_text = tk.Text(
            out_frame,
            bg="#0f0f0f",
            fg="#eaeaea",
            insertbackground="#ffffff",
            relief="flat",
            wrap="none"
        )
        self.tab2_text.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(out_frame, orient="vertical", command=self.tab2_text.yview)
        yscroll.pack(side="right", fill="y")
        self.tab2_text.configure(yscrollcommand=yscroll.set)

    def _tab2_pick_folder(self):
        folder = choose_directory()
        if folder:
            self.tab2_path.set(folder)

    def _tab2_generate(self):
        folder = self.tab2_path.get().strip()
        if not folder:
            messagebox.showwarning("Atenção", "Selecione uma pasta primeiro.")
            return

        lines = list_files_grouped_by_subfolder(folder)

        self.tab2_text.delete("1.0", "end")
        self.tab2_text.insert("end", f"Pasta: {folder}\n\n")
        self.tab2_text.insert("end", "\n".join(lines).rstrip() + "\n")


if __name__ == "__main__":
    app = App()
    app.mainloop()