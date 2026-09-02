"""
app.py — Interface gráfica principal (Tkinter).
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List, Optional

import layout_store
import mapper


# ─── Paleta de cores ──────────────────────────────────────────────────────────
BG_DARK    = "#1e1e2e"
BG_PANEL   = "#2a2a3e"
BG_CARD    = "#313145"
ACCENT     = "#7c6af7"
ACCENT_HVR = "#9b8dff"
SUCCESS    = "#4caf87"
DANGER     = "#e06c75"
TEXT_PRI   = "#e8e8f0"
TEXT_SEC   = "#8888aa"
BORDER     = "#44445a"
MAPPED_BG  = "#2d3b4f"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mapeador de Planilhas")
        self.geometry("1100x700")
        self.minsize(900, 600)
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        # Estado
        self.source_path: Optional[str] = None
        self.dest_path: Optional[str] = None
        self.source_cols: List[str] = []
        self.dest_cols: List[str] = []
        self.source_sheet_idx: int = 0
        self.dest_sheet_idx: int = 0

        # Mapeamentos: {dest_col -> src_col}
        self.mappings: Dict[str, str] = {}

        # Coluna de origem selecionada aguardando destino
        self._pending_src: Optional[str] = None

        self._build_styles()
        self._build_ui()
        self._refresh_layout_list()

    # ─── Estilos ttk ──────────────────────────────────────────────────────────
    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=BG_DARK, foreground=TEXT_PRI,
                         font=("Segoe UI", 10))

        style.configure("TButton", background=ACCENT, foreground="#ffffff",
                         padding=(10, 6), relief="flat", borderwidth=0)
        style.map("TButton",
                  background=[("active", ACCENT_HVR), ("pressed", "#6455d4")])

        style.configure("Danger.TButton", background=DANGER, foreground="#ffffff",
                         padding=(8, 5), relief="flat", borderwidth=0)
        style.map("Danger.TButton",
                  background=[("active", "#f08090"), ("pressed", "#c05060")])

        style.configure("Success.TButton", background=SUCCESS, foreground="#ffffff",
                         padding=(12, 8), relief="flat", borderwidth=0,
                         font=("Segoe UI", 11, "bold"))
        style.map("Success.TButton",
                  background=[("active", "#5dcf9f"), ("pressed", "#3a9f70")])

        style.configure("Ghost.TButton", background=BG_PANEL, foreground=TEXT_SEC,
                         padding=(8, 5), relief="flat", borderwidth=0)
        style.map("Ghost.TButton",
                  background=[("active", BG_CARD)])

        style.configure("TLabel", background=BG_DARK, foreground=TEXT_PRI)
        style.configure("Sec.TLabel", background=BG_DARK, foreground=TEXT_SEC,
                         font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=BG_DARK, foreground=TEXT_PRI,
                         font=("Segoe UI", 15, "bold"))
        style.configure("Card.TLabel", background=BG_PANEL, foreground=TEXT_PRI)
        style.configure("Section.TLabel", background=BG_PANEL, foreground=TEXT_SEC,
                         font=("Segoe UI", 8, "bold"))

        style.configure("TCombobox", fieldbackground=BG_CARD, background=BG_CARD,
                         foreground=TEXT_PRI, selectbackground=ACCENT,
                         selectforeground="#ffffff")

        style.configure("TProgressbar", troughcolor=BG_CARD, background=ACCENT,
                         thickness=6)

        style.configure("TSeparator", background=BORDER)

    # ─── Layout principal ──────────────────────────────────────────────────────
    def _build_ui(self):
        # Cabeçalho
        header = tk.Frame(self, bg=BG_PANEL, height=60)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        tk.Label(header, text="🗂  Mapeador de Planilhas", bg=BG_PANEL,
                 fg=TEXT_PRI, font=("Segoe UI", 14, "bold")).pack(
            side="left", padx=20, pady=15)
        tk.Label(header, text="Mapeie colunas entre planilhas e salve seus layouts",
                 bg=BG_PANEL, fg=TEXT_SEC, font=("Segoe UI", 9)).pack(
            side="left", pady=15)

        # Área principal
        main = tk.Frame(self, bg=BG_DARK)
        main.pack(fill="both", expand=True, padx=0, pady=0)

        # Painel esquerdo — Layouts
        left = tk.Frame(main, bg=BG_PANEL, width=220)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_left_panel(left)

        # Separador
        ttk.Separator(main, orient="vertical").pack(side="left", fill="y")

        # Painel direito — Conteúdo
        right = tk.Frame(main, bg=BG_DARK)
        right.pack(side="left", fill="both", expand=True)
        self._build_right_panel(right)

        # Barra de status inferior
        status_bar = tk.Frame(self, bg=BG_PANEL, height=36)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        self.status_var = tk.StringVar(value="Pronto")
        tk.Label(status_bar, textvariable=self.status_var, bg=BG_PANEL,
                 fg=TEXT_SEC, font=("Segoe UI", 9)).pack(
            side="left", padx=16, pady=8)
        self.progress = ttk.Progressbar(status_bar, mode="determinate",
                                         style="TProgressbar", length=200)
        self.progress.pack(side="right", padx=16, pady=14)

    # ─── Painel esquerdo ───────────────────────────────────────────────────────
    def _build_left_panel(self, parent: tk.Frame):
        tk.Label(parent, text="LAYOUTS SALVOS", bg=BG_PANEL, fg=TEXT_SEC,
                 font=("Segoe UI", 8, "bold")).pack(
            anchor="w", padx=16, pady=(16, 6))

        # Lista de layouts
        list_frame = tk.Frame(parent, bg=BG_PANEL)
        list_frame.pack(fill="both", expand=True, padx=8)

        scrollbar = tk.Scrollbar(list_frame, bg=BG_PANEL, troughcolor=BG_PANEL,
                                  relief="flat", borderwidth=0)
        scrollbar.pack(side="right", fill="y")

        self.layout_listbox = tk.Listbox(
            list_frame,
            bg=BG_CARD, fg=TEXT_PRI,
            selectbackground=ACCENT, selectforeground="#ffffff",
            font=("Segoe UI", 10),
            borderwidth=0, highlightthickness=0,
            relief="flat",
            yscrollcommand=scrollbar.set,
            activestyle="none",
            cursor="hand2",
        )
        self.layout_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.layout_listbox.yview)
        self.layout_listbox.bind("<<ListboxSelect>>", self._on_layout_select)

        # Botões de layout
        btn_frame = tk.Frame(parent, bg=BG_PANEL)
        btn_frame.pack(fill="x", padx=8, pady=10)

        ttk.Button(btn_frame, text="💾 Salvar Layout",
                   command=self._save_layout, style="TButton").pack(
            fill="x", pady=(0, 4))
        ttk.Button(btn_frame, text="🗑 Deletar", style="Danger.TButton",
                   command=self._delete_layout).pack(fill="x")

    # ─── Painel direito ────────────────────────────────────────────────────────
    def _build_right_panel(self, parent: tk.Frame):
        # Botão de execução e opções fixados no rodapé
        exec_frame = tk.Frame(parent, bg=BG_DARK)
        exec_frame.pack(side="bottom", fill="x", padx=16, pady=12)

        # Flag para importar fórmulas/funções do Excel
        self.import_formulas_var = tk.BooleanVar(value=False)
        self.formulas_check = tk.Checkbutton(
            exec_frame,
            text="fx  Importar funções/fórmulas",
            variable=self.import_formulas_var,
            bg=BG_DARK,
            fg=TEXT_PRI,
            selectcolor=BG_CARD,
            activebackground=BG_DARK,
            activeforeground=TEXT_PRI,
            font=("Segoe UI", 9),
            cursor="hand2",
        )
        self.formulas_check.pack(side="left")

        ttk.Button(exec_frame, text="▶  Preencher Planilha",
                   style="Success.TButton",
                   command=self._execute).pack(side="right")
        ttk.Button(exec_frame, text="⚡  Relacionar Automaticamente",
                   style="TButton",
                   command=self._auto_map).pack(side="right", padx=(0, 8))
        ttk.Button(exec_frame, text="↺  Limpar Mapeamento",
                   style="Ghost.TButton",
                   command=self._clear_mapping).pack(side="right", padx=(0, 8))

        # Seção de arquivos
        files_frame = tk.Frame(parent, bg=BG_PANEL)
        files_frame.pack(side="top", fill="x", padx=16, pady=(16, 0))

        self._build_file_section(files_frame)

        ttk.Separator(parent, orient="horizontal").pack(
            side="top", fill="x", padx=16, pady=12)

        # Instrução de mapeamento
        self.instruction_label = tk.Label(
            parent,
            text="① Selecione as planilhas acima para começar o mapeamento",
            bg=BG_DARK, fg=TEXT_SEC, font=("Segoe UI", 9, "italic"))
        self.instruction_label.pack(side="top", anchor="w", padx=16)

        # Área de mapeamento
        mapping_area = tk.Frame(parent, bg=BG_DARK)
        mapping_area.pack(side="top", fill="both", expand=True, padx=16, pady=8)

        self._build_mapping_area(mapping_area)

    # ─── Seção de arquivos ─────────────────────────────────────────────────────
    def _build_file_section(self, parent: tk.Frame):
        parent.columnconfigure(1, weight=1)

        # Origem
        tk.Label(parent, text="Planilha Origem:", bg=BG_PANEL,
                 fg=TEXT_SEC, font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=(8, 8), pady=4)

        self.source_label = tk.Label(parent, text="Nenhum arquivo selecionado",
                                      bg=BG_CARD, fg=TEXT_SEC,
                                      font=("Segoe UI", 9), anchor="w",
                                      padx=8, pady=6, relief="flat")
        self.source_label.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(parent, text="📂 Abrir", command=self._pick_source).grid(
            row=0, column=2, padx=(6, 0), pady=4)

        self.source_sheet_var = tk.StringVar()
        self.source_sheet_combo = ttk.Combobox(
            parent, textvariable=self.source_sheet_var, state="disabled",
            width=18, font=("Segoe UI", 9))
        self.source_sheet_combo.grid(row=0, column=3, padx=(6, 8), pady=4)
        self.source_sheet_combo.bind("<<ComboboxSelected>>",
                                      lambda e: self._on_sheet_change("source"))

        # Destino
        tk.Label(parent, text="Planilha Destino:", bg=BG_PANEL,
                 fg=TEXT_SEC, font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", padx=(8, 8), pady=4)

        self.dest_label = tk.Label(parent, text="Nenhum arquivo selecionado",
                                    bg=BG_CARD, fg=TEXT_SEC,
                                    font=("Segoe UI", 9), anchor="w",
                                    padx=8, pady=6, relief="flat")
        self.dest_label.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Button(parent, text="📂 Abrir", command=self._pick_dest).grid(
            row=1, column=2, padx=(6, 0), pady=4)

        self.dest_sheet_var = tk.StringVar()
        self.dest_sheet_combo = ttk.Combobox(
            parent, textvariable=self.dest_sheet_var, state="disabled",
            width=18, font=("Segoe UI", 9))
        self.dest_sheet_combo.grid(row=1, column=3, padx=(6, 8), pady=4)
        self.dest_sheet_combo.bind("<<ComboboxSelected>>",
                                    lambda e: self._on_sheet_change("dest"))

    # ─── Área de mapeamento ────────────────────────────────────────────────────
    def _build_mapping_area(self, parent: tk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(2, weight=1)
        parent.rowconfigure(1, weight=1)

        # Cabeçalhos das colunas
        tk.Label(parent, text="ORIGEM", bg=BG_DARK, fg=TEXT_SEC,
                 font=("Segoe UI", 8, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        tk.Label(parent, text="", bg=BG_DARK).grid(row=0, column=1)
        tk.Label(parent, text="DESTINO", bg=BG_DARK, fg=TEXT_SEC,
                 font=("Segoe UI", 8, "bold")).grid(
            row=0, column=2, sticky="w", pady=(0, 4))

        # Lista de colunas origem
        src_frame = tk.Frame(parent, bg=BG_CARD, bd=0)
        src_frame.grid(row=1, column=0, sticky="nsew")

        src_scroll = tk.Scrollbar(src_frame, bg=BG_CARD, troughcolor=BG_CARD,
                                   relief="flat", borderwidth=0)
        src_scroll.pack(side="right", fill="y")

        self.src_listbox = tk.Listbox(
            src_frame, bg=BG_CARD, fg=TEXT_PRI,
            selectbackground=ACCENT, selectforeground="#fff",
            font=("Segoe UI", 10),
            borderwidth=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            relief="flat", activestyle="none", cursor="hand2",
            yscrollcommand=src_scroll.set,
        )
        self.src_listbox.pack(side="left", fill="both", expand=True)
        src_scroll.config(command=self.src_listbox.yview)
        self.src_listbox.bind("<<ListboxSelect>>", self._on_src_select)

        # Seta central e botão rápido
        arrow_frame = tk.Frame(parent, bg=BG_DARK, width=80)
        arrow_frame.grid(row=1, column=1, padx=8)
        arrow_frame.pack_propagate(False)
        self.arrow_label = tk.Label(arrow_frame, text="➡", bg=BG_DARK,
                                     fg=TEXT_SEC, font=("Segoe UI", 18))
        self.arrow_label.pack(expand=True)
        tk.Label(arrow_frame, text="clique\norig.→dest.", bg=BG_DARK,
                 fg=TEXT_SEC, font=("Segoe UI", 7)).pack()
        ttk.Button(arrow_frame, text="⚡ Auto", style="TButton",
                   command=self._auto_map).pack(pady=(6, 8))

        # Lista de colunas destino
        dst_frame = tk.Frame(parent, bg=BG_CARD, bd=0)
        dst_frame.grid(row=1, column=2, sticky="nsew")

        dst_scroll = tk.Scrollbar(dst_frame, bg=BG_CARD, troughcolor=BG_CARD,
                                   relief="flat", borderwidth=0)
        dst_scroll.pack(side="right", fill="y")

        self.dst_listbox = tk.Listbox(
            dst_frame, bg=BG_CARD, fg=TEXT_PRI,
            selectbackground=SUCCESS, selectforeground="#fff",
            font=("Segoe UI", 10),
            borderwidth=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=SUCCESS,
            relief="flat", activestyle="none", cursor="hand2",
            yscrollcommand=dst_scroll.set,
        )
        self.dst_listbox.pack(side="left", fill="both", expand=True)
        dst_scroll.config(command=self.dst_listbox.yview)
        self.dst_listbox.bind("<<ListboxSelect>>", self._on_dst_select)

        # Painel de mapeamentos ativos (abaixo das listas)
        tk.Label(parent, text="MAPEAMENTOS ATIVOS", bg=BG_DARK, fg=TEXT_SEC,
                 font=("Segoe UI", 8, "bold")).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(10, 4))
        parent.rowconfigure(3, weight=1)

        mapped_container = tk.Frame(parent, bg=BG_DARK)
        mapped_container.grid(row=3, column=0, columnspan=3, sticky="nsew")

        self.mapped_canvas = tk.Canvas(
            mapped_container, bg=BG_DARK, highlightthickness=0, bd=0)
        mapped_scroll = tk.Scrollbar(
            mapped_container, orient="vertical", command=self.mapped_canvas.yview,
            bg=BG_CARD, troughcolor=BG_DARK, relief="flat", borderwidth=0)
        self.mapped_canvas.configure(yscrollcommand=mapped_scroll.set)

        mapped_scroll.pack(side="right", fill="y")
        self.mapped_canvas.pack(side="left", fill="both", expand=True)

        self.mapping_canvas = tk.Frame(self.mapped_canvas, bg=BG_DARK)
        self._mapped_window = self.mapped_canvas.create_window(
            (0, 0), window=self.mapping_canvas, anchor="nw")

        def _on_inner_configure(e):
            self.mapped_canvas.configure(
                scrollregion=self.mapped_canvas.bbox("all"))

        def _on_canvas_configure(e):
            self.mapped_canvas.itemconfig(self._mapped_window, width=e.width)

        self.mapping_canvas.bind("<Configure>", _on_inner_configure)
        self.mapped_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            self.mapped_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(e):
            self.mapped_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(e):
            self.mapped_canvas.unbind_all("<MouseWheel>")

        self.mapped_canvas.bind("<Enter>", _bind_mousewheel)
        self.mapped_canvas.bind("<Leave>", _unbind_mousewheel)
        self.mapping_canvas.bind("<Enter>", _bind_mousewheel)
        self.mapping_canvas.bind("<Leave>", _unbind_mousewheel)

    # ─── Lógica de seleção de arquivo ─────────────────────────────────────────
    def _pick_source(self):
        path = filedialog.askopenfilename(
            title="Selecionar Planilha de Origem",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not path:
            return
        self.source_path = path
        self.source_label.config(text=path.split("/")[-1].split("\\")[-1],
                                  fg=TEXT_PRI)
        sheets = mapper.get_sheet_names(path)
        self.source_sheet_combo.config(values=sheets, state="readonly")
        self.source_sheet_combo.set(sheets[0])
        self.source_sheet_idx = 0
        self._load_source_cols()

    def _pick_dest(self):
        path = filedialog.askopenfilename(
            title="Selecionar Planilha de Destino",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not path:
            return
        self.dest_path = path
        self.dest_label.config(text=path.split("/")[-1].split("\\")[-1],
                                fg=TEXT_PRI)
        sheets = mapper.get_sheet_names(path)
        self.dest_sheet_combo.config(values=sheets, state="readonly")
        self.dest_sheet_combo.set(sheets[0])
        self.dest_sheet_idx = 0
        self._load_dest_cols()

    def _on_sheet_change(self, side: str):
        if side == "source":
            self.source_sheet_idx = self.source_sheet_combo.current()
            self._load_source_cols()
        else:
            self.dest_sheet_idx = self.dest_sheet_combo.current()
            self._load_dest_cols()

    def _load_source_cols(self):
        if not self.source_path:
            return
        try:
            cols, _ = mapper.get_columns(self.source_path, self.source_sheet_idx)
            self.source_cols = cols
            self._refresh_src_list()
            self._update_instruction()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler origem:\n{e}")

    def _load_dest_cols(self):
        if not self.dest_path:
            return
        try:
            cols, _ = mapper.get_columns(self.dest_path, self.dest_sheet_idx)
            self.dest_cols = cols
            self._refresh_dst_list()
            self._update_instruction()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler destino:\n{e}")

    def _update_instruction(self):
        if self.source_cols and self.dest_cols:
            self.instruction_label.config(
                text="② Clique em uma coluna de ORIGEM, depois em uma de DESTINO para mapear",
                fg=TEXT_PRI)
        elif self.source_cols:
            self.instruction_label.config(
                text="Selecione a planilha de DESTINO", fg=TEXT_SEC)
        elif self.dest_cols:
            self.instruction_label.config(
                text="Selecione a planilha de ORIGEM", fg=TEXT_SEC)

    # ─── Lógica de mapeamento ──────────────────────────────────────────────────
    def _refresh_src_list(self):
        mapped_srcs = set(self.mappings.values())
        self.src_listbox.delete(0, tk.END)
        for col in self.source_cols:
            self.src_listbox.insert(tk.END, f"  {col}")
            if col in mapped_srcs:
                self.src_listbox.itemconfig(tk.END, fg=ACCENT)

    def _refresh_dst_list(self):
        self.dst_listbox.delete(0, tk.END)
        for col in self.dest_cols:
            self.dst_listbox.insert(tk.END, f"  {col}")
            if col in self.mappings:
                self.dst_listbox.itemconfig(tk.END, fg=SUCCESS)

    def _refresh_mapping_display(self):
        # Limpa painel
        for widget in self.mapping_canvas.winfo_children():
            widget.destroy()

        if not self.mappings:
            tk.Label(self.mapping_canvas, text="Nenhum mapeamento definido ainda",
                     bg=BG_DARK, fg=TEXT_SEC,
                     font=("Segoe UI", 9, "italic")).pack(anchor="w")
            return

        for dst_col, src_col in self.mappings.items():
            row = tk.Frame(self.mapping_canvas, bg=MAPPED_BG, pady=3, padx=8)
            row.pack(fill="x", pady=2)

            tk.Label(row, text=src_col, bg=MAPPED_BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), width=20, anchor="w").pack(
                side="left")
            tk.Label(row, text="──▶", bg=MAPPED_BG, fg=TEXT_SEC,
                     font=("Segoe UI", 9)).pack(side="left", padx=6)
            tk.Label(row, text=dst_col, bg=MAPPED_BG, fg=SUCCESS,
                     font=("Segoe UI", 9, "bold"), width=20, anchor="w").pack(
                side="left")

            # Botão de remover
            dst = dst_col  # captura para closure
            tk.Button(row, text="✕", bg=MAPPED_BG, fg=DANGER,
                      font=("Segoe UI", 8, "bold"), relief="flat",
                      borderwidth=0, cursor="hand2",
                      command=lambda d=dst: self._remove_mapping(d)).pack(
                side="right", padx=4)

    def _on_src_select(self, event):
        sel = self.src_listbox.curselection()
        if not sel:
            return
        col = self.source_cols[sel[0]]
        self._pending_src = col
        self.arrow_label.config(fg=ACCENT, text="➡")
        self.status_var.set(f"Origem selecionada: '{col}' — agora clique em uma coluna de DESTINO")

    def _on_dst_select(self, event):
        sel = self.dst_listbox.curselection()
        if not sel:
            return
        dst_col = self.dest_cols[sel[0]]

        if self._pending_src is None:
            self.status_var.set("Selecione uma coluna de ORIGEM primeiro")
            return

        # Registra mapeamento
        self.mappings[dst_col] = self._pending_src
        self._pending_src = None
        self.arrow_label.config(fg=TEXT_SEC)
        self.status_var.set(f"✓ Mapeamento adicionado: '{list(self.mappings.keys())[-1]}'")

        self._refresh_src_list()
        self._refresh_dst_list()
        self._refresh_mapping_display()
        self.src_listbox.selection_clear(0, tk.END)
        self.dst_listbox.selection_clear(0, tk.END)

    def _remove_mapping(self, dst_col: str):
        self.mappings.pop(dst_col, None)
        self._refresh_src_list()
        self._refresh_dst_list()
        self._refresh_mapping_display()
        self.status_var.set(f"Mapeamento para '{dst_col}' removido")

    def _auto_map(self):
        if not self.source_cols or not self.dest_cols:
            messagebox.showwarning(
                "Atenção",
                "Selecione as planilhas de origem e destino antes de relacionar as colunas automaticamente."
            )
            return

        matches = mapper.auto_match_columns(self.source_cols, self.dest_cols)
        if not matches:
            messagebox.showinfo(
                "Auto Relacionamento",
                "Nenhuma correspondência automática foi encontrada entre as colunas."
            )
            return

        # Atualiza o dicionário de mapeamentos com as relações encontradas
        self.mappings.update(matches)
        self._pending_src = None
        self.arrow_label.config(fg=TEXT_SEC)

        self._refresh_src_list()
        self._refresh_dst_list()
        self._refresh_mapping_display()
        self.src_listbox.selection_clear(0, tk.END)
        self.dst_listbox.selection_clear(0, tk.END)

        count = len(matches)
        self.status_var.set(f"✓ {count} coluna(s) relacionada(s) automaticamente!")

    def _clear_mapping(self):
        if self.mappings and messagebox.askyesno(
                "Limpar", "Tem certeza que deseja limpar todos os mapeamentos?"):
            self.mappings.clear()
            self._pending_src = None
            self._refresh_src_list()
            self._refresh_dst_list()
            self._refresh_mapping_display()
            self.status_var.set("Mapeamentos limpos")

    # ─── Layouts ───────────────────────────────────────────────────────────────
    def _refresh_layout_list(self):
        self.layout_listbox.delete(0, tk.END)
        for name in layout_store.list_layouts():
            self.layout_listbox.insert(tk.END, f"  {name}")

    def _on_layout_select(self, event):
        sel = self.layout_listbox.curselection()
        if not sel:
            return
        raw = self.layout_listbox.get(sel[0]).strip()
        rules = layout_store.load_layout(raw)
        if rules is None:
            return
        self.mappings = {r["destination"]: r["source"] for r in rules}
        self._refresh_src_list()
        self._refresh_dst_list()
        self._refresh_mapping_display()
        self.status_var.set(f"Layout '{raw}' carregado")

    def _save_layout(self):
        if not self.mappings:
            messagebox.showwarning("Aviso", "Não há mapeamentos para salvar.")
            return
        name = simpledialog.askstring(
            "Salvar Layout", "Nome do layout:",
            parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        rules = [{"source": src, "destination": dst}
                 for dst, src in self.mappings.items()]
        layout_store.save_layout(name, rules)
        self._refresh_layout_list()
        self.status_var.set(f"Layout '{name}' salvo com sucesso!")

    def _delete_layout(self):
        sel = self.layout_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Selecione um layout para deletar.")
            return
        raw = self.layout_listbox.get(sel[0]).strip()
        if messagebox.askyesno("Deletar", f"Deletar o layout '{raw}'?"):
            layout_store.delete_layout(raw)
            self._refresh_layout_list()
            self.status_var.set(f"Layout '{raw}' deletado")

    # ─── Execução ──────────────────────────────────────────────────────────────
    def _execute(self):
        if not self.source_path:
            messagebox.showwarning("Atenção", "Selecione a planilha de origem.")
            return
        if not self.dest_path:
            messagebox.showwarning("Atenção", "Selecione a planilha de destino.")
            return
        if not self.mappings:
            messagebox.showwarning("Atenção", "Defina pelo menos um mapeamento de coluna.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Salvar Planilha Preenchida",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="planilha_preenchida.xlsx")
        if not output_path:
            return

        rules = [{"source": src, "destination": dst}
                 for dst, src in self.mappings.items()]
        import_formulas = self.import_formulas_var.get()

        self.progress["value"] = 0
        self.status_var.set("Executando preenchimento…")
        self.update_idletasks()

        def run():
            try:
                def cb(current, total):
                    pct = int(current / total * 100) if total else 100
                    self.progress["value"] = pct
                    self.status_var.set(f"Copiando linha {current}/{total}…")
                    self.update_idletasks()

                total = mapper.execute_mapping(
                    source_path=self.source_path,
                    dest_path=self.dest_path,
                    output_path=output_path,
                    mapping=rules,
                    source_sheet=self.source_sheet_idx,
                    dest_sheet=self.dest_sheet_idx,
                    import_formulas=import_formulas,
                    progress_callback=cb,
                )
                self.progress["value"] = 100
                self.status_var.set(
                    f"✓ Concluído! {total} linhas copiadas → {output_path.split('/')[-1].split(chr(92))[-1]}")
                messagebox.showinfo(
                    "Concluído",
                    f"Planilha preenchida com sucesso!\n{total} linhas copiadas.\n\nSalvo em:\n{output_path}")
            except Exception as e:
                self.status_var.set(f"Erro: {e}")
                messagebox.showerror("Erro", f"Falha ao preencher:\n{e}")

        threading.Thread(target=run, daemon=True).start()
