"""Aplicación local de escritorio para Windows 10 y 11."""
import csv
import calendar
import ctypes
import math
import queue
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, font, messagebox, simpledialog, ttk

from core import (DataError, batch_label, check_date, customer_message, delete_selection,
                  entry_to_result, exclude_saved, filter_checks, find_combination, find_fewest, find_progressive,
                  list_sheets, load_registry, money, parse_amount, parse_date, read_excel,
                  registry_path, save_selection, saved_overlap)


# --- PALETA INTERMEDIA SLATE-DARK / GRISES NEUTROS ---
BG_CANVAS    = "#12151B"    # Fondo exterior profundo
BG_PANEL     = "#1A1E26"    # Tarjetas y paneles principales
BG_SURFACE   = "#232934"    # Contenedores internos e inputs
BG_HOVER     = "#2E3644"    # Estados activos / hover
BORDER       = "#363E4F"    # Bordes finos estructurales
BORDER_LIGHT = "#475267"    # Bordes de inputs / divisores

TEXT_HIGH    = "#F1F5F9"    # Texto destacado (Slate 100)
TEXT_MID     = "#CBD5E1"    # Texto estándar de lectura (Slate 300)
TEXT_MUTED   = "#8493A8"    # Etiquetas y descripciones (Slate 400/500)

CYAN         = "#38BDF8"    # Acento primario interactivo (Sky 400)
CYAN_HOVER   = "#0284C7"    # Hover botones principales
CYAN_DIM     = "#16384C"    # Fondo para selecciones e indicadores
EMERALD      = "#34D399"    # Montos en e-cheques (Emerald 400)
AMBER        = "#FBBF24"    # Diferencia transferencia (Amber 400)


class Application:
    def __init__(self, root):
        self.root = root
        self.path = None
        self.imported = None
        self.result = None
        self.busy = False
        self.events = queue.Queue()
        self.cancel = threading.Event()
        self.clients = {}
        self.sheet = tk.StringVar()
        self.client = tk.StringVar(value="Todos los clientes")
        self.target = tk.StringVar()
        self.date_field = tk.StringVar(value="G")
        self.date_from = tk.StringVar()
        self.date_to = tk.StringVar()
        self.strategy = tk.StringVar(value="Progresiva · mayor a menor")
        self.prefer_previous = tk.BooleanVar(value=True)
        self.skip_saved = tk.BooleanVar(value=True)
        self.search_context = {}
        self.file_label = tk.StringVar(value="Cargá un archivo Excel para comenzar")
        self.file_detail = tk.StringVar(value="Formato .xlsx con estructura oficial de e-cheques.")
        self.notice = tk.StringVar(value="Las fechas e identificadores originales se conservan sin modificaciones.")
        self.status = tk.StringVar(value="Listo para iniciar. Seleccioná un archivo .xlsx.")
        self.result_title = tk.StringVar(value="Resultados de la búsqueda")
        self.result_detail = tk.StringVar(value="Primero el monto exacto; si no, el más cercano sin pasarse.")
        self.target_total = tk.StringVar(value="$ —")
        self.check_total = tk.StringVar(value="$ —")
        self.transfer_total = tk.StringVar(value="$ —")
        self.message = tk.StringVar(value="Al calcular, vas a poder copiar el resumen para el cliente.")

        # Controladores internos para animaciones
        self._anim_pulse_step = 0
        self._anim_pulse_timer = None
        self._spinner_angle = 0
        self._spinner_timer = None
        self._result_timers = set()

        self._build()
        for variable in (self.target, self.client, self.date_field, self.date_from,
                         self.date_to, self.strategy, self.prefer_previous, self.skip_saved):
            variable.trace_add("write", self._invalidate)
        self.poll_timer = self.root.after(100, self._poll)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def _build(self):
        root = self.root
        root.title("eCheck · Combinador de pagos")

        # Efecto de transparencia sutil estilo vidrio/mica en Windows
        try:
            root.attributes("-alpha", 0.985)
        except Exception:
            pass

        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        width = min(1280, max(1020, screen_w - 100))
        height = min(940, max(690, screen_h - 90))
        root.geometry(f"{width}x{height}")
        root.minsize(980, 680)
        root.configure(bg=BG_CANVAS)

        try:
            available_fonts = font.families(root)
        except Exception:
            available_fonts = []
        FONT_FAMILY = "Segoe UI Variable Text" if "Segoe UI Variable Text" in available_fonts else "Segoe UI"
        self.FONT_FAMILY = FONT_FAMILY
        root.option_add("*Font", f"{{{FONT_FAMILY}}} 10")

        style = ttk.Style(root)
        style.theme_use("clam")

        # Botón Estándar (Gris intermedio)
        style.configure(
            "TButton",
            padding=(14, 8),
            font=(FONT_FAMILY, 9, "bold"),
            background=BG_SURFACE,
            foreground=TEXT_HIGH,
            borderwidth=1,
            relief="solid",
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER
        )
        style.map(
            "TButton",
            background=[("active", BG_HOVER), ("disabled", BG_PANEL)],
            foreground=[("disabled", TEXT_MUTED)],
            bordercolor=[("active", BORDER_LIGHT), ("disabled", BORDER)]
        )

        # Botón Primario (Acento Cyan)
        style.configure(
            "Primary.TButton",
            padding=(16, 9),
            font=(FONT_FAMILY, 10, "bold"),
            background=CYAN,
            foreground="#08141F",
            borderwidth=0
        )
        style.map(
            "Primary.TButton",
            background=[("active", CYAN_HOVER), ("disabled", "#1E3A4B")],
            foreground=[("disabled", "#52758F")]
        )

        # Botón Detener / Cancelar
        style.configure(
            "Danger.TButton",
            padding=(12, 6),
            font=(FONT_FAMILY, 9),
            background="#3D1D24",
            foreground="#FCA5A5",
            borderwidth=1,
            bordercolor="#682A36"
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#52252F"), ("disabled", BG_PANEL)],
            foreground=[("disabled", TEXT_MUTED)]
        )

        # Selector Desplegable
        style.configure(
            "TCombobox",
            padding=6,
            background=BG_SURFACE,
            fieldbackground=BG_SURFACE,
            foreground=TEXT_HIGH,
            darkcolor=BORDER,
            lightcolor=BORDER,
            bordercolor=BORDER_LIGHT,
            arrowcolor=TEXT_MID
        )
        style.map("TCombobox", fieldbackground=[("readonly", BG_SURFACE)], selectbackground=[("readonly", BG_HOVER)])

        # Tabla de Datos (Treeview)
        style.configure(
            "Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=TEXT_MID,
            rowheight=34,
            borderwidth=0,
            font=(FONT_FAMILY, 9)
        )
        style.configure(
            "Treeview.Heading",
            background=BG_SURFACE,
            foreground=TEXT_HIGH,
            padding=(10, 9),
            font=(FONT_FAMILY, 9, "bold"),
            borderwidth=0,
            relief="flat"
        )
        style.map("Treeview.Heading", background=[("active", BG_HOVER)])
        style.map("Treeview", background=[("selected", CYAN_DIM)], foreground=[("selected", TEXT_HIGH)])

        # Scrollbar estilizado
        style.configure("TScrollbar", background=BG_SURFACE, troughcolor=BG_CANVAS, borderwidth=0, arrowsize=11)

        # --- HEADER PRINCIPAL ---
        header = tk.Frame(root, bg=BG_PANEL, padx=26, pady=15, highlightthickness=1, highlightbackground=BORDER)
        header.pack(fill="x")

        head_left = tk.Frame(header, bg=BG_PANEL)
        head_left.pack(side="left")
        tk.Label(head_left, text="eCheck", bg=BG_PANEL, fg=CYAN, font=(FONT_FAMILY, 18, "bold")).pack(side="left")
        tk.Label(head_left, text="│", bg=BG_PANEL, fg=BORDER_LIGHT, font=(FONT_FAMILY, 15)).pack(side="left", padx=12)
        tk.Label(head_left, text="COMBINADOR DE PAGOS", bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 9, "bold")).pack(side="left")

        badge = tk.Label(header, text="ENTORNO LOCAL", bg=BG_SURFACE, fg=TEXT_MID,
                         font=(FONT_FAMILY, 8, "bold"), padx=10, pady=4, highlightthickness=1, highlightbackground=BORDER)
        badge.pack(side="right")

        # --- CONTENEDOR CENTRAL ---
        main_host = tk.Frame(root, bg=BG_CANVAS)
        main_host.pack(fill="both", expand=True)
        self.main_canvas = tk.Canvas(main_host, bg=BG_CANVAS, highlightthickness=0)
        page_scroll = ttk.Scrollbar(main_host, orient="vertical", command=self.main_canvas.yview)
        page_scroll.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_canvas.configure(yscrollcommand=page_scroll.set)
        main_container = tk.Frame(self.main_canvas, bg=BG_CANVAS, padx=22, pady=16)
        main_window = self.main_canvas.create_window(0, 0, window=main_container, anchor="nw")

        def resize_content(event=None):
            self.main_canvas.itemconfigure(main_window, width=self.main_canvas.winfo_width(),
                height=max(main_container.winfo_reqheight(), self.main_canvas.winfo_height()))
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        main_container.bind("<Configure>", resize_content)
        self.main_canvas.bind("<Configure>", resize_content)
        def page_wheel(event):
            if event.widget is not self.table and event.widget.winfo_toplevel() == root:
                self.main_canvas.yview_scroll(-int(event.delta / 120), "units")
        root.bind("<MouseWheel>", page_wheel)

        # PANELES DE ENTRADA (2 COLUMNAS)
        top_grid = tk.Frame(main_container, bg=BG_CANVAS)
        top_grid.pack(fill="x", pady=(0, 12))
        top_grid.columnconfigure(0, weight=1, uniform="top_cards")
        top_grid.columnconfigure(1, weight=1, uniform="top_cards")

        # TARJETA 1: CARGA DE ARCHIVO
        file_card = tk.Frame(top_grid, bg=BG_PANEL, padx=18, pady=15, highlightthickness=1, highlightbackground=BORDER)
        file_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        fc_head = tk.Frame(file_card, bg=BG_PANEL)
        fc_head.pack(fill="x", pady=(0, 8))
        tk.Label(fc_head, text="1. LISTADO DE ORIGEN", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8, "bold")).pack(side="left")
        self.notes_button = ttk.Button(fc_head, text="Ver detalle", command=self.show_import_notes, state="disabled")
        self.notes_button.pack(side="right")

        tk.Label(file_card, textvariable=self.file_label, bg=BG_PANEL, fg=TEXT_HIGH,
                 font=(FONT_FAMILY, 11, "bold"), anchor="w").pack(fill="x")
        tk.Label(file_card, textvariable=self.file_detail, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 9), anchor="w").pack(fill="x", pady=(2, 10))

        fc_actions = tk.Frame(file_card, bg=BG_PANEL)
        fc_actions.pack(fill="x", pady=(4, 8))
        self.load_button = ttk.Button(fc_actions, text="Abrir Excel…", command=self.load_file, style="Primary.TButton")
        self.load_button.pack(side="left", padx=(0, 10))

        sheet_container = tk.Frame(fc_actions, bg=BG_PANEL)
        sheet_container.pack(side="left", fill="x", expand=True)
        tk.Label(sheet_container, text="Hoja:", bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8, "bold")).pack(anchor="w")
        self.sheet_box = ttk.Combobox(sheet_container, textvariable=self.sheet, state="disabled")
        self.sheet_box.pack(fill="x")
        self.sheet_box.bind("<<ComboboxSelected>>", self._sheet_changed)

        tk.Label(file_card, textvariable=self.notice, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8), anchor="w", wraplength=480, justify="left").pack(fill="x")

        dates_frame = tk.Frame(file_card, bg=BG_PANEL)
        dates_frame.pack(fill="x", pady=(10, 0))
        for col in (1, 2):
            dates_frame.columnconfigure(col, weight=1)
        for col, label in enumerate(("Fecha Excel", "Desde · dd/mm/aaaa", "Hasta · dd/mm/aaaa")):
            tk.Label(dates_frame, text=label, bg=BG_PANEL, fg=TEXT_MUTED,
                     font=(FONT_FAMILY, 8)).grid(row=0, column=col, sticky="w", padx=(0, 6))
        self.date_box = ttk.Combobox(dates_frame, textvariable=self.date_field, values=("F", "G"),
                                     width=4, state="readonly")
        self.date_box.grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.date_entries, self.calendar_buttons = [], []
        for col, variable in ((1, self.date_from), (2, self.date_to)):
            wrapper = tk.Frame(dates_frame, bg=BG_PANEL)
            wrapper.grid(row=1, column=col, sticky="ew", padx=(0, 6))
            entry = ttk.Entry(wrapper, textvariable=variable, width=11)
            entry.pack(side="left", fill="x", expand=True)
            button = ttk.Button(wrapper, text="…", width=2,
                                command=lambda value=variable: self._open_calendar(value))
            button.pack(side="left", padx=(2, 0))
            self.date_entries.append(entry)
            self.calendar_buttons.append(button)
        self.clear_dates_button = tk.Button(dates_frame, text="Limpiar", command=self._clear_dates,
            bg=BG_PANEL, fg=CYAN, activebackground=BG_PANEL, activeforeground=TEXT_HIGH,
            relief="flat", borderwidth=0, font=(FONT_FAMILY, 8), padx=3, pady=0)
        self.clear_dates_button.grid(row=0, column=2, sticky="e")

        # TARJETA 2: IMPORTE Y PARÁMETROS
        entry_card = tk.Frame(top_grid, bg=BG_PANEL, padx=18, pady=15, highlightthickness=1, highlightbackground=BORDER)
        entry_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ec_head = tk.Frame(entry_card, bg=BG_PANEL)
        ec_head.pack(fill="x", pady=(0, 8))
        tk.Label(ec_head, text="2. IMPORTE A CUBRIR", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8, "bold")).pack(side="left")
        self.cancel_button = ttk.Button(ec_head, text="Detener búsqueda", command=self.stop, state="disabled", style="Danger.TButton")
        self.cancel_button.pack(side="right")

        inputs_frame = tk.Frame(entry_card, bg=BG_PANEL)
        inputs_frame.pack(fill="x", pady=(0, 8))

        amount_box = tk.Frame(inputs_frame, bg=BG_PANEL)
        amount_box.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Label(amount_box, text="Monto objetivo ($):", bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8, "bold")).pack(anchor="w")
        self.amount_entry = ttk.Entry(amount_box, textvariable=self.target, font=(FONT_FAMILY, 12, "bold"))
        self.amount_entry.pack(fill="x", pady=(2, 0))
        self.amount_entry.bind("<Return>", lambda event: self.search())

        client_box_wrapper = tk.Frame(inputs_frame, bg=BG_PANEL)
        client_box_wrapper.pack(side="left", fill="x", expand=True)
        tk.Label(client_box_wrapper, text="Cliente:", bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8, "bold")).pack(anchor="w")
        self.client_box = ttk.Combobox(client_box_wrapper, textvariable=self.client, state="disabled")
        self.client_box.pack(fill="x", pady=(2, 0))
        self.client_box.bind("<<ComboboxSelected>>", self._invalidate)

        search_options = tk.Frame(entry_card, bg=BG_PANEL)
        search_options.pack(fill="x", pady=(2, 8))
        self.strategy_box = ttk.Combobox(search_options, textvariable=self.strategy, state="readonly",
            values=("Progresiva · mayor a menor", "Mejor suma posible", "Mejor suma · menos e-cheques"))
        self.strategy_box.pack(fill="x")
        self.previous_toggle = tk.Checkbutton(search_options, text="Priorizar la tanda de una semana anterior",
            variable=self.prefer_previous, bg=BG_PANEL, fg=TEXT_MID, selectcolor=BG_SURFACE,
            activebackground=BG_PANEL, activeforeground=TEXT_HIGH, font=(FONT_FAMILY, 9), anchor="w")
        self.previous_toggle.pack(fill="x", pady=(4, 0))
        self.saved_toggle = tk.Checkbutton(search_options, text="Excluir e-cheques ya guardados en el registro",
            variable=self.skip_saved, bg=BG_PANEL, fg=TEXT_MID, selectcolor=BG_SURFACE,
            activebackground=BG_PANEL, activeforeground=TEXT_HIGH, font=(FONT_FAMILY, 9), anchor="w")
        self.saved_toggle.pack(fill="x")
        tk.Label(search_options, text="Exacto al saldo primero. Semanas de lunes a domingo.",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8)).pack(anchor="w")

        ec_actions = tk.Frame(entry_card, bg=BG_PANEL)
        ec_actions.pack(fill="x", pady=(4, 0))
        self.search_button = ttk.Button(ec_actions, text="Buscar combinación", command=self.search,
                                        style="Primary.TButton", state="disabled")
        self.search_button.pack(side="left")
        tk.Label(ec_actions, text="Presioná Enter para calcular", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8)).pack(side="left", padx=10)

        # PANELES DE MÉTRICAS (KPIs)
        summary = tk.Frame(main_container, bg=BG_CANVAS)
        self.summary_frame = summary
        summary.pack(fill="x", pady=(0, 12))

        kpis = (
            ("IMPORTE SOLICITADO", self.target_total, TEXT_HIGH),
            ("TOTAL EN E-CHEQUES", self.check_total, EMERALD),
            ("TRANSFERENCIA A COMPLETAR", self.transfer_total, AMBER)
        )

        self.kpi_labels = []
        for col_idx, (title, variable, color) in enumerate(kpis):
            box = tk.Frame(summary, bg=BG_PANEL, padx=16, pady=12, highlightthickness=1, highlightbackground=BORDER)
            box.grid(row=0, column=col_idx, sticky="nsew", padx=(0 if col_idx == 0 else 6, 0 if col_idx == 2 else 6))
            tk.Label(box, text=title, bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8, "bold")).pack(anchor="w")
            val_lbl = tk.Label(box, textvariable=variable, bg=BG_PANEL, fg=color, font=(FONT_FAMILY, 18, "bold"))
            val_lbl.pack(anchor="w", pady=(4, 0))
            self.kpi_labels.append((val_lbl, color))
            summary.columnconfigure(col_idx, weight=1, uniform="kpi")

        # TABLA DE CHEQUES
        table_card = tk.Frame(main_container, bg=BG_PANEL, padx=18, pady=14, highlightthickness=1, highlightbackground=BORDER)
        table_card.pack(fill="both", expand=True, pady=(0, 12))

        tc_header = tk.Frame(table_card, bg=BG_PANEL)
        tc_header.pack(fill="x", pady=(0, 10))

        tc_titles = tk.Frame(tc_header, bg=BG_PANEL)
        tc_titles.pack(side="left", fill="x", expand=True)
        tk.Label(tc_titles, textvariable=self.result_title, bg=BG_PANEL, fg=TEXT_HIGH,
                 font=(FONT_FAMILY, 11, "bold")).pack(anchor="w")
        self.result_detail_label = tk.Label(tc_titles, textvariable=self.result_detail, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8), wraplength=620, justify="left")
        self.result_detail_label.pack(anchor="w", pady=(2, 0))

        tc_actions = tk.Frame(tc_header, bg=BG_PANEL)
        tc_actions.pack(side="right")
        self.save_button = ttk.Button(tc_actions, text="Guardar selección", command=self.save_selection, state="disabled")
        self.save_button.pack(side="left", padx=(0, 8))
        self.registry_button = ttk.Button(tc_actions, text="Registro…", command=self.show_registry)
        self.registry_button.pack(side="left", padx=(0, 8))
        self.copy_button = ttk.Button(tc_actions, text="Copiar mensaje", command=self.copy_message, state="disabled")
        self.copy_button.pack(side="left", padx=(0, 8))
        self.export_button = ttk.Button(tc_actions, text="Exportar CSV…", command=self.export, state="disabled")
        self.export_button.pack(side="left")

        table_frame = tk.Frame(table_card, bg=BG_PANEL)
        table_frame.pack(fill="both", expand=True)

        columns = ("row", "reference", "date_f", "date_g", "id", "name", "amount", "receipt", "batch", "balance")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", height=5)
        headings = ("Fila", "Referencia", "Fecha F", "Fecha G", "ID Cliente", "Cliente", "Importe", "Recibo", "Tanda semanal", "Saldo restante")
        widths = (55, 110, 95, 95, 90, 160, 120, 110, 195, 125)
        for col, title, width in zip(columns, headings, widths):
            self.table.heading(col, text=title)
            self.table.column(col, width=width, minwidth=width, anchor="e" if col in ("amount", "balance") else "w", stretch=col == "name")
        self.table.configure(displaycolumns=("reference", "amount", "batch", "balance", "date_f", "date_g", "id", "name", "receipt", "row"))
        self.table.tag_configure("alternate", background="#202530")
        self.table.tag_configure("glow", background="#1E394A", foreground=CYAN)

        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)

        self.table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # FOOTER CON RESUMEN DEL MENSAJE
        footer_card = tk.Frame(main_container, bg=BG_PANEL, padx=16, pady=10, highlightthickness=1, highlightbackground=BORDER)
        footer_card.pack(fill="x", side="bottom", before=table_card)

        tk.Label(footer_card, text="VISTA PREVIA DE COMUNICACIÓN:", bg=BG_PANEL, fg=TEXT_MUTED, font=(FONT_FAMILY, 8, "bold")).pack(anchor="w")
        self.message_label = tk.Label(footer_card, textvariable=self.message, bg=BG_PANEL, fg=TEXT_MID,
                                      font=(FONT_FAMILY, 9), anchor="w", justify="left")
        self.message_label.pack(fill="x", pady=(2, 0))

        # BARRA INFERIOR DE ESTADO CON SPINNER VECTORIAL
        status_bar = tk.Frame(root, bg=BG_PANEL, padx=20, pady=7, highlightthickness=1, highlightbackground=BORDER)
        status_bar.pack(fill="x", side="bottom", before=main_host)

        # ANIMACIÓN 2: Radar Spinner Canvas en lugar de la barra gris estática
        self.spinner = tk.Canvas(status_bar, width=22, height=22, bg=BG_PANEL, highlightthickness=0)
        self.spinner.pack(side="right", padx=(10, 0))

        tk.Label(status_bar, textvariable=self.status, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(FONT_FAMILY, 8, "bold"), anchor="w").pack(side="left", fill="x", expand=True)

        root.bind("<Configure>", self._on_window_resize)

    # --- SISTEMA DE ANIMACIONES ---
    def _start_animations(self):
        """Inicia las animaciones de pulso en botón y spinner de radar rotativo."""
        self._anim_pulse_step = 0
        self._spinner_angle = 0
        self._loop_button_pulse()
        self._loop_spinner()

    def _stop_animations(self):
        """Detiene todas las animaciones activas y restaura el estado limpio."""
        if self._anim_pulse_timer:
            self.root.after_cancel(self._anim_pulse_timer)
            self._anim_pulse_timer = None
        if self._spinner_timer:
            self.root.after_cancel(self._spinner_timer)
            self._spinner_timer = None

        style = ttk.Style(self.root)
        style.configure("Primary.TButton", background=CYAN, foreground="#08141F")
        self.search_button.configure(text="Buscar combinación")
        self.spinner.delete("all")

    def _loop_button_pulse(self):
        """ANIMACIÓN 1: Glow / Pulse oscilante de color en el botón en ejecución."""
        if not self.busy:
            return
        # Ciclo senoidal suave entre tonos Cyan y Deep Cyan
        pulse_colors = ["#38BDF8", "#22A6E0", "#0EA5E9", "#0284C7", "#0369A1", "#0284C7", "#0EA5E9", "#22A6E0"]
        color = pulse_colors[self._anim_pulse_step % len(pulse_colors)]
        style = ttk.Style(self.root)
        style.configure("Primary.TButton", background=color, foreground="#FFFFFF")
        
        dots = "." * ((self._anim_pulse_step // 2) % 4)
        if self.cancel_button.cget("state") == "normal":
            self.search_button.configure(text=f"Calculando{dots:<3}")

        self._anim_pulse_step += 1
        self._anim_pulse_timer = self.root.after(110, self._loop_button_pulse)

    def _loop_spinner(self):
        """ANIMACIÓN 2: Spinner de radar vectorizado en canvas con estela dinámica."""
        if not self.busy:
            self.spinner.delete("all")
            return
        self.spinner.delete("all")
        cx, cy, r = 11, 11, 8
        self._spinner_angle = (self._spinner_angle + 24) % 360
        start_rad = math.radians(self._spinner_angle)

        # Arco base tenue
        self.spinner.create_oval(cx - r, cy - r, cx + r, cy + r, outline=BORDER, width=2)

        # Arco brillante principal en rotación
        extent = 100
        self.spinner.create_arc(cx - r, cy - r, cx + r, cy + r,
                                start=self._spinner_angle, extent=extent,
                                style="arc", outline=CYAN, width=2.5)

        self._spinner_timer = self.root.after(35, self._loop_spinner)

    def _animate_kpi_reveal(self, step=0):
        """ANIMACIÓN 3: Fade-in progresivo con transparencia cromática en los totales."""
        steps = ["#2A3241", "#4B5A75", "#7288AC", "#9FB5D6", "#C7D8EE"]
        if step < len(steps):
            for lbl, target_color in self.kpi_labels:
                lbl.configure(fg=steps[step])
            self._result_after(40, self._animate_kpi_reveal, step + 1)
        else:
            for lbl, target_color in self.kpi_labels:
                lbl.configure(fg=target_color)

    def _animate_table_rows(self, items, index=0):
        """ANIMACIÓN 4: Cascada de iluminación secuencial sobre los e-cheques encontrados."""
        if index < len(items) and self.table.exists(items[index]):
            item_id = items[index]
            self.table.item(item_id, tags=("glow",))
            # Vuelve a su estado alternado luego de un pulso de luz
            def restore(tid, idx):
                if self.table.exists(tid):
                    self.table.item(tid, tags=("alternate",) if idx % 2 else ())
            self._result_after(160, restore, item_id, index)
            self._result_after(45, self._animate_table_rows, items, index + 1)

    def _result_after(self, delay, callback, *args):
        def run():
            self._result_timers.discard(timer)
            callback(*args)
        timer = self.root.after(delay, run)
        self._result_timers.add(timer)

    def _stop_result_animations(self):
        for timer in self._result_timers:
            self.root.after_cancel(timer)
        self._result_timers.clear()
        for label, color in self.kpi_labels:
            label.configure(fg=color)

    def _on_window_resize(self, event):
        if event.widget == self.root:
            self.message_label.configure(wraplength=max(400, self.root.winfo_width() - 80))
            self.result_detail_label.configure(wraplength=max(300, self.root.winfo_width() - 430))

    def _invalidate(self, *args):
        if self.busy:
            return
        self.result = None
        self._stop_result_animations()
        self.search_context = {}
        self.table.delete(*self.table.get_children())
        for value in (self.target_total, self.check_total, self.transfer_total):
            value.set("$ —")
        self.copy_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.result_title.set("Resultados de la búsqueda")
        self.result_detail.set("Cada fila del Excel se puede usar una sola vez en cada cálculo.")
        self.message.set("Al calcular, vas a poder copiar el resumen para el cliente.")
        self.status.set("Ingresá el importe y presioná Buscar combinación.")
        self.previous_toggle.configure(state="normal" if self.strategy.get().startswith("Progresiva") else "disabled")

    def _set_busy(self, busy, searching=False):
        self.busy = busy
        self.load_button.configure(state="disabled" if busy else "normal")
        self.amount_entry.configure(state="disabled" if busy else "normal")
        ready = self.imported is not None and bool(self.imported.checks)
        self.sheet_box.configure(state="disabled" if busy or not self.path else "readonly")
        self.client_box.configure(state="disabled" if busy or not ready else "readonly")
        self.search_button.configure(state="disabled" if busy or not ready else "normal")
        self.registry_button.configure(state="disabled" if busy else "normal")
        self.cancel_button.configure(state="normal" if busy and searching else "disabled")
        self.notes_button.configure(state="normal" if not busy and self.imported else "disabled")
        for widget in self.date_entries + self.calendar_buttons + [self.clear_dates_button]:
            widget.configure(state="disabled" if busy else "normal")
        for widget in (self.date_box, self.strategy_box):
            widget.configure(state="disabled" if busy else "readonly")
        self.saved_toggle.configure(state="disabled" if busy else "normal")
        self.previous_toggle.configure(state="normal" if not busy and self.strategy.get().startswith("Progresiva") else "disabled")
        if busy:
            self._start_animations()
        else:
            self._stop_animations()

    def _worker(self, kind, function):
        def run():
            try:
                self.events.put((kind, function()))
            except Exception as exc:
                self.events.put(("error", (kind, str(exc))))
        threading.Thread(target=run, daemon=True).start()

    def load_file(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(title="Seleccionar listado de e-cheques", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self._load(path, None)

    def _load(self, path, sheet):
        self._invalidate()
        self._set_busy(True)
        self.status.set("Leyendo el Excel…")
        def read():
            names = list_sheets(path)
            return path, names, read_excel(path, sheet)
        self._worker("loaded", read)

    def _sheet_changed(self, event=None):
        if self.path and not self.busy:
            self._load(self.path, self.sheet.get())

    def _clear_dates(self):
        if not self.busy:
            self.date_from.set("")
            self.date_to.set("")

    def _open_calendar(self, variable):
        if self.busy:
            return
        try:
            selected = parse_date(variable.get())
        except DataError:
            selected = None
        dates = [check_date(c, self.date_field.get()) for c in self.imported.checks] if self.imported else []
        dates = [d for d in dates if d is not None]
        initial = selected or ((min(dates) if variable is self.date_from else max(dates)) if dates else date.today())
        current = [initial.year, initial.month]
        window = tk.Toplevel(self.root)
        window.title("Elegir fecha · columna " + self.date_field.get())
        window.configure(bg=BG_PANEL, padx=12, pady=12)
        window.resizable(False, False)
        window.transient(self.root)
        title = tk.StringVar()
        navigation = tk.Frame(window, bg=BG_PANEL)
        navigation.pack(fill="x")
        days = tk.Frame(window, bg=BG_PANEL)
        days.pack(pady=8)
        months = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")

        def choose(value):
            variable.set(value.strftime("%d/%m/%Y") if value else "")
            window.destroy()

        def draw(delta=0):
            month_index = current[0] * 12 + current[1] - 1 + delta
            year, month = divmod(month_index, 12)
            if not 1 <= year <= 9999:
                return
            current[:] = year, month + 1
            title.set("{} {}".format(months[month], year))
            for child in days.winfo_children():
                child.destroy()
            for col, day in enumerate(("Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do")):
                tk.Label(days, text=day, bg=BG_PANEL, fg=TEXT_MUTED).grid(row=0, column=col)
            for row, week in enumerate(calendar.monthcalendar(year, month + 1), 1):
                for col, day in enumerate(week):
                    if day:
                        value = date(year, month + 1, day)
                        ttk.Button(days, text=str(day), width=3,
                                   style="Primary.TButton" if value == selected else "TButton",
                                   command=lambda value=value: choose(value)).grid(row=row, column=col, padx=1, pady=1)
        ttk.Button(navigation, text="‹", width=3, command=lambda: draw(-1)).pack(side="left")
        tk.Label(navigation, textvariable=title, bg=BG_PANEL, fg=TEXT_HIGH, width=20).pack(side="left")
        ttk.Button(navigation, text="›", width=3, command=lambda: draw(1)).pack(side="right")
        ttk.Button(window, text="Sin límite de fecha", command=lambda: choose(None)).pack(fill="x")
        window.bind("<Escape>", lambda event: window.destroy())
        draw()
        window.grab_set()

    def _loaded(self, payload):
        self.path, sheets, self.imported = payload
        self.sheet_box.configure(values=sheets)
        self.sheet.set(self.imported.sheet)
        self.file_label.set(Path(self.path).name)
        checks = self.imported.checks
        client_names = {check.client_id: check.client_name for check in checks}
        self.clients = {"{} · {}".format(identifier, name): identifier for identifier, name in sorted(client_names.items())}
        self.client_box.configure(values=["Todos los clientes"] + list(self.clients))
        self.client.set("Todos los clientes")
        self.date_from.set("")
        self.date_to.set("")
        self.file_detail.set("{} e-cheques · {} clientes · Total disponible: {}".format(
            len(checks), len(client_names), money(sum(check.amount for check in checks))))
        if self.imported.rejected:
            self.notice.set("{} filas excluidas por datos incompletos o importes inválidos. Revisá el detalle antes de buscar. {} observaciones conservadas.".format(
                len(self.imported.rejected), len(self.imported.notes)))
        elif self.imported.notes:
            self.notice.set("{} observaciones de fechas o datos de origen. Los registros están incluidos; ver detalle.".format(len(self.imported.notes)))
        else:
            self.notice.set("Todos los registros se cargaron. Las fechas F y G se muestran como figuran en el archivo.")
        self._set_busy(False)
        self.status.set("Excel cargado correctamente. Ingresá el importe que querés cubrir.")
        self.amount_entry.focus_set()
        if not checks:
            messagebox.showwarning("No hay e-cheques para buscar", "No se encontraron registros utilizables. Se espera ID en J, importe en V y fechas en F/G. Revisá Ver detalle.")
        elif self.imported.rejected:
            self.show_import_notes()

    def search(self):
        if self.busy or not self.imported or not self.imported.checks:
            return
        try:
            target = parse_amount(self.target.get())
            start, end = parse_date(self.date_from.get()), parse_date(self.date_to.get())
            date_field = self.date_field.get()
            identifier = self.clients.get(self.client.get())
            records = [check for check in self.imported.checks if identifier is None or check.client_id == identifier]
            checks, excluded_dates = filter_checks(records, date_field, start, end)
            excluded_saved = 0
            if self.skip_saved.get():
                checks, excluded_saved = exclude_saved(checks, Path(self.path).name, self.imported.sheet)
        except DataError as exc:
            messagebox.showwarning("Revisá el importe y las fechas", str(exc))
            self.amount_entry.focus_set()
            return
        self._invalidate()
        if not checks:
            messagebox.showwarning("Sin e-cheques para esas fechas", "No hay registros disponibles que cumplan el cliente y el rango elegidos. {} tienen la fecha vacía o inválida y {} ya están guardados en el registro.".format(excluded_dates, excluded_saved))
            return
        progressive = self.strategy.get().startswith("Progresiva")
        fewest = self.strategy.get().endswith("menos e-cheques")
        prefer_previous = self.prefer_previous.get()
        self.search_context = {"file": Path(self.path).name, "sheet": self.imported.sheet,
                               "client": self.client.get(), "date_field": date_field, "from": self.date_from.get(), "to": self.date_to.get(),
                               "excluded_dates": excluded_dates, "excluded_saved": excluded_saved, "previous": progressive and prefer_previous}
        self.target_total.set(money(target))
        self.cancel.clear()
        self._set_busy(True, searching=True)
        self.status.set("Buscando la combinación más cercana sin superar {}…".format(money(target)))
        def calculate():
            progress = lambda message: self.events.put(("progress", message))
            if progressive:
                return find_progressive(checks, target, self.cancel, progress, date_field, prefer_previous)
            if fewest:
                result = find_fewest(checks, target, self.cancel, progress)
                result.date_field = date_field
                return result
            result = find_combination(checks, target, self.cancel, progress)
            result.date_field = date_field
            return result
        self._worker("found", calculate)

    def stop(self):
        self.cancel.set()
        self.cancel_button.configure(state="disabled")
        self.status.set("Terminando y recuperando una combinación válida…")

    def _found(self, result):
        self._stop_result_animations()
        self.table.delete(*self.table.get_children())
        self.result = result
        self._set_busy(False)
        self.target_total.set(money(result.target))
        self.check_total.set(money(result.total))
        self.transfer_total.set(money(result.transfer))
        if self.search_context.get("saved"):
            title = "Selección guardada · del registro"
            detail = "{} e-cheques guardados el {}. Faltan {} por transferencia.".format(
                len(result.checks), self.search_context["saved"], money(result.transfer))
        elif not result.completed:
            title = "Búsqueda detenida · mejor combinación encontrada"
            detail = "El total no supera lo solicitado. Podría existir una combinación más cercana; volvé a buscar para completar el cálculo."
        elif result.transfer == 0:
            title = "Coincidencia exacta"
            detail = "{} e-cheques cubren todo el importe. No hace falta completar por transferencia.".format(len(result.checks))
        elif not result.checks:
            title = "Ningún e-cheque entra en el importe solicitado"
            detail = "Los e-cheques disponibles superan el objetivo. El importe completo queda a completar por transferencia."
        elif result.strategy == "fewest":
            title = "Mejor suma · con los e-cheques más altos"
            detail = "{} e-cheques, de mayor a menor. {} Faltan {} por transferencia.".format(
                len(result.checks),
                "Es la menor cantidad posible para esa suma." if result.count_minimal
                else "Menor cantidad encontrada; podría existir una con menos.", money(result.transfer))
        elif result.strategy == "progressive":
            title = "Selección progresiva · de mayor a menor"
            detail = "{} e-cheques elegidos por monto y tandas. Faltan {}. Otra combinación podría cubrir más.".format(len(result.checks), money(result.transfer))
        else:
            title = "La combinación más cercana sin pasarse"
            detail = "{} e-cheques seleccionados. La diferencia a completar por transferencia es {}.".format(len(result.checks), money(result.transfer))
        self.result_title.set(title)
        self.result_detail.set(detail)

        inserted_ids = []
        balance = result.target
        self.table.heading("batch", text="Tanda semanal · fecha " + result.date_field)
        for i, check in enumerate(result.checks):
            balance -= check.amount
            row_id = self.table.insert("", "end", values=(check.row, check.reference, check.date_f, check.date_g,
                                                         check.client_id, check.client_name, money(check.amount), check.receipt,
                                                         batch_label(check, result.date_field), money(balance)),
                                       tags=("alternate",) if i % 2 else ())
            inserted_ids.append(row_id)

        # Dispara la cascada y la transición de métricas
        self._animate_kpi_reveal()
        self._animate_table_rows(inserted_ids)

        self.message.set(customer_message(result))
        self.copy_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self.save_button.configure(state="disabled" if self.search_context.get("saved") else "normal")
        if self.search_context.get("saved"):
            self.status.set("Selección del registro cargada. Podés copiar el mensaje o exportarla.")
            return
        self.status.set("{} · {:.2f} s · Fecha {} · {} registros excluidos por fecha inválida · {} ya guardados en el registro.".format(
            "Cálculo completo" if result.completed else "Cálculo detenido", result.elapsed, result.date_field,
            self.search_context.get("excluded_dates", 0), self.search_context.get("excluded_saved", 0)))

    def _poll(self):
        self.root.after_cancel(self.poll_timer)
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "loaded":
                    self._loaded(payload)
                elif kind == "found":
                    self._found(payload)
                elif kind == "progress":
                    if not self.cancel.is_set():
                        self.status.set(payload)
                elif kind == "error":
                    source, text = payload
                    if source == "loaded":
                        self.path, self.imported = None, None
                        self.sheet.set("")
                        self.file_label.set("No se pudo cargar el archivo")
                        self.file_detail.set("Elegí un archivo .xlsx con el formato esperado.")
                        self.notice.set("")
                    self._set_busy(False)
                    self.status.set("No se pudo completar la operación.")
                    messagebox.showerror("No se pudo completar", text)
        except queue.Empty:
            pass
        self.poll_timer = self.root.after(100, self._poll)

    def show_import_notes(self):
        if not self.imported:
            return
        window = tk.Toplevel(self.root)
        window.title("Detalle de importación")
        window.geometry("800x480")
        window.configure(bg=BG_PANEL)
        try:
            window.attributes("-alpha", 0.98)
        except Exception:
            pass
        text = tk.Text(window, wrap="word", padx=16, pady=16, font=("Segoe UI", 10),
                       bg=BG_PANEL, fg=TEXT_HIGH, insertbackground=TEXT_HIGH, relief="flat")
        scrollbar = ttk.Scrollbar(window, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        lines = ["Archivo: {}".format(self.path), "Hoja: {}".format(self.imported.sheet), "",
                 "FILAS EXCLUIDAS ({})".format(len(self.imported.rejected))]
        lines.extend(self.imported.rejected or ["Ninguna."])
        lines.extend(["", "OBSERVACIONES: REGISTROS INCLUIDOS ({})".format(len(self.imported.notes))])
        lines.extend(self.imported.notes or ["Ninguna."])
        text.insert("1.0", "\n".join(lines))
        text.configure(state="disabled")

    def copy_message(self):
        if self.result is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(customer_message(self.result))
        self.status.set("Mensaje copiado al portapapeles. Listo para enviar al cliente.")

    def export(self):
        if self.result is None:
            return
        path = filedialog.asksaveasfilename(title="Guardar selección", defaultextension=".csv",
                                            initialfile="seleccion_echeques.csv", filetypes=[("CSV para Excel", "*.csv")])
        if not path:
            return
        def safe(value):
            text = str(value)
            if text.startswith(("=", "+", "-", "@", "\t", "\r")):
                return "'" + text
            return text
        def decimal(cents):
            return "{},{:02d}".format(cents // 100, cents % 100)
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(["Archivo", safe(self.search_context.get("file", ""))])
                writer.writerow(["Hoja", safe(self.search_context.get("sheet", ""))])
                state = ("Selección guardada en el registro" if self.search_context.get("saved") else
                         "Búsqueda detenida; óptimo no confirmado" if not self.result.completed else
                         "Selección progresiva" if self.result.strategy == "progressive" else
                         "Mejor suma con menos e-cheques" if self.result.strategy == "fewest" else "Óptimo confirmado")
                writer.writerow(["Estado", state])
                writer.writerow(["Fecha usada", self.result.date_field])
                writer.writerow(["Desde", safe(self.search_context.get("from", ""))])
                writer.writerow(["Hasta", safe(self.search_context.get("to", ""))])
                writer.writerow(["Priorizar semanas anteriores", "Sí" if self.search_context.get("previous") else "No"])
                writer.writerow(["Importe solicitado", decimal(self.result.target)])
                writer.writerow(["Total e-cheques", decimal(self.result.total)])
                writer.writerow(["A completar por transferencia", decimal(self.result.transfer)])
                writer.writerow([])
                writer.writerow(["Fila Excel", "Referencia C", "Fecha F", "Fecha G", "ID cliente", "Cliente", "Importe", "Recibo", "Orden", "Tanda semanal", "Saldo restante"])
                balance = self.result.target
                for order, check in enumerate(self.result.checks, 1):
                    balance -= check.amount
                    writer.writerow([check.row, "'" + check.reference, safe(check.date_f), safe(check.date_g),
                                     "'" + check.client_id, safe(check.client_name), decimal(check.amount), safe(check.receipt),
                                     order, batch_label(check, self.result.date_field), decimal(balance)])
            self.status.set("Selección exportada en {}.".format(path))
        except OSError as exc:
            messagebox.showerror("No se pudo guardar", str(exc))

    def save_selection(self):
        if self.result is None or not self.result.checks or self.search_context.get("saved"):
            return
        context = self.search_context
        try:
            repeated = saved_overlap(self.result, context.get("file", ""), context.get("sheet", ""))
        except DataError as exc:
            messagebox.showerror("No se pudo leer el registro", str(exc))
            return
        if repeated and not messagebox.askyesno(
                "E-cheques ya guardados",
                "{} de estos e-cheques ya figuran en otra selección guardada ({}). ¿Guardar igualmente?".format(
                    len(repeated), ", ".join(repeated[:5]) + ("…" if len(repeated) > 5 else ""))):
            return
        note = simpledialog.askstring("Guardar selección", "Nota opcional (por ejemplo, nombre del pago):", parent=self.root)
        if note is None:
            return
        try:
            save_selection(self.result, context, note)
        except DataError as exc:
            messagebox.showerror("No se pudo guardar", str(exc))
            return
        self.status.set("Selección guardada en el registro: {}".format(registry_path()))

    def show_registry(self):
        if self.busy:
            return
        window = tk.Toplevel(self.root)
        window.title("Registro de selecciones guardadas")
        window.geometry("920x420")
        window.configure(bg=BG_PANEL, padx=12, pady=12)
        window.transient(self.root)
        columns = ("saved", "note", "file", "target", "total", "transfer", "count")
        tree = ttk.Treeview(window, columns=columns, show="headings", selectmode="browse")
        for col, title, width in zip(columns, ("Guardado", "Nota", "Archivo", "Solicitado", "E-cheques", "Faltante", "Cant."),
                                     (140, 180, 170, 110, 110, 100, 50)):
            tree.heading(col, text=title)
            tree.column(col, width=width, anchor="e" if col in ("target", "total", "transfer", "count") else "w",
                        stretch=col == "note")
        buttons = tk.Frame(window, bg=BG_PANEL)
        buttons.pack(side="bottom", fill="x", pady=(10, 0))
        tree.pack(side="top", fill="both", expand=True)
        entries = {}

        def refresh():
            tree.delete(*tree.get_children())
            entries.clear()
            for entry in reversed(load_registry()):
                entries[entry["id"]] = entry
                tree.insert("", "end", iid=entry["id"], values=(
                    entry.get("saved_at", ""), entry.get("note", ""), entry.get("file", ""),
                    money(entry.get("target", 0)), money(entry.get("total", 0)),
                    money(entry.get("transfer", 0)), len(entry.get("checks", []))))

        def selected():
            chosen = tree.selection()
            return entries.get(chosen[0]) if chosen else None

        def load():
            entry = selected()
            if entry is None:
                return
            try:
                result = entry_to_result(entry)
            except DataError as exc:
                messagebox.showerror("No se pudo cargar", str(exc), parent=window)
                return
            self._invalidate()
            self.search_context = {"file": entry.get("file", ""), "sheet": entry.get("sheet", ""),
                                   "from": entry.get("from", ""), "to": entry.get("to", ""),
                                   "previous": entry.get("previous", False), "excluded_dates": 0,
                                   "saved": entry.get("saved_at", "")}
            window.destroy()
            self._found(result)

        def remove():
            entry = selected()
            if entry is None or not messagebox.askyesno(
                    "Eliminar selección", "¿Eliminar esta selección del registro?", parent=window):
                return
            try:
                delete_selection(entry["id"])
            except DataError as exc:
                messagebox.showerror("No se pudo eliminar", str(exc), parent=window)
            refresh()

        ttk.Button(buttons, text="Cargar selección", style="Primary.TButton", command=load).pack(side="left")
        ttk.Button(buttons, text="Eliminar", style="Danger.TButton", command=remove).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cerrar", command=window.destroy).pack(side="right")
        tk.Label(buttons, text="Archivo: " + str(registry_path()), bg=BG_PANEL, fg=TEXT_MUTED,
                 font=(self.FONT_FAMILY, 8)).pack(side="right", padx=10)
        tree.bind("<Double-1>", lambda event: load())
        window.bind("<Escape>", lambda event: window.destroy())
        refresh()

    def close(self):
        self.cancel.set()
        self._stop_result_animations()
        self._stop_animations()
        self.root.after_cancel(self.poll_timer)
        self.root.update_idletasks()
        self.root.destroy()


def _self_test(destination):
    """Verificación automática del ejecutable distribuido, sin mostrar ventanas."""
    import json
    import os
    import tempfile
    import time
    os.environ["ECHECK_REGISTRO"] = os.path.join(tempfile.mkdtemp(), "registro.json")
    from core import Check, ImportResult
    root = tk.Tk()
    root.withdraw()
    callback_errors = []
    root.report_callback_exception = lambda kind, error, trace: callback_errors.append(str(error))
    app = Application(root)
    report = {"ok": False}
    try:
        records = [Check(i + 2, "DEMO-{}".format(i), "01/09/2025", "50/09/2025", "00123",
                         "Cliente de prueba", amount, "Recibo de prueba")
                   for i, amount in enumerate((30000000, 15000000, 60000000))]
        app._loaded(("prueba.xlsx", ["Pagos"], ImportResult(records, [], [], "Pagos")))
        app.target.set("500.000")
        app.search()
        deadline = time.monotonic() + 15
        while app.busy and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        result = app.result
        report = {"ok": not callback_errors and result is not None and result.optimal and result.total == 45000000
                        and result.transfer == 5000000 and len(app.table.get_children()) == 2,
                  "total_cents": result.total if result else None,
                  "transfer_cents": result.transfer if result else None,
                  "tk": root.tk.call("info", "patchlevel"), "frozen": bool(getattr(sys, "frozen", False)),
                  "callback_errors": callback_errors}
        dated = [Check(i+2,"SEM-{}".format(i),value,value,"00123","Demo",amount,"")
                 for i,(value,amount) in enumerate((("21/09/2026",90000000),
                     ("14/09/2026",6000000),("07/09/2026",4000000)))]
        app._loaded(("prueba.xlsx",["Pagos"],ImportResult(dated,[],[],"Pagos")))
        app.target.set("1.000.000")
        app.search()
        deadline = time.monotonic() + 15
        while app.busy and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        report["weekly_selection"] = bool(app.result and [c.amount for c in app.result.checks] == [90000000,6000000,4000000])
        app.date_from.set("14/09/2026")
        app.date_to.set("14/09/2026")
        app.search()
        deadline = time.monotonic() + 15
        while app.busy and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        report["date_filter"] = bool(app.result and app.result.total == 6000000 and len(app.result.checks) == 1)
        report["ok"] = report["ok"] and report["weekly_selection"] and report["date_filter"] and not callback_errors
    except Exception as exc:
        report["error"] = str(exc)
    finally:
        root.update_idletasks()
        app.close()
        Path(destination).write_text(json.dumps(report, indent=2), encoding="utf-8")
    return 0 if report["ok"] else 1


def main():
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        return _self_test(sys.argv[2])
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except (AttributeError, OSError):
                pass
    root = tk.Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    sys.exit(main())
