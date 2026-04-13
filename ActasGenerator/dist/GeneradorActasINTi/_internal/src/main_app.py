"""
main_app.py - Interfaz gráfica principal del Generador de Actas INTi

Este módulo contiene la clase ModernActaGeneratorApp que construye y gestiona
toda la interfaz gráfica usando tkinter (con soporte opcional de ttkbootstrap
para un estilo más moderno).

Estructura de la interfaz:
  - Barra de menú (Archivo / Ver / Ayuda)
  - Header con nombre y versión de la app
  - Notebook con 3 pestañas:
      1. Dashboard   : estadísticas globales y actas recientes
      2. Generar Actas: carga de Excel, selección de sesiones y generación
      3. Historial   : registro de todas las actas generadas con búsqueda
  - Barra de estado inferior

Bugs corregidos respecto a la versión anterior:
  - _abrir_archivo()        : ahora abre el .docx con os.startfile()
  - _ver_detalles()         : muestra información real del acta seleccionada
  - _actualizar_estadisticas(): actualiza la card 'total_sesiones' (antes siempre en 0)
  - _generar_actas_thread() : llama a guardar_estadistica_sesion() para poblar
                              el panel de estadísticas por estado del dashboard
  - Búsqueda en historial   : filtra en tiempo real mientras el usuario escribe
"""

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import threading
from pathlib import Path
from datetime import datetime
from typing import Optional

# pandas se usa para exportar el historial a Excel
# Es opcional: si no está instalado la exportación muestra un aviso
try:
    import pandas as pd
    PANDAS_DISPONIBLE = True
except ImportError:
    PANDAS_DISPONIBLE = False

# ttkbootstrap es opcional: si está instalado mejora el aspecto visual
try:
    import ttkbootstrap as ttkb
    from ttkbootstrap.constants import *
    USE_TTKBOOTSTRAP = True
except ImportError:
    USE_TTKBOOTSTRAP = False

from config import APP_NAME, APP_VERSION, APP_WIDTH, APP_HEIGHT, OUTPUT_DIR
from excel_processor import ExcelProcessor, ResumenSesion
from document_generator import ActaDocumentGenerator, generar_nombre_archivo
from database import db_manager, ActaRecord
from config_editor import abrir_configuracion


# ===========================================================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ===========================================================================

class ModernActaGeneratorApp:
    """
    Aplicación principal del Generador de Actas INTi.

    Gestiona toda la interfaz gráfica y coordina las acciones del usuario
    con los módulos de procesamiento (ExcelProcessor, ActaDocumentGenerator,
    DatabaseManager).
    """

    def __init__(self):
        """
        Inicializa la ventana principal, variables de estado y construye
        todos los componentes de la interfaz.
        """
        # Crear la ventana raíz con ttkbootstrap si está disponible
        if USE_TTKBOOTSTRAP:
            self.root = ttkb.Window(themename="flatly")
        else:
            self.root = tk.Tk()

        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.root.minsize(1200, 700)

        # Centrar la ventana en la pantalla al abrir
        self._center_window()

        # ---------------------------------------------------------------
        # Variables de estado de la aplicación
        # ---------------------------------------------------------------

        # Ruta del archivo Excel actualmente cargado (vinculada al Entry)
        self.excel_path = tk.StringVar()

        # Instancia del procesador de Excel (None hasta que se cargue un archivo)
        self.excel_processor: Optional[ExcelProcessor] = None

        # Lista de IDs de sesiones que el usuario ha marcado para generar
        self.sesiones_seleccionadas = []

        # ---------------------------------------------------------------
        # Construir la interfaz
        # ---------------------------------------------------------------
        self._create_styles()
        self._create_menu()
        self._create_main_interface()

        # Cargar estadísticas del historial existente al arrancar
        self._actualizar_estadisticas()

    # -----------------------------------------------------------------------
    # CONFIGURACIÓN INICIAL DE LA VENTANA
    # -----------------------------------------------------------------------

    def _center_window(self):
        """
        Centra la ventana en la pantalla del usuario.
        Se llama después de crear la ventana para calcular las coordenadas
        correctas basándose en el tamaño real de la ventana.
        """
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth()  // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def _create_styles(self):
        """
        Define estilos personalizados de ttk cuando ttkbootstrap no está
        disponible. Con ttkbootstrap los estilos los maneja el tema elegido.
        """
        if not USE_TTKBOOTSTRAP:
            style = ttk.Style()
            style.configure("Title.TLabel",    font=("Helvetica", 16, "bold"), foreground="#1a5f7a")
            style.configure("Subtitle.TLabel", font=("Helvetica", 12), foreground="#2c3e50")
            style.configure("Card.TFrame",     background="white", relief="raised", borderwidth=2)
            style.configure("Header.TLabel",   font=("Helvetica", 14, "bold"), foreground="#159895")
            style.configure("Info.TLabel",     font=("Helvetica", 10), foreground="#57c5b6")

    # -----------------------------------------------------------------------
    # BARRA DE MENÚ
    # -----------------------------------------------------------------------

    def _create_menu(self):
        """
        Construye la barra de menú superior con tres menús:
          - Archivo : cargar Excel y salir
          - Ver     : navegación directa a cada pestaña
          - Ayuda   : información sobre la aplicación
        """
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Menú Archivo
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=file_menu)
        file_menu.add_command(label="Cargar Excel", command=self._cargar_excel)
        file_menu.add_separator()
        # Abre la ventana de configuración para editar directores, horas, etc.
        file_menu.add_command(label="Configuración", command=lambda: abrir_configuracion(self.root))
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self.root.quit)

        # Menú Ver: acceso directo a cada pestaña del notebook
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver", menu=view_menu)
        view_menu.add_command(label="Dashboard",     command=lambda: self.notebook.select(0))
        view_menu.add_command(label="Generar Actas", command=lambda: self.notebook.select(1))
        view_menu.add_command(label="Historial",     command=lambda: self.notebook.select(2))

        # Menú Ayuda
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=help_menu)
        help_menu.add_command(label="Acerca de", command=self._mostrar_acerca)

    # -----------------------------------------------------------------------
    # ESTRUCTURA PRINCIPAL
    # -----------------------------------------------------------------------

    def _create_main_interface(self):
        """
        Construye el layout principal de la ventana:
          1. Frame contenedor con padding
          2. Header con título y versión
          3. Notebook con las tres pestañas
          4. Barra de estado en la parte inferior
        """
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._create_header(main_frame)

        # Notebook (sistema de pestañas)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # Crear los frames de cada pestaña
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_generar   = ttk.Frame(self.notebook)
        self.tab_historial = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text=" Dashboard ")
        self.notebook.add(self.tab_generar,   text=" Generar Actas ")
        self.notebook.add(self.tab_historial, text=" Historial ")

        # Poblar cada pestaña con su contenido
        self._create_dashboard_tab()
        self._create_generar_tab()
        self._create_historial_tab()

        # Barra de estado: muestra mensajes de progreso y resultado
        self.status_bar = ttk.Label(
            main_frame, text="Listo", relief=tk.SUNKEN, anchor=tk.W
        )
        self.status_bar.pack(fill=tk.X, pady=(5, 0))

    def _create_header(self, parent):
        """
        Crea el encabezado superior con el logo, nombre de la aplicación y su versión.
        """
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        # Cargar y mostrar el logo
        try:
            from config import ASSETS_DIR
            logo_path = ASSETS_DIR / "logo.jpeg"
            if logo_path.exists():
                # Cargar imagen y redimensionar
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((40, 40), Image.Resampling.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = ttk.Label(header_frame, image=logo_photo)
                logo_label.image = logo_photo  # Mantener referencia
                logo_label.pack(side=tk.LEFT, padx=(0, 10))
        except Exception as e:
            # Si falla cargar el logo, continuar sin él
            pass

        # Título de la aplicación
        ttk.Label(
            header_frame,
            text=f"{APP_NAME}",
            font=("Helvetica", 18, "bold"),
            foreground="#1a5f7a"
        ).pack(side=tk.LEFT)

        # Versión
        ttk.Label(
            header_frame,
            text=f"v{APP_VERSION}",
            font=("Helvetica", 11),
            foreground="#57c5b6"
        ).pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))

    # -----------------------------------------------------------------------
    # PESTAÑA 1: DASHBOARD
    # -----------------------------------------------------------------------

    def _create_dashboard_tab(self):
        """
        Construye la pestaña Dashboard con:
          - 4 tarjetas de estadísticas (Total Actas, Aprobados, Negados, Sesiones)
          - Tabla de estadísticas por estado (alimentada por estadisticas_sesion)
          - Tabla de actas recientes (últimas 10)
          - Botón para refrescar manualmente
        """
        container = ttk.Frame(self.tab_dashboard)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(
            container,
            text="Dashboard - Estadísticas",
            font=("Helvetica", 16, "bold")
        ).pack(anchor=tk.W, pady=(0, 20))

        # --- Fila de tarjetas de estadísticas ---
        stats_frame = ttk.Frame(container)
        stats_frame.pack(fill=tk.X, pady=(0, 20))

        # Configuración de cada tarjeta: (clave_interna, título, icono, color)
        card_configs = [
            ("total_actas",     "Total Actas", "📄", "primary"),
            ("total_aprobados", "Aprobados",   "✅", "success"),
            ("total_negados",   "Negados",     "❌", "danger"),
            ("total_sesiones",  "Sesiones",    "📊", "info"),
        ]

        self.stat_cards = {}
        for i, (key, title, icon, color) in enumerate(card_configs):
            card = self._create_stat_card(stats_frame, title, icon, color)
            card.grid(row=0, column=i, padx=10, pady=5, sticky="nsew")
            self.stat_cards[key] = card
            stats_frame.columnconfigure(i, weight=1)

        # --- Panel de detalles (dos columnas) ---
        details_frame = ttk.Frame(container)
        details_frame.pack(fill=tk.BOTH, expand=True)

        # Panel izquierdo: estadísticas por estado
        left_panel = ttk.LabelFrame(
            details_frame, text=" Estadísticas por Estado ", padding=10
        )
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        cols_estados = ("estado", "aprobados", "negados", "total")
        self.tree_estados = ttk.Treeview(
            left_panel, columns=cols_estados, show="headings", height=15
        )
        self.tree_estados.heading("estado",    text="Estado")
        self.tree_estados.heading("aprobados", text="Aprobados")
        self.tree_estados.heading("negados",   text="Negados")
        self.tree_estados.heading("total",     text="Total")

        self.tree_estados.column("estado",    width=200)
        self.tree_estados.column("aprobados", width=100, anchor=tk.CENTER)
        self.tree_estados.column("negados",   width=100, anchor=tk.CENTER)
        self.tree_estados.column("total",     width=100, anchor=tk.CENTER)

        sb_estados = ttk.Scrollbar(left_panel, orient=tk.VERTICAL,
                                   command=self.tree_estados.yview)
        self.tree_estados.configure(yscrollcommand=sb_estados.set)
        self.tree_estados.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_estados.pack(side=tk.RIGHT, fill=tk.Y)

        # Panel derecho: actas recientes
        right_panel = ttk.LabelFrame(
            details_frame, text=" Actas Recientes ", padding=10
        )
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        cols_recientes = ("sesion", "fecha", "aprobados", "negados")
        self.tree_recientes = ttk.Treeview(
            right_panel, columns=cols_recientes, show="headings", height=15
        )
        self.tree_recientes.heading("sesion",    text="Sesión")
        self.tree_recientes.heading("fecha",     text="Fecha")
        self.tree_recientes.heading("aprobados", text="Aprobados")
        self.tree_recientes.heading("negados",   text="Negados")

        self.tree_recientes.column("sesion",    width=120)
        self.tree_recientes.column("fecha",     width=120)
        self.tree_recientes.column("aprobados", width=80, anchor=tk.CENTER)
        self.tree_recientes.column("negados",   width=80, anchor=tk.CENTER)

        sb_recientes = ttk.Scrollbar(right_panel, orient=tk.VERTICAL,
                                     command=self.tree_recientes.yview)
        self.tree_recientes.configure(yscrollcommand=sb_recientes.set)
        self.tree_recientes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_recientes.pack(side=tk.RIGHT, fill=tk.Y)

        # Botón para refrescar estadísticas manualmente
        ttk.Button(
            container,
            text="🔄 Actualizar Estadísticas",
            command=self._actualizar_estadisticas
        ).pack(anchor=tk.E, pady=(10, 0))

    def _create_stat_card(self, parent, title: str, icon: str, color: str):
        """
        Crea una tarjeta de estadística con icono, valor numérico y título.

        El atributo value_label se guarda en la tarjeta para poder actualizar
        el número desde _actualizar_estadisticas() sin reconstruir el widget.

        Args:
            parent: widget padre donde se colocará la tarjeta
            title : texto descriptivo debajo del número
            icon  : emoji o carácter unicode para el icono
            color : nombre de estilo de ttkbootstrap (ignorado sin ttkbootstrap)

        Retorna el frame de la tarjeta con el atributo .value_label añadido.
        """
        if USE_TTKBOOTSTRAP:
            card = ttkb.Frame(parent, bootstyle=color)
        else:
            card = ttk.Frame(parent, relief="raised", padding=10)

        ttk.Label(card, text=icon, font=("Helvetica", 32)).pack(pady=(10, 5))

        value_label = ttk.Label(card, text="0", font=("Helvetica", 24, "bold"))
        value_label.pack()
        card.value_label = value_label  # Referencia para actualizar el valor

        ttk.Label(card, text=title, font=("Helvetica", 10)).pack(pady=(5, 10))

        return card

    # -----------------------------------------------------------------------
    # PESTAÑA 2: GENERAR ACTAS
    # -----------------------------------------------------------------------

    def _create_generar_tab(self):
        """
        Construye la pestaña de generación de actas con:
          - Campo de ruta del Excel + botones Cargar y Procesar
          - Tabla de sesiones detectadas con columna de selección
          - Botones Seleccionar Todas / Deseleccionar Todas / Generar
          - Barra de progreso (oculta hasta que se inicia la generación)
        """
        container = ttk.Frame(self.tab_generar)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(
            container,
            text="Generar Actas desde Excel",
            font=("Helvetica", 16, "bold")
        ).pack(anchor=tk.W, pady=(0, 20))

        # --- Sección de carga de archivo ---
        file_frame = ttk.LabelFrame(container, text=" Archivo Excel ", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 20))

        ttk.Entry(
            file_frame, textvariable=self.excel_path,
            state="readonly", width=80
        ).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        ttk.Button(
            file_frame, text="📁 Cargar Excel",
            command=self._cargar_excel
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            file_frame, text="🔄 Procesar",
            command=self._procesar_excel
        ).pack(side=tk.LEFT)

        # --- Tabla de sesiones detectadas ---
        self.sesiones_frame = ttk.LabelFrame(
            container, text=" Sesiones Detectadas ", padding=10
        )
        self.sesiones_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        cols_sesiones = ("seleccionar", "sesion", "fecha", "aprobados", "negados", "estados")
        self.tree_sesiones = ttk.Treeview(
            self.sesiones_frame, columns=cols_sesiones,
            show="headings", height=10
        )

        self.tree_sesiones.heading("seleccionar", text="✓")
        self.tree_sesiones.heading("sesion",      text="Sesión")
        self.tree_sesiones.heading("fecha",       text="Fecha")
        self.tree_sesiones.heading("aprobados",   text="Aprobados")
        self.tree_sesiones.heading("negados",     text="Negados")
        self.tree_sesiones.heading("estados",     text="Estados Involucrados")

        self.tree_sesiones.column("seleccionar", width=40,  anchor=tk.CENTER)
        self.tree_sesiones.column("sesion",      width=120)
        self.tree_sesiones.column("fecha",       width=120)
        self.tree_sesiones.column("aprobados",   width=80,  anchor=tk.CENTER)
        self.tree_sesiones.column("negados",     width=80,  anchor=tk.CENTER)
        self.tree_sesiones.column("estados",     width=300)

        sb_sesiones = ttk.Scrollbar(
            self.sesiones_frame, orient=tk.VERTICAL,
            command=self.tree_sesiones.yview
        )
        self.tree_sesiones.configure(yscrollcommand=sb_sesiones.set)
        self.tree_sesiones.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_sesiones.pack(side=tk.RIGHT, fill=tk.Y)

        # Clic en la columna ✓ para seleccionar/deseleccionar una sesión
        self.tree_sesiones.bind("<ButtonRelease-1>", self._on_sesion_click)

        # --- Botones de acción ---
        actions_frame = ttk.Frame(container)
        actions_frame.pack(fill=tk.X)

        ttk.Button(
            actions_frame, text="✅ Seleccionar Todas",
            command=self._seleccionar_todas
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            actions_frame, text="⬜ Deseleccionar Todas",
            command=self._deseleccionar_todas
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            actions_frame, text="📄 Generar Actas Seleccionadas",
            command=self._generar_actas
        ).pack(side=tk.RIGHT)

        # --- Barra de progreso (oculta inicialmente) ---
        self.progress_frame = ttk.Frame(container)
        self.progress_frame.pack(fill=tk.X, pady=(10, 0))
        self.progress_frame.pack_forget()  # Se muestra solo durante la generación

        self.progress_label = ttk.Label(self.progress_frame, text="Generando...")
        self.progress_label.pack(anchor=tk.W)

        self.progress_bar = ttk.Progressbar(
            self.progress_frame, mode='determinate', length=100
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

    # -----------------------------------------------------------------------
    # PESTAÑA 3: HISTORIAL
    # -----------------------------------------------------------------------

    def _create_historial_tab(self):
        """
        Construye la pestaña de historial con:
          - Barra de búsqueda con filtrado en tiempo real mientras se escribe
          - Tabla con todas las actas generadas
          - Menú contextual (clic derecho) con opciones: Abrir, Ver detalles, Eliminar
        """
        container = ttk.Frame(self.tab_historial)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(
            container,
            text="Historial de Actas Generadas",
            font=("Helvetica", 16, "bold")
        ).pack(anchor=tk.W, pady=(0, 20))

        # --- Barra de búsqueda ---
        search_frame = ttk.Frame(container)
        search_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(search_frame, text="Buscar:").pack(side=tk.LEFT, padx=(0, 5))

        self.search_var = tk.StringVar()
        # trace_add dispara _buscar_historial cada vez que cambia el texto
        # esto implementa la búsqueda en tiempo real
        self.search_var.trace_add("write", lambda *_: self._buscar_historial())

        ttk.Entry(
            search_frame, textvariable=self.search_var, width=40
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            search_frame, text="🔄 Actualizar",
            command=self._cargar_historial
        ).pack(side=tk.LEFT)

        # Exporta todo el historial a un archivo .xlsx elegido por el usuario
        ttk.Button(
            search_frame, text="📊 Exportar a Excel",
            command=self._exportar_historial_excel
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Botón para limpiar toda la base de datos (con confirmación)
        ttk.Button(
            search_frame, text="🗑️ Limpiar Base de Datos",
            command=self._limpiar_base_datos
        ).pack(side=tk.LEFT, padx=(10, 0))

        # --- Tabla de historial ---
        cols_historial = (
            "id", "sesion", "fecha_sesion", "fecha_gen",
            "aprobados", "negados", "archivo"
        )
        self.tree_historial = ttk.Treeview(
            container, columns=cols_historial,
            show="headings", height=20
        )

        self.tree_historial.heading("id",          text="ID")
        self.tree_historial.heading("sesion",      text="Sesión")
        self.tree_historial.heading("fecha_sesion",text="Fecha Sesión")
        self.tree_historial.heading("fecha_gen",   text="Fecha Generación")
        self.tree_historial.heading("aprobados",   text="Aprobados")
        self.tree_historial.heading("negados",     text="Negados")
        self.tree_historial.heading("archivo",     text="Archivo")

        self.tree_historial.column("id",           width=50,  anchor=tk.CENTER)
        self.tree_historial.column("sesion",       width=120)
        self.tree_historial.column("fecha_sesion", width=120)
        self.tree_historial.column("fecha_gen",    width=150)
        self.tree_historial.column("aprobados",    width=80,  anchor=tk.CENTER)
        self.tree_historial.column("negados",      width=80,  anchor=tk.CENTER)
        self.tree_historial.column("archivo",      width=400)

        sb_historial = ttk.Scrollbar(
            container, orient=tk.VERTICAL,
            command=self.tree_historial.yview
        )
        self.tree_historial.configure(yscrollcommand=sb_historial.set)
        self.tree_historial.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_historial.pack(side=tk.RIGHT, fill=tk.Y)

        # Clic derecho abre el menú contextual
        self.tree_historial.bind("<Button-3>", self._show_context_menu)

        # Cargar el historial existente al construir la pestaña
        self._cargar_historial()

    # ===========================================================================
    # MÉTODOS DE ACCIÓN - CARGA Y PROCESAMIENTO DE EXCEL
    # ===========================================================================

    def _cargar_excel(self):
        """
        Abre el diálogo de selección de archivo y guarda la ruta elegida
        en self.excel_path para que el Entry la muestre.
        No procesa el archivo todavía, solo lo registra.
        """
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            self.excel_path.set(filepath)
            self.status_bar.config(text=f"Archivo cargado: {filepath}")

    def _procesar_excel(self):
        """
        Lee el archivo Excel cargado, detecta las sesiones disponibles
        y las muestra en la tabla de la pestaña 'Generar Actas'.

        Pasos:
          1. Validar que hay una ruta de archivo
          2. Crear ExcelProcessor y cargar el archivo (con validación de columnas)
          3. Procesar los datos y agrupar por sesión
          4. Poblar el Treeview con una fila por sesión
          5. Marcar en verde las sesiones que ya tienen acta generada en el historial
        """
        filepath = self.excel_path.get()
        if not filepath:
            messagebox.showwarning("Advertencia", "Por favor seleccione un archivo Excel")
            return

        try:
            self.status_bar.config(text="Procesando archivo...")
            self.root.update()

            self.excel_processor = ExcelProcessor(filepath)

            if not self.excel_processor.cargar_archivo():
                # Mostrar el mensaje específico si faltan columnas requeridas
                error_msg = getattr(self.excel_processor, 'error_columnas', None)
                if error_msg:
                    messagebox.showerror(
                        "Columnas faltantes",
                        f"El archivo Excel no tiene el formato correcto:\n\n{error_msg}\n\n"
                        f"Verifique que el archivo tenga todas las columnas requeridas."
                    )
                else:
                    messagebox.showerror("Error", "No se pudo cargar el archivo Excel")
                return

            sesiones = self.excel_processor.procesar_datos()

            # Obtener sesiones que ya tienen acta en el historial
            sesiones_con_acta = {acta.numero_sesion for acta in db_manager.obtener_historial(limite=500)}

            # Tag de color verde para sesiones ya generadas anteriormente
            self.tree_sesiones.tag_configure(
                "ya_generada", background="#d4edda", foreground="#155724"
            )

            for item in self.tree_sesiones.get_children():
                self.tree_sesiones.delete(item)

            self.sesiones_seleccionadas = []

            for sesion_id, sesion in sesiones.items():
                estados_str = ", ".join(sesion.estados_involucrados[:5])
                if len(sesion.estados_involucrados) > 5:
                    estados_str += f" (+{len(sesion.estados_involucrados) - 5} mas)"

                # Las sesiones ya generadas se muestran en verde y pre-marcadas
                ya_generada = sesion_id in sesiones_con_acta

                self.tree_sesiones.insert(
                    "", tk.END,
                    values=(
                        "☑" if ya_generada else "☐",
                        sesion_id,
                        sesion.fecha,
                        sesion.total_aprobados,
                        sesion.total_negados,
                        estados_str,
                    ),
                    tags=("ya_generada",) if ya_generada else ()
                )

            ya_generadas_count = sum(1 for s in sesiones if s in sesiones_con_acta)
            self.status_bar.config(
                text=f"Procesamiento completado. {len(sesiones)} sesiones "
                     f"({ya_generadas_count} ya generadas anteriormente)."
            )
            messagebox.showinfo(
                "Exito",
                f"Se encontraron {len(sesiones)} sesiones en el archivo.\n"
                f"Las sesiones en verde ya tienen acta generada."
            )

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el archivo:\n{str(e)}")
            self.status_bar.config(text="Error en el procesamiento")

    # ===========================================================================
    # MÉTODOS DE ACCIÓN - SELECCIÓN DE SESIONES
    # ===========================================================================

    def _on_sesion_click(self, event):
        """
        Maneja el clic del mouse sobre la tabla de sesiones.

        Solo reacciona si el clic fue en la primera columna (columna ✓).
        Alterna el estado de selección de la fila clickeada:
          - Si estaba desmarcada (☐) → la marca (☑) y agrega a la lista
          - Si estaba marcada (☑)    → la desmarca (☐) y la quita de la lista
        """
        # Verificar que el clic fue sobre una celda
        if self.tree_sesiones.identify_region(event.x, event.y) != "cell":
            return

        # Solo actuar si el clic fue en la primera columna (#1)
        if self.tree_sesiones.identify_column(event.x) != "#1":
            return

        item = self.tree_sesiones.identify_row(event.y)
        if not item:
            return

        values    = self.tree_sesiones.item(item, "values")
        sesion_id = values[1]

        if sesion_id in self.sesiones_seleccionadas:
            # Desmarcar: quitar de la lista y cambiar el icono
            self.sesiones_seleccionadas.remove(sesion_id)
            self.tree_sesiones.item(item, values=("☐",) + values[1:])
        else:
            # Marcar: agregar a la lista y cambiar el icono
            self.sesiones_seleccionadas.append(sesion_id)
            self.tree_sesiones.item(item, values=("☑",) + values[1:])

    def _seleccionar_todas(self):
        """
        Marca todas las sesiones de la tabla como seleccionadas.
        Reemplaza la lista de seleccionadas con todos los IDs disponibles.
        """
        self.sesiones_seleccionadas = []
        for item in self.tree_sesiones.get_children():
            values    = self.tree_sesiones.item(item, "values")
            sesion_id = values[1]
            self.sesiones_seleccionadas.append(sesion_id)
            self.tree_sesiones.item(item, values=("☑",) + values[1:])

    def _deseleccionar_todas(self):
        """
        Desmarca todas las sesiones de la tabla.
        Vacía la lista de seleccionadas.
        """
        self.sesiones_seleccionadas = []
        for item in self.tree_sesiones.get_children():
            values = self.tree_sesiones.item(item, "values")
            self.tree_sesiones.item(item, values=("☐",) + values[1:])

    # ===========================================================================
    # MÉTODOS DE ACCIÓN - GENERACIÓN DE ACTAS
    # ===========================================================================

    def _generar_actas(self):
        """
        Inicia el proceso de generación de actas para las sesiones seleccionadas.

        Antes de lanzar el hilo, verifica si alguna sesión seleccionada ya tiene
        acta generada y pide confirmación al usuario para regenerarla.

        Valida que haya sesiones seleccionadas y que el Excel esté procesado,
        muestra la barra de progreso y lanza el trabajo en un hilo separado
        para no bloquear la interfaz gráfica.
        """
        if not self.sesiones_seleccionadas:
            messagebox.showwarning("Advertencia", "No hay sesiones seleccionadas")
            return

        if not self.excel_processor:
            messagebox.showerror("Error", "Primero debe procesar el archivo Excel")
            return

        # Detectar cuáles sesiones seleccionadas ya tienen acta en el historial
        sesiones_con_acta = {acta.numero_sesion for acta in db_manager.obtener_historial(limite=500)}
        ya_generadas = [s for s in self.sesiones_seleccionadas if s in sesiones_con_acta]

        # Si hay sesiones ya generadas, pedir confirmación antes de continuar
        if ya_generadas:
            lista = "\n".join(f"  - {s}" for s in ya_generadas)
            confirmar = messagebox.askyesno(
                "Sesiones ya generadas",
                f"Las siguientes sesiones ya tienen acta generada:\n\n{lista}\n\n"
                f"\u00bfDesea regenerarlas de todas formas?"
            )
            if not confirmar:
                return

        # Mostrar y configurar la barra de progreso
        self.progress_frame.pack(fill=tk.X, pady=(10, 0))
        self.progress_bar['maximum'] = len(self.sesiones_seleccionadas)
        self.progress_bar['value']   = 0
        self.root.update()

        # Ejecutar en hilo separado para no congelar la UI
        threading.Thread(target=self._generar_actas_thread, daemon=True).start()

    def _generar_actas_thread(self):
        """
        Hilo de trabajo que genera los documentos Word para cada sesión
        seleccionada.

        Por cada sesión:
          1. Obtiene el ResumenSesion del ExcelProcessor
          2. Genera el documento Word con ActaDocumentGenerator
          3. Guarda el .docx en output_actas/
          4. Registra el acta en la tabla 'actas' de la BD
          5. Registra el detalle por estado en 'estadisticas_sesion' (CORREGIDO:
             antes no se llamaba, dejando el dashboard de estados siempre vacío)

        Al terminar, llama a _generacion_completada() en el hilo principal
        usando root.after() para evitar problemas de concurrencia con tkinter.
        """
        generadas = []
        errores   = []

        for i, sesion_id in enumerate(self.sesiones_seleccionadas):
            try:
                # Actualizar la barra de progreso y la etiqueta desde el hilo principal
                self.root.after(0, lambda v=i: self.progress_bar.configure(value=v))
                self.root.after(0, lambda s=sesion_id: self.progress_label.configure(
                    text=f"Generando {s}..."
                ))

                # Obtener los datos de la sesión desde el procesador de Excel
                sesion = self.excel_processor.obtener_resumen_sesion(sesion_id)
                if not sesion:
                    errores.append(f"{sesion_id}: No se encontraron datos")
                    continue

                # Generar el documento Word
                generator = ActaDocumentGenerator()
                generator.generar_acta(sesion)

                # Construir la ruta de salida y guardar el archivo
                filename = generar_nombre_archivo(sesion_id, sesion.fecha)
                filepath = OUTPUT_DIR / filename
                generator.guardar(str(filepath))

                # Guardar el registro principal en la tabla 'actas'
                db_manager.guardar_acta(
                    numero_sesion=sesion_id,
                    fecha_sesion=sesion.fecha,
                    total_aprobados=sesion.total_aprobados,
                    total_negados=sesion.total_negados,
                    estados_involucrados=sesion.estados_involucrados,
                    tipos_solicitud=sesion.tipos_solicitud,
                    archivo_generado=str(filepath),
                    observaciones=""
                )

                # Guardar el detalle por estado y tipo en 'estadisticas_sesion'
                # Esto alimenta el panel "Estadísticas por Estado" del dashboard
                for estado, conteos in sesion.por_estado.items():
                    # Calcular aprobados y negados totales para este estado
                    aprobados_estado = sum(
                        v for k, v in conteos.items() if k.endswith("_aprobados")
                    )
                    negados_estado = sum(
                        v for k, v in conteos.items() if k.endswith("_negados")
                    )
                    # Determinar el tipo de solicitud principal del estado
                    # (puede haber varios; se registra uno por combinación)
                    tipos_en_estado = set(
                        k.replace("_aprobados", "").replace("_negados", "")
                        for k in conteos.keys()
                    )
                    for tipo in tipos_en_estado:
                        aprobados_tipo = conteos.get(f"{tipo}_aprobados", 0)
                        negados_tipo   = conteos.get(f"{tipo}_negados",   0)
                        db_manager.guardar_estadistica_sesion(
                            numero_sesion=sesion_id,
                            estado=estado,
                            tipo_solicitud=tipo,
                            cantidad_aprobados=aprobados_tipo,
                            cantidad_negados=negados_tipo,
                        )

                generadas.append(sesion_id)

            except Exception as e:
                errores.append(f"{sesion_id}: {str(e)}")

        # Notificar al hilo principal que terminó
        self.root.after(0, self._generacion_completada, generadas, errores)

    def _generacion_completada(self, generadas: list, errores: list):
        """
        Callback ejecutado en el hilo principal cuando termina la generación.

        Oculta la barra de progreso, muestra el resumen de resultados,
        actualiza el historial y refresca las estadísticas del dashboard.

        Args:
            generadas: lista de IDs de sesiones generadas exitosamente
            errores  : lista de mensajes de error para sesiones fallidas
        """
        self.progress_frame.pack_forget()

        # Construir el mensaje de resumen
        mensaje = f"Se generaron {len(generadas)} actas:\n"
        mensaje += "\n".join(f"  ✓ {g}" for g in generadas)

        if errores:
            mensaje += f"\n\nErrores ({len(errores)}):\n"
            mensaje += "\n".join(f"  ✗ {e}" for e in errores)

        messagebox.showinfo("Generación Completada", mensaje)
        self.status_bar.config(text=f"Generación completada: {len(generadas)} actas")

        # Refrescar historial y estadísticas con los nuevos datos
        self._cargar_historial()
        self._actualizar_estadisticas()

    # ===========================================================================
    # MÉTODOS DE ACCIÓN - DASHBOARD
    # ===========================================================================

    def _actualizar_estadisticas(self):
        """
        Consulta la base de datos y actualiza todos los widgets del dashboard:
          - Las 4 tarjetas de estadísticas (incluyendo total_sesiones, que antes
            nunca se actualizaba y siempre mostraba 0)
          - La tabla de estadísticas por estado
          - La tabla de actas recientes (últimas 10)
        """
        try:
            stats = db_manager.obtener_estadisticas_globales()

            # Actualizar las 4 tarjetas con los valores de la BD
            for key in ("total_actas", "total_aprobados", "total_negados", "total_sesiones"):
                if key in self.stat_cards:
                    self.stat_cards[key].value_label.config(text=str(stats[key]))

            # Refrescar la tabla de estadísticas por estado
            for item in self.tree_estados.get_children():
                self.tree_estados.delete(item)

            for estado, aprobados, negados in db_manager.obtener_estadisticas_por_estado():
                self.tree_estados.insert(
                    "", tk.END,
                    values=(estado, aprobados, negados, aprobados + negados)
                )

            # Refrescar la tabla de actas recientes (últimas 10)
            for item in self.tree_recientes.get_children():
                self.tree_recientes.delete(item)

            for acta in db_manager.obtener_historial(limite=10):
                self.tree_recientes.insert(
                    "", tk.END,
                    values=(
                        acta.numero_sesion,
                        acta.fecha_sesion,
                        acta.total_aprobados,
                        acta.total_negados,
                    )
                )

        except Exception as e:
            print(f"[Dashboard] Error al actualizar estadísticas: {e}")

    # ===========================================================================
    # MÉTODOS DE ACCIÓN - HISTORIAL
    # ===========================================================================

    def _cargar_historial(self):
        """
        Carga las últimas 100 actas del historial y las muestra en la tabla.
        Se llama al iniciar la app, después de generar actas y al hacer clic
        en el botón Actualizar.
        """
        for item in self.tree_historial.get_children():
            self.tree_historial.delete(item)

        for acta in db_manager.obtener_historial(limite=100):
            self.tree_historial.insert(
                "", tk.END,
                values=(
                    acta.id,
                    acta.numero_sesion,
                    acta.fecha_sesion,
                    acta.fecha_generacion,
                    acta.total_aprobados,
                    acta.total_negados,
                    Path(acta.archivo_generado).name,  # Solo el nombre, no la ruta completa
                )
            )

    def _buscar_historial(self):
        """
        Filtra el historial según el texto del campo de búsqueda.

        Se ejecuta automáticamente cada vez que el usuario escribe en el
        campo de búsqueda (búsqueda en tiempo real via trace_add).

        Si el campo está vacío, muestra todos los registros.
        """
        termino = self.search_var.get().strip()

        if not termino:
            # Sin término de búsqueda: mostrar todo el historial
            self._cargar_historial()
            return

        for item in self.tree_historial.get_children():
            self.tree_historial.delete(item)

        for acta in db_manager.buscar_actas(termino):
            self.tree_historial.insert(
                "", tk.END,
                values=(
                    acta.id,
                    acta.numero_sesion,
                    acta.fecha_sesion,
                    acta.fecha_generacion,
                    acta.total_aprobados,
                    acta.total_negados,
                    Path(acta.archivo_generado).name,
                )
            )

    def _show_context_menu(self, event):
        """
        Muestra un menú contextual al hacer clic derecho sobre una fila
        del historial.

        Opciones disponibles:
          - Abrir archivo  : abre el .docx con la aplicación predeterminada
          - Ver detalles   : muestra una ventana con información completa del acta
          - Eliminar registro: elimina el registro de la BD (pide confirmación)
        """
        item = self.tree_historial.identify_row(event.y)
        if not item:
            return

        self.tree_historial.selection_set(item)

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="📂 Abrir archivo",     command=lambda: self._abrir_archivo(item))
        menu.add_command(label="🔍 Ver detalles",      command=lambda: self._ver_detalles(item))
        menu.add_separator()
        menu.add_command(label="🗑 Eliminar registro", command=lambda: self._eliminar_registro(item))
        menu.post(event.x_root, event.y_root)

    def _abrir_archivo(self, item):
        """
        Abre el archivo .docx del acta seleccionada con la aplicación
        predeterminada del sistema (Microsoft Word u OpenOffice).

        Reconstruye la ruta completa combinando OUTPUT_DIR con el nombre
        del archivo almacenado en la columna 'archivo' de la tabla.

        CORRECCIÓN: antes solo mostraba un messagebox con el nombre del archivo.
        """
        values   = self.tree_historial.item(item, "values")
        filename = values[6]  # Columna 'archivo' (solo el nombre)
        filepath = OUTPUT_DIR / filename

        if not filepath.exists():
            messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontró el archivo:\n{filepath}\n\n"
                "Es posible que haya sido movido o eliminado."
            )
            return

        try:
            # os.startfile abre el archivo con la app predeterminada en Windows
            os.startfile(str(filepath))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{str(e)}")

    def _ver_detalles(self, item):
        """
        Muestra una ventana emergente con la información completa del acta
        seleccionada, incluyendo estados involucrados y tipos de solicitud.

        CORRECCIÓN: antes solo mostraba sesión y fecha. Ahora muestra todos
        los campos relevantes del registro.
        """
        values = self.tree_historial.item(item, "values")
        acta_id = values[0]

        # Buscar el registro completo en la BD para obtener los campos JSON
        acta = db_manager.obtener_acta_por_sesion(values[1])

        if not acta:
            messagebox.showinfo("Detalles", "No se encontró información adicional.")
            return

        # Deserializar los campos JSON almacenados en la BD
        try:
            import json
            estados = json.loads(acta.estados_involucrados) if acta.estados_involucrados else []
            tipos   = json.loads(acta.tipos_solicitud)      if acta.tipos_solicitud      else []
        except Exception:
            estados = []
            tipos   = []

        detalle = (
            f"ID:                  {acta.id}\n"
            f"Sesión:              {acta.numero_sesion}\n"
            f"Fecha de sesión:     {acta.fecha_sesion}\n"
            f"Fecha de generación: {acta.fecha_generacion}\n"
            f"Aprobados:           {acta.total_aprobados}\n"
            f"Negados:             {acta.total_negados}\n"
            f"Archivo:             {Path(acta.archivo_generado).name}\n\n"
            f"Estados involucrados ({len(estados)}):\n"
            f"  {', '.join(estados) if estados else 'N/A'}\n\n"
            f"Tipos de solicitud:\n"
            f"  {', '.join(tipos) if tipos else 'N/A'}\n"
        )
        if acta.observaciones:
            detalle += f"\nObservaciones:\n  {acta.observaciones}"

        messagebox.showinfo(f"Detalles — {acta.numero_sesion}", detalle)

    def _eliminar_registro(self, item):
        """
        Elimina el registro seleccionado del historial previa confirmación.

        No elimina el archivo .docx del disco, solo el registro en la BD.
        """
        values  = self.tree_historial.item(item, "values")
        acta_id = values[0]

        if not messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Eliminar el registro de la sesión {values[1]} del historial?\n\n"
            "El archivo .docx no será eliminado del disco."
        ):
            return

        if db_manager.eliminar_acta(acta_id):
            self.tree_historial.delete(item)
            self._actualizar_estadisticas()
            messagebox.showinfo("Éxito", "Registro eliminado del historial.")
        else:
            messagebox.showerror("Error", "No se pudo eliminar el registro.")

    # ===========================================================================
    # MÉTODOS DE AYUDA
    # ===========================================================================

    def _exportar_historial_excel(self):
        """
        Exporta el historial completo de actas a un archivo Excel (.xlsx).

        Usa pandas para construir el DataFrame y openpyxl para escribir el archivo.
        Si pandas no está disponible muestra un aviso de instalación.
        El usuario elige la ruta de destino mediante un diálogo de guardado.
        """
        if not PANDAS_DISPONIBLE:
            messagebox.showerror(
                "Error",
                "pandas no está instalado.\nEjecute: pip install pandas openpyxl"
            )
            return

        historial = db_manager.obtener_historial(limite=10000)
        if not historial:
            messagebox.showinfo("Sin datos", "No hay actas en el historial para exportar.")
            return

    def _limpiar_base_datos(self):
        """
        Limpia toda la base de datos (borra todos los registros de actas y estadísticas).
        
        Solicita confirmación del usuario antes de proceder.
        Hace un backup automático antes de limpiar.
        """
        # Confirmación del usuario
        respuesta = messagebox.askyesno(
            "⚠️ Confirmar Limpieza",
            "¿Está seguro de que desea eliminar TODOS los registros de la base de datos?\n\n"
            "Esta acción:\n"
            "  • Eliminará todas las actas del historial\n"
            "  • Eliminará todas las estadísticas\n"
            "  • Hacerá un backup automáticamente antes de limpiar\n\n"
            "Esta operación es IRREVERSIBLE.\n\n"
            "¿Desea continuar?",
            icon=messagebox.WARNING
        )
        
        if not respuesta:
            return
        
        # Segunda confirmación (para evitar errores accidentales)
        respuesta2 = messagebox.askyesno(
            "⚠️ ÚLTIMA CONFIRMACIÓN",
            "¿REALMENTE desea eliminar TODOS los datos?\n\n"
            "Esta acción no se puede deshacer.",
            icon=messagebox.WARNING
        )
        
        if not respuesta2:
            return
        
        # Ejecutar la limpieza
        try:
            if db_manager.limpiar_base_datos():
                messagebox.showinfo(
                    "✅ Limpieza Exitosa",
                    "La base de datos ha sido limpiada correctamente.\n\n"
                    "Se ha creado un backup automáticamente en la carpeta de historial."
                )
                # Recargar el historial y estadísticas
                self._cargar_historial()
                self._actualizar_estadisticas()
            else:
                messagebox.showerror(
                    "❌ Error",
                    "No se pudo limpiar la base de datos.\n"
                    "Verifique el log para más detalles."
                )
        except Exception as e:
            messagebox.showerror(
                "❌ Error",
                f"Ocurrió un error al limpiar la base de datos:\n{str(e)}"
            )

    def _exportar_historial_excel(self):
        """
        Exporta el historial completo de actas a un archivo Excel (.xlsx).

        Usa pandas para construir el DataFrame y openpyxl para escribir el archivo.
        Si pandas no está disponible muestra un aviso de instalación.
        El usuario elige la ruta de destino mediante un diálogo de guardado.
        """
        if not PANDAS_DISPONIBLE:
            messagebox.showerror(
                "Error",
                "pandas no está instalado.\nEjecute: pip install pandas openpyxl"
            )
            return

        historial = db_manager.obtener_historial(limite=10000)
        if not historial:
            messagebox.showinfo("Sin datos", "No hay actas en el historial para exportar.")
            return

        # Diálogo para elegir dónde guardar el archivo
        filepath = filedialog.asksaveasfilename(
            title="Guardar historial como Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            initialfile=f"historial_actas_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        if not filepath:
            return

        try:
            datos = []
            for acta in historial:
                # Deserializar los campos JSON para mostrarlos como texto legible
                try:
                    estados = ", ".join(json.loads(acta.estados_involucrados))
                except Exception:
                    estados = acta.estados_involucrados
                try:
                    tipos = ", ".join(json.loads(acta.tipos_solicitud))
                except Exception:
                    tipos = acta.tipos_solicitud

                datos.append({
                    "ID":                 acta.id,
                    "Sesion":             acta.numero_sesion,
                    "Fecha Sesion":       acta.fecha_sesion,
                    "Fecha Generacion":   acta.fecha_generacion,
                    "Aprobados":          acta.total_aprobados,
                    "Negados":            acta.total_negados,
                    "Estados":            estados,
                    "Tipos de Solicitud": tipos,
                    "Archivo":            Path(acta.archivo_generado).name,
                    "Observaciones":      acta.observaciones,
                })

            pd.DataFrame(datos).to_excel(filepath, index=False, sheet_name="Historial")

            self.status_bar.config(text=f"Historial exportado: {Path(filepath).name}")
            messagebox.showinfo(
                "Exportacion completada",
                f"Historial exportado exitosamente:\n{filepath}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el historial:\n{str(e)}")

    def _mostrar_acerca(self):
        """
        Muestra el diálogo 'Acerca de' con información de la aplicación.
        """
        mensaje = (
            f"{APP_NAME}\n"
            f"Versión: {APP_VERSION}\n\n"
            f"Aplicación para la generación automática de actas\n"
            f"del Instituto Nacional de Tierras (INTi).\n\n"
            f"Desarrollada para automatizar el proceso de creación\n"
            f"de actas a partir del archivo Excel de Cargas Atancha.\n\n"
            f"Características:\n"
            f"• Procesamiento automático de archivos Excel\n"
            f"• Generación de actas en formato Word\n"
            f"• Dashboard con estadísticas en tiempo real\n"
            f"• Historial de actas con búsqueda en tiempo real\n"
            f"• Configuración editable desde config.json\n"
        )
        messagebox.showinfo("Acerca de", mensaje)

    # ===========================================================================
    # PUNTO DE ENTRADA
    # ===========================================================================

    def run(self):
        """Inicia el loop principal de la interfaz gráfica."""
        self.root.mainloop()


# ===========================================================================
# PUNTO DE ENTRADA DEL SCRIPT
# ===========================================================================

def main():
    """
    Función principal. Crea la instancia de la aplicación y la ejecuta.
    Se llama desde iniciar.py o directamente con: python main_app.py
    """
    app = ModernActaGeneratorApp()
    app.run()


if __name__ == "__main__":
    main()
