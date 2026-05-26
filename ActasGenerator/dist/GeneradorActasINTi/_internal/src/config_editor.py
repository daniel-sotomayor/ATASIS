"""
config_editor.py - Ventana de configuración editable del Generador de Actas INTi

Permite al usuario modificar desde la interfaz gráfica todos los valores
de config.json sin necesidad de abrir el archivo manualmente:

  - Presidente del directorio
  - Directores principales (hasta 4)
  - Directores suplentes (hasta 4)
  - Secretaria del directorio
  - Hora de inicio y fin de sesión
  - Número de reunión continua

Al guardar, escribe el config.json actualizado y recarga config.py en memoria
para que los cambios se reflejen en la próxima acta generada sin reiniciar la app.

Uso desde main_app.py:
    from config_editor import abrir_configuracion
    abrir_configuracion(parent_window)
"""

import json
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

from config import CONFIG_JSON_PATH


# ---------------------------------------------------------------------------
# FUNCIÓN PÚBLICA
# ---------------------------------------------------------------------------

def abrir_configuracion(parent: tk.Tk):
    """
    Abre la ventana de configuración como diálogo modal.

    Args:
        parent: ventana principal de la aplicación (para centrar el diálogo)
    """
    ventana = ConfigEditorWindow(parent)
    ventana.window.grab_set()       # Bloquear la ventana principal mientras está abierta
    parent.wait_window(ventana.window)  # Esperar a que se cierre


# ---------------------------------------------------------------------------
# CLASE PRINCIPAL
# ---------------------------------------------------------------------------

class ConfigEditorWindow:
    """
    Ventana modal para editar la configuración del sistema.

    Lee config.json al abrirse y lo escribe al guardar.
    Recarga los valores en config.py para que los cambios sean inmediatos.
    """

    def __init__(self, parent: tk.Tk):
        """
        Construye la ventana y carga los valores actuales de config.json.
        """
        self.parent = parent

        # Crear ventana secundaria (Toplevel)
        self.window = tk.Toplevel(parent)
        self.window.title("Configuración del Sistema")
        self.window.geometry("700x650")
        self.window.resizable(False, False)
        self._centrar_ventana()

        # Cargar configuración actual
        self.cfg = self._leer_config()

        # Construir la interfaz
        self._build_ui()

    def _centrar_ventana(self):
        """Centra la ventana sobre la ventana principal."""
        self.window.update_idletasks()
        pw = self.parent.winfo_x() + self.parent.winfo_width()  // 2
        ph = self.parent.winfo_y() + self.parent.winfo_height() // 2
        w  = 700
        h  = 650
        self.window.geometry(f"{w}x{h}+{pw - w//2}+{ph - h//2}")

    def _leer_config(self) -> dict:
        """
        Lee y retorna el contenido actual de config.json.
        Si el archivo no existe o tiene error, retorna un dict vacío.
        """
        if not CONFIG_JSON_PATH.exists():
            return {}
        try:
            with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    # -----------------------------------------------------------------------
    # CONSTRUCCIÓN DE LA INTERFAZ
    # -----------------------------------------------------------------------

    def _build_ui(self):
        """
        Construye todos los widgets de la ventana organizados en secciones:
          1. Directorio (presidente, directores, secretaria)
          2. Formato del acta (horas, número de reunión)
          3. Botones Guardar / Cancelar
        """
        main_frame = ttk.Frame(self.window, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main_frame,
            text="Configuración del Sistema",
            font=("Helvetica", 14, "bold")
        ).pack(anchor=tk.W, pady=(0, 15))

        # Notebook con dos pestañas para organizar los campos
        nb = ttk.Notebook(main_frame)
        nb.pack(fill=tk.BOTH, expand=True)

        tab_directorio = ttk.Frame(nb, padding=10)
        tab_formato    = ttk.Frame(nb, padding=10)

        nb.add(tab_directorio, text=" Miembros del Directorio ")
        nb.add(tab_formato,    text=" Formato del Acta ")

        self._build_tab_directorio(tab_directorio)
        self._build_tab_formato(tab_formato)

        # Botones de acción
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(
            btn_frame, text="💾 Guardar cambios",
            command=self._guardar
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(
            btn_frame, text="Cancelar",
            command=self.window.destroy
        ).pack(side=tk.RIGHT)

    def _build_tab_directorio(self, parent):
        """
        Construye la pestaña de miembros del directorio con campos para:
          - Presidente
          - 4 directores principales
          - 4 directores suplentes
          - Secretaria
        """
        dir_cfg = self.cfg.get("directorio", {})

        # --- Presidente ---
        ttk.Label(parent, text="Presidente:", font=("Helvetica", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5)
        )
        self.var_presidente = tk.StringVar(
            value=dir_cfg.get("presidente", "")
        )
        ttk.Entry(parent, textvariable=self.var_presidente, width=55).grid(
            row=0, column=1, sticky=tk.W, pady=(0, 5), padx=(10, 0)
        )

        # --- Directores principales ---
        ttk.Label(parent, text="Directores Principales:", font=("Helvetica", 10, "bold")).grid(
            row=1, column=0, sticky=tk.NW, pady=(10, 5)
        )
        principales = dir_cfg.get("directores_principales", ["", "", "", ""])
        self.vars_principales = []
        for i in range(4):
            var = tk.StringVar(value=principales[i] if i < len(principales) else "")
            self.vars_principales.append(var)
            ttk.Label(parent, text=f"  {i+1}.").grid(
                row=2 + i, column=0, sticky=tk.W, padx=(15, 0)
            )
            ttk.Entry(parent, textvariable=var, width=55).grid(
                row=2 + i, column=1, sticky=tk.W, pady=2, padx=(10, 0)
            )

        # --- Directores suplentes ---
        ttk.Label(parent, text="Directores Suplentes:", font=("Helvetica", 10, "bold")).grid(
            row=6, column=0, sticky=tk.NW, pady=(10, 5)
        )
        suplentes = dir_cfg.get("directores_suplentes", ["", "", "", ""])
        self.vars_suplentes = []
        for i in range(4):
            var = tk.StringVar(value=suplentes[i] if i < len(suplentes) else "")
            self.vars_suplentes.append(var)
            ttk.Label(parent, text=f"  {i+1}.").grid(
                row=7 + i, column=0, sticky=tk.W, padx=(15, 0)
            )
            ttk.Entry(parent, textvariable=var, width=55).grid(
                row=7 + i, column=1, sticky=tk.W, pady=2, padx=(10, 0)
            )

        # --- Secretaria ---
        ttk.Label(parent, text="Secretaria:", font=("Helvetica", 10, "bold")).grid(
            row=11, column=0, sticky=tk.W, pady=(10, 5)
        )
        self.var_secretaria = tk.StringVar(
            value=dir_cfg.get("secretaria", "")
        )
        ttk.Entry(parent, textvariable=self.var_secretaria, width=55).grid(
            row=11, column=1, sticky=tk.W, pady=(10, 5), padx=(10, 0)
        )

    def _build_tab_formato(self, parent):
        """
        Construye la pestaña de formato del acta con campos para:
          - Hora de inicio de la sesión
          - Hora de fin de la sesión
          - Número de reunión continua del directorio
        """
        fmt_cfg = self.cfg.get("formato_acta", {})

        campos = [
            ("Hora de inicio:",       "hora_inicio",    "9:00 a.m."),
            ("Hora de fin:",          "hora_fin",       "11:30 a.m."),
            ("Número de reunión:",    "reunion_numero", "98"),
        ]

        self.vars_formato = {}
        for i, (label, key, default) in enumerate(campos):
            ttk.Label(parent, text=label, font=("Helvetica", 10, "bold")).grid(
                row=i, column=0, sticky=tk.W, pady=8
            )
            var = tk.StringVar(value=fmt_cfg.get(key, default))
            self.vars_formato[key] = var
            ttk.Entry(parent, textvariable=var, width=30).grid(
                row=i, column=1, sticky=tk.W, padx=(15, 0), pady=8
            )

        # Nota informativa
        nota = (
            "Nota: Los cambios se aplican a partir de la próxima acta generada.\n"
            "No afectan actas ya creadas."
        )
        ttk.Label(
            parent, text=nota,
            font=("Helvetica", 9), foreground="gray"
        ).grid(row=len(campos) + 1, column=0, columnspan=2, sticky=tk.W, pady=(20, 0))

    # -----------------------------------------------------------------------
    # GUARDAR
    # -----------------------------------------------------------------------

    def _guardar(self):
        """
        Valida los campos, escribe config.json con los nuevos valores
        y recarga los valores en config.py para que sean inmediatos.

        Validaciones:
          - El campo Presidente no puede estar vacío
          - Debe haber al menos 1 director principal
          - La Secretaria no puede estar vacía
        """
        # Recoger valores de la pestaña Directorio
        presidente = self.var_presidente.get().strip()
        principales = [v.get().strip() for v in self.vars_principales if v.get().strip()]
        suplentes   = [v.get().strip() for v in self.vars_suplentes   if v.get().strip()]
        secretaria  = self.var_secretaria.get().strip()

        # Validaciones básicas
        if not presidente:
            messagebox.showwarning("Campo requerido", "El nombre del Presidente no puede estar vacío.")
            return
        if not principales:
            messagebox.showwarning("Campo requerido", "Debe haber al menos un Director Principal.")
            return
        if not secretaria:
            messagebox.showwarning("Campo requerido", "El nombre de la Secretaria no puede estar vacío.")
            return

        # Recoger valores de la pestaña Formato
        hora_inicio    = self.vars_formato["hora_inicio"].get().strip()
        hora_fin       = self.vars_formato["hora_fin"].get().strip()
        reunion_numero = self.vars_formato["reunion_numero"].get().strip()

        # Construir el nuevo diccionario de configuración
        # Preservar las secciones que no se editan (app, excel)
        nueva_cfg = dict(self.cfg)  # Copia de la config actual
        nueva_cfg["directorio"] = {
            "presidente":             presidente,
            "directores_principales": principales,
            "directores_suplentes":   suplentes,
            "secretaria":             secretaria,
        }
        nueva_cfg["formato_acta"] = {
            "hora_inicio":    hora_inicio,
            "hora_fin":       hora_fin,
            "reunion_numero": reunion_numero,
        }

        # Escribir config.json
        try:
            with open(CONFIG_JSON_PATH, "w", encoding="utf-8") as f:
                json.dump(nueva_cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la configuración:\n{e}")
            return

        # Recargar los valores en config.py para que sean inmediatos
        # sin necesidad de reiniciar la aplicación
        self._recargar_config_en_memoria(nueva_cfg)

        messagebox.showinfo(
            "Configuración guardada",
            "Los cambios se guardaron correctamente.\n"
            "Se aplicarán en la próxima acta generada."
        )
        self.window.destroy()

    def _recargar_config_en_memoria(self, nueva_cfg: dict):
        """
        Actualiza las variables de config.py en memoria con los nuevos valores.

        Esto evita tener que reiniciar la aplicación para que los cambios
        en config.json surtan efecto.

        Args:
            nueva_cfg: diccionario con la configuración recién guardada
        """
        import config  # Importar el módulo para modificar sus variables globales

        dir_cfg = nueva_cfg.get("directorio", {})
        fmt_cfg = nueva_cfg.get("formato_acta", {})

        # Actualizar MIEMBROS_DIRECTORIO
        config.MIEMBROS_DIRECTORIO["presidente"]             = dir_cfg.get("presidente", config.MIEMBROS_DIRECTORIO["presidente"])
        config.MIEMBROS_DIRECTORIO["directores_principales"] = dir_cfg.get("directores_principales", config.MIEMBROS_DIRECTORIO["directores_principales"])
        config.MIEMBROS_DIRECTORIO["directores_suplentes"]   = dir_cfg.get("directores_suplentes",   config.MIEMBROS_DIRECTORIO["directores_suplentes"])
        config.MIEMBROS_DIRECTORIO["secretaria"]             = dir_cfg.get("secretaria",             config.MIEMBROS_DIRECTORIO["secretaria"])

        # Actualizar horas y número de reunión
        config.HORA_INICIO_DEFAULT = fmt_cfg.get("hora_inicio",    config.HORA_INICIO_DEFAULT)
        config.HORA_FIN_DEFAULT    = fmt_cfg.get("hora_fin",       config.HORA_FIN_DEFAULT)
        config.REUNION_NUMERO      = fmt_cfg.get("reunion_numero", config.REUNION_NUMERO)
