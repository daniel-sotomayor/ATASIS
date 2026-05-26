"""
config.py - Configuración central de la aplicación Generador de Actas INTi

Este módulo es el único punto de verdad para toda la configuración.
Lee los datos dinámicos (miembros del directorio, horas, número de reunión)
desde el archivo config.json ubicado en la raíz del proyecto, de modo que
cualquier cambio en ese archivo se refleja automáticamente sin tocar código.

Si config.json no existe o tiene un error, se usan valores por defecto
para que la aplicación no falle al iniciar.
"""

import json
import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# RUTAS BASE DEL PROYECTO
# ---------------------------------------------------------------------------

# Detectar si corremos dentro de un .exe de PyInstaller o como script normal
if getattr(sys, 'frozen', False):
    # Dentro del .exe: usar la carpeta donde está el ejecutable
    BASE_DIR = Path(sys.executable).parent
else:
    # Como script: un nivel arriba de /src
    BASE_DIR = Path(__file__).parent.parent

# Carpetas de trabajo
ASSETS_DIR    = BASE_DIR / "assets"        # Recursos gráficos (logos, iconos)

# Usar AppData para historial y reports para tener permisos de escritura
# cuando la app está instalada en Program Files
APP_DATA_DIR = Path.home() / "AppData" / "Local" / "Generador de Actas INTi"
HISTORIAL_DIR = APP_DATA_DIR / "historial"    # Base de datos SQLite
REPORTS_DIR   = APP_DATA_DIR / "reports"       # Reportes exportados
DB_PATH       = HISTORIAL_DIR / "actas_historial.db"

# Las actas se guardan en Documentos/Actas de ATASIS para que el usuario
# las encuentre fácilmente sin importar desde dónde se ejecute la app
OUTPUT_DIR = Path.home() / "Documents" / "Actas de ATASIS"

# Ruta al archivo de configuración editable por el usuario
# Usar AppData para tener permisos de escritura cuando está instalado en Program Files
CONFIG_JSON_PATH = APP_DATA_DIR / "config.json"

# Crear las carpetas si no existen (se ejecuta al importar el módulo)
for dir_path in [ASSETS_DIR, OUTPUT_DIR, HISTORIAL_DIR, REPORTS_DIR, APP_DATA_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# Copiar config.json por defecto de BASE_DIR a APP_DATA_DIR si no existe
# Esto permite que la configuración por defecto se cargue la primera vez
# y luego el usuario pueda modificarla en AppData (donde tiene permisos de escritura)
DEFAULT_CONFIG_PATH = BASE_DIR / "config.json"
if not CONFIG_JSON_PATH.exists() and DEFAULT_CONFIG_PATH.exists():
    import shutil
    try:
        shutil.copy2(DEFAULT_CONFIG_PATH, CONFIG_JSON_PATH)
    except Exception as e:
        print(f"[config] Advertencia: no se pudo copiar config.json por defecto: {e}")

# ---------------------------------------------------------------------------
# LEER config.json EN TIEMPO DE EJECUCIÓN
# ---------------------------------------------------------------------------

def _cargar_config_json() -> dict:
    """
    Lee y retorna el contenido de config.json.

    Si el archivo no existe o tiene un error de sintaxis, retorna un
    diccionario vacío para que el resto del módulo use los valores por defecto.
    """
    if not CONFIG_JSON_PATH.exists():
        print(f"[config] Advertencia: no se encontró {CONFIG_JSON_PATH}. Usando valores por defecto.")
        return {}
    try:
        with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print(f"[config] Error al leer config.json: {e}. Usando valores por defecto.")
        return {}

# Cargar el JSON una sola vez al importar el módulo
_cfg = _cargar_config_json()

# ---------------------------------------------------------------------------
# CONFIGURACIÓN DE LA APLICACIÓN
# ---------------------------------------------------------------------------

APP_NAME    = _cfg.get("app", {}).get("name",    "Generador de Actas INTi")
APP_VERSION = _cfg.get("app", {}).get("version", "1.0.0")
APP_WIDTH   = 1400
APP_HEIGHT  = 900

# ---------------------------------------------------------------------------
# CONFIGURACIÓN DEL FORMATO DEL ACTA
# Estos valores se leen desde config.json para que el usuario pueda
# cambiarlos sin modificar código fuente.
# ---------------------------------------------------------------------------

_fmt = _cfg.get("formato_acta", {})

# Hora de inicio de la sesión (aparece en el encabezado del acta)
HORA_INICIO_DEFAULT = _fmt.get("hora_inicio", "9:00 a.m.")

# Hora de cierre de la sesión (aparece en la sección de lectura y aprobación)
HORA_FIN_DEFAULT = _fmt.get("hora_fin", "11:30 a.m.")

# Número de reunión continua del directorio (aparece en el encabezado)
REUNION_NUMERO = _fmt.get("reunion_numero", "98")

# ---------------------------------------------------------------------------
# MIEMBROS DEL DIRECTORIO
# Se leen desde config.json. Si no existe la clave, se usan los valores
# por defecto definidos aquí como respaldo.
# ---------------------------------------------------------------------------

_dir_cfg = _cfg.get("directorio", {})

MIEMBROS_DIRECTORIO = {
    "presidente": _dir_cfg.get(
        "presidente",
        "David José Hernández Giménez"
    ),
    "directores_principales": _dir_cfg.get(
        "directores_principales",
        [
            "Faiez Kassen Castillo",
            "Giomar Gisela Blanco Espinoza",
            "Lusmialit Lliliana Perdomo Torrealba",
            "David Arturo Gómez Rodríguez",
        ]
    ),
    "directores_suplentes": _dir_cfg.get(
        "directores_suplentes",
        [
            "Virginia Coromoto Delgado Oropeza",
            "Francisco Ramon Ojeda Barrera",
            "Alfredo José Mora Álvarez",
            "Gisela Ninoska Faraco De Cuellar",
        ]
    ),
    "secretaria": _dir_cfg.get(
        "secretaria",
        "Viviana Del Carmen Rivas Tovar"
    ),
}

# ---------------------------------------------------------------------------
# TIPOS DE SOLICITUD RECONOCIDOS
# Mapeo de nombre completo → código interno usado en estadísticas
# ---------------------------------------------------------------------------

TIPOS_SOLICITUD = {
    "Adjudicación de Tierras":               "ADJUDICACION",
    "Declaratoria de Garantía de Permanencia": "DECLARATORIA",
    "Liberación de Predio - Revocatoria":    "LIBERACION",
}

# ---------------------------------------------------------------------------
# ESTADOS DE VENEZUELA (orden alfabético oficial)
# ---------------------------------------------------------------------------

ESTADOS_VENEZUELA = [
    "Amazonas", "Anzoátegui", "Apure", "Aragua", "Barinas",
    "Bolívar", "Carabobo", "Cojedes", "Delta Amacuro",
    "Distrito Capital", "Falcón", "Guárico", "La Guaira",
    "Lara", "Mérida", "Miranda", "Monagas", "Nueva Esparta",
    "Portuguesa", "Sucre", "Táchira", "Trujillo", "Yaracuy",
    "Zulia norte", "Zulia sur",
]

# ---------------------------------------------------------------------------
# PALETA DE COLORES DE LA INTERFAZ
# Usados por main_app.py para estilos cuando ttkbootstrap no está disponible
# ---------------------------------------------------------------------------

COLORS = {
    "primary":    "#1a5f7a",
    "secondary":  "#57c5b6",
    "accent":     "#159895",
    "background": "#f5f5f5",
    "text":       "#2c3e50",
    "success":    "#27ae60",
    "warning":    "#f39c12",
    "danger":     "#e74c3c",
    "info":       "#3498db",
}
