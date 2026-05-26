"""
crear_ejecutable.py - Genera el archivo .exe del Generador de Actas INTi

Usa PyInstaller para empaquetar toda la aplicación en un único ejecutable
portable que no requiere tener Python instalado.

El ejecutable generado se guarda en:
    dist/GeneradorActasINTi/GeneradorActasINTi.exe

Uso:
    python crear_ejecutable.py

Requisitos:
    pip install pyinstaller
"""

import subprocess
import sys
import shutil
from pathlib import Path

# ---------------------------------------------------------------------------
# RUTAS
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).parent
SRC_DIR    = BASE_DIR / "src"
ASSETS_DIR = BASE_DIR / "assets"
DIST_DIR   = BASE_DIR / "dist" / "GeneradorActasINTi"
BUILD_DIR  = BASE_DIR / "build"

APP_NAME   = "GeneradorActasINTi"
ICON_PATH  = ASSETS_DIR / "logo.ico"


def limpiar_builds_anteriores():
    """Elimina carpetas build/ y dist/ de ejecuciones anteriores."""
    for carpeta in [BUILD_DIR, BASE_DIR / "dist"]:
        if carpeta.exists():
            shutil.rmtree(carpeta)
            print(f"  Eliminado: {carpeta}")


def construir_exe():
    """
    Ejecuta PyInstaller con los parámetros necesarios para empaquetar
    la aplicación correctamente.

    Parámetros importantes:
      --onedir      : genera una carpeta con el exe y sus dependencias
                      (más rápido de iniciar que --onefile)
      --windowed    : no muestra consola negra al abrir el exe
      --add-data    : incluye carpetas de recursos (assets, config.json)
      --hidden-import: fuerza la inclusión de módulos que PyInstaller
                       no detecta automáticamente
      --name        : nombre del ejecutable generado
    """
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",                          # Sobreescribir sin preguntar
        "--onedir",                             # Carpeta con exe + dependencias
        "--windowed",                           # Sin consola negra
        f"--icon={ICON_PATH}",                   # Icono del .exe (logo INTi)
        "--name", APP_NAME,
        "--distpath", str(BASE_DIR / "dist"),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(BASE_DIR),

        # Incluir carpeta src/ completa para que los módulos estén disponibles
        f"--add-data={SRC_DIR};src",

        # Incluir carpeta de assets (logo)
        f"--add-data={ASSETS_DIR};assets",

        # Incluir config.json en la raíz del exe
        f"--add-data={BASE_DIR / 'config.json'};.",

        # Agregar src/ al path de búsqueda de módulos
        f"--paths={SRC_DIR}",

        # Módulos del proyecto (PyInstaller no los detecta automáticamente)
        "--hidden-import=config",
        "--hidden-import=database",
        "--hidden-import=excel_processor",
        "--hidden-import=document_generator",
        "--hidden-import=main_app",
        "--hidden-import=config_editor",

        # Dependencias externas
        "--hidden-import=ttkbootstrap",
        "--hidden-import=openpyxl",
        "--hidden-import=docx",
        "--hidden-import=pandas",
        "--hidden-import=PIL",

        # Punto de entrada
        str(BASE_DIR / "iniciar.py"),
    ]

    print("\nEjecutando PyInstaller...")
    print("Comando:", " ".join(cmd))
    print()

    resultado = subprocess.run(cmd, cwd=str(BASE_DIR))
    return resultado.returncode == 0


def copiar_recursos():
    """
    Copia los archivos y carpetas necesarios para que el exe funcione
    correctamente al ejecutarse desde dist/GeneradorActasINTi/:
      - config.json  (configuración editable por el usuario)
      - historial/   (se crea vacía para la BD)
      - output_actas/(se crea vacía para las actas generadas)
      - reports/     (se crea vacía para reportes)
    """
    if not DIST_DIR.exists():
        print("ERROR: No se encontró la carpeta dist/. El build falló.")
        return False

    # Crear carpetas de trabajo vacías dentro del dist
    for carpeta in ["historial", "output_actas", "reports",
                    "historial/backup", "assets"]:
        (DIST_DIR / carpeta).mkdir(parents=True, exist_ok=True)

    # Copiar config.json si no fue incluido por PyInstaller
    config_src = BASE_DIR / "config.json"
    config_dst = DIST_DIR / "config.json"
    if config_src.exists() and not config_dst.exists():
        shutil.copy2(config_src, config_dst)
        print(f"  Copiado: config.json")

    # Copiar logo
    logo_src = ASSETS_DIR / "logo.jpeg"
    logo_dst = DIST_DIR / "assets" / "logo.jpeg"
    if logo_src.exists() and not logo_dst.exists():
        shutil.copy2(logo_src, logo_dst)
        print(f"  Copiado: assets/logo.jpeg")

    # Crear .bat de inicio dentro del dist para facilitar el uso
    bat_content = (
        "@echo off\n"
        "cd /d \"%~dp0\"\n"
        f"start GeneradorActasINTi.exe\n"
    )
    (DIST_DIR / "Iniciar.bat").write_text(bat_content, encoding="utf-8")
    print("  Creado: Iniciar.bat")

    return True


def main():
    print("=" * 60)
    print(f"  Generando ejecutable: {APP_NAME}.exe")
    print("=" * 60)

    print("\n[1/3] Limpiando builds anteriores...")
    limpiar_builds_anteriores()

    print("\n[2/3] Construyendo ejecutable...")
    if not construir_exe():
        print("\nERROR: PyInstaller falló. Revise los mensajes anteriores.")
        input("\nPresione Enter para salir...")
        sys.exit(1)

    print("\n[3/3] Copiando recursos adicionales...")
    if not copiar_recursos():
        input("\nPresione Enter para salir...")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  EJECUTABLE GENERADO EXITOSAMENTE")
    print("=" * 60)
    print(f"\n  Ubicacion: {DIST_DIR}")
    print(f"  Ejecutable: {DIST_DIR / (APP_NAME + '.exe')}")
    print("\n  Para distribuir, copie toda la carpeta:")
    print(f"  dist/{APP_NAME}/")
    print("\n  El usuario solo necesita ejecutar:")
    print(f"  {APP_NAME}.exe  o  Iniciar.bat")
    print()

    input("Presione Enter para salir...")


if __name__ == "__main__":
    main()
