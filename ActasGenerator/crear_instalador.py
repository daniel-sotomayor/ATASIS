"""
crear_instalador.py - Genera el instalador completo del Generador de Actas INTi

Pasos que ejecuta:
  1. Convierte logo.jpeg a logo.ico
  2. Regenera el .exe con PyInstaller (con el ícono correcto)
  3. Compila el instalador .exe con Inno Setup

El instalador final queda en:
    instalador/Instalador_GeneradorActasINTi_v1.0.0.exe

Uso:
    python crear_instalador.py
"""

import subprocess
import sys
import shutil
from pathlib import Path

BASE_DIR   = Path(__file__).parent
SRC_DIR    = BASE_DIR / "src"
ASSETS_DIR = BASE_DIR / "assets"
DIST_DIR   = BASE_DIR / "dist" / "GeneradorActasINTi"
BUILD_DIR  = BASE_DIR / "build"
INST_DIR   = BASE_DIR / "instalador"
APP_NAME   = "GeneradorActasINTi"
ICON_PATH  = ASSETS_DIR / "logo.ico"
ISS_FILE   = BASE_DIR / "instalador.iss"

# Rutas posibles de Inno Setup en Windows
INNO_PATHS = [
    Path(r"C:\Users\usuario\AppData\Local\Programs\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
]


def paso1_generar_icono():
    """Convierte logo.jpeg a logo.ico con múltiples tamaños."""
    print("[1/3] Generando icono .ico desde logo.jpeg...")
    try:
        from PIL import Image
        img = Image.open(ASSETS_DIR / "logo.jpeg").convert("RGBA")
        img.save(ICON_PATH, format="ICO",
                 sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])
        print(f"      Icono generado: {ICON_PATH.name} ({ICON_PATH.stat().st_size} bytes)")
        return True
    except Exception as e:
        print(f"      ERROR: {e}")
        return False


def paso2_generar_exe():
    """Regenera el .exe con PyInstaller incluyendo el ícono."""
    print("\n[2/3] Generando ejecutable con PyInstaller...")

    # Limpiar builds anteriores
    for carpeta in [BUILD_DIR, BASE_DIR / "dist"]:
        if carpeta.exists():
            shutil.rmtree(carpeta)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", APP_NAME,
        f"--icon={ICON_PATH}",
        "--distpath", str(BASE_DIR / "dist"),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(BASE_DIR),
        f"--add-data={SRC_DIR};src",
        f"--add-data={ASSETS_DIR};assets",
        f"--add-data={BASE_DIR / 'config.json'};.",
        f"--paths={SRC_DIR}",
        "--hidden-import=config",
        "--hidden-import=database",
        "--hidden-import=excel_processor",
        "--hidden-import=document_generator",
        "--hidden-import=main_app",
        "--hidden-import=config_editor",
        "--hidden-import=ttkbootstrap",
        "--hidden-import=openpyxl",
        "--hidden-import=docx",
        "--hidden-import=pandas",
        "--hidden-import=PIL",
        str(BASE_DIR / "iniciar.py"),
    ]

    resultado = subprocess.run(cmd, cwd=str(BASE_DIR))
    if resultado.returncode != 0:
        print("      ERROR: PyInstaller falló.")
        return False

    # Copiar recursos adicionales al dist
    INST_DIR.mkdir(exist_ok=True)
    for carpeta in ["historial", "output_actas", "reports",
                    "historial/backup", "assets"]:
        (DIST_DIR / carpeta).mkdir(parents=True, exist_ok=True)

    # Copiar logo.ico al dist/assets para que el instalador lo use
    shutil.copy2(ICON_PATH, DIST_DIR / "assets" / "logo.ico")
    shutil.copy2(ASSETS_DIR / "logo.jpeg", DIST_DIR / "assets" / "logo.jpeg")

    print(f"      Ejecutable listo: {DIST_DIR / (APP_NAME + '.exe')}")
    return True


def paso3_compilar_instalador():
    """Compila el instalador .exe con Inno Setup."""
    print("\n[3/3] Compilando instalador con Inno Setup...")

    # Buscar ISCC.exe
    iscc = next((p for p in INNO_PATHS if p.exists()), None)
    if not iscc:
        print("      ERROR: No se encontró Inno Setup.")
        print("      Instale desde: https://jrsoftware.org/isdl.php")
        return False

    INST_DIR.mkdir(exist_ok=True)

    resultado = subprocess.run(
        [str(iscc), str(ISS_FILE)],
        cwd=str(BASE_DIR)
    )

    if resultado.returncode != 0:
        print("      ERROR: Inno Setup falló.")
        return False

    # Buscar el instalador generado
    instaladores = list(INST_DIR.glob("Instalador_*.exe"))
    if instaladores:
        inst = instaladores[0]
        size_mb = inst.stat().st_size / (1024 * 1024)
        print(f"      Instalador listo: {inst.name} ({size_mb:.1f} MB)")

    return True


def main():
    print("=" * 60)
    print("  GENERADOR DE INSTALADOR - Actas INTi")
    print("=" * 60)

    if not paso1_generar_icono():
        input("\nPresione Enter para salir...")
        sys.exit(1)

    if not paso2_generar_exe():
        input("\nPresione Enter para salir...")
        sys.exit(1)

    if not paso3_compilar_instalador():
        input("\nPresione Enter para salir...")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  INSTALADOR GENERADO EXITOSAMENTE")
    print("=" * 60)
    print(f"\n  Ubicacion: {INST_DIR}")
    print("\n  El instalador:")
    print("    - Instala en C:\\Program Files\\Generador de Actas INTi\\")
    print("    - Crea acceso directo en el Escritorio con el logo")
    print("    - Crea acceso directo en el Menu Inicio")
    print("    - Crea la carpeta Documentos\\Actas de ATASIS")
    print("    - Aparece en 'Agregar o quitar programas'")
    print()

    input("Presione Enter para salir...")


if __name__ == "__main__":
    main()
