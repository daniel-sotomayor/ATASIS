"""
iniciar.py - Punto de entrada principal del Generador de Actas INTi

Este script es el que ejecuta el .bat y el usuario directamente.
Se encarga de:
  1. Agregar la carpeta src/ al sys.path para que los imports funcionen
     sin importar desde qué directorio se ejecute el script
  2. Verificar que todas las dependencias estén instaladas antes de
     intentar arrancar la interfaz gráfica
  3. Capturar cualquier error inesperado al iniciar y mostrarlo de forma
     legible en lugar de que la ventana se cierre sola sin explicación

Uso:
    python iniciar.py
    (o doble clic en "Iniciar Generador de Actas.bat")
"""

import sys
import os
from pathlib import Path

# ---------------------------------------------------------------------------
# PASO 1: Detectar si corremos como .exe (PyInstaller) o como script normal
# y agregar src/ al path correctamente en ambos casos.
# ---------------------------------------------------------------------------

# sys.frozen es True cuando la app corre dentro de un .exe de PyInstaller
if getattr(sys, 'frozen', False):
    # Dentro del .exe: la raíz es la carpeta donde está el .exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Como script normal: la raíz es donde está iniciar.py
    BASE_DIR = Path(__file__).parent

src_dir = BASE_DIR / "src"

if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

# Cambiar el directorio de trabajo a la raíz para que las rutas de config.py
# funcionen correctamente tanto en script como en .exe
os.chdir(BASE_DIR)


# ---------------------------------------------------------------------------
# PASO 2: Verificar dependencias antes de intentar importar la app
# ---------------------------------------------------------------------------

def verificar_dependencias() -> bool:
    """
    Comprueba que todas las librerías requeridas estén instaladas.

    Retorna True si todo está bien, False si falta alguna dependencia.
    En caso de faltar dependencias, imprime instrucciones claras para
    que el usuario sepa exactamente qué instalar.
    """
    # Mapa: nombre_import → nombre_paquete_pip
    dependencias = {
        "pandas":      "pandas",
        "docx":        "python-docx",
        "openpyxl":    "openpyxl",
        "ttkbootstrap": "ttkbootstrap",  # Opcional pero recomendado
    }

    faltantes_criticas = []   # Sin estas la app no arranca
    faltantes_opcionales = [] # Sin estas la app arranca con funciones reducidas

    criticas   = {"pandas", "docx", "openpyxl"}
    opcionales = {"ttkbootstrap"}

    for modulo, paquete in dependencias.items():
        try:
            __import__(modulo)
        except ImportError:
            if modulo in criticas:
                faltantes_criticas.append(paquete)
            else:
                faltantes_opcionales.append(paquete)

    # Mostrar advertencia por dependencias opcionales faltantes
    if faltantes_opcionales:
        print(f"[AVISO] Dependencias opcionales no instaladas: {', '.join(faltantes_opcionales)}")
        print("        La app funcionará sin ellas pero con estilo básico.")
        print(f"        Para instalarlas: pip install {' '.join(faltantes_opcionales)}\n")

    # Si faltan dependencias críticas, no se puede continuar
    if faltantes_criticas:
        print("=" * 60)
        print("ERROR: Faltan dependencias requeridas")
        print("=" * 60)
        print("\nLas siguientes librerías no están instaladas:")
        for pkg in faltantes_criticas:
            print(f"  - {pkg}")
        print("\nInstale todas las dependencias con:")
        print("  pip install -r requirements.txt")
        print("\nO instale solo las faltantes con:")
        print(f"  pip install {' '.join(faltantes_criticas)}")
        print("=" * 60)
        return False

    return True


# ---------------------------------------------------------------------------
# PASO 3: Arrancar la aplicación con captura de errores
# ---------------------------------------------------------------------------

def main():
    """
    Verifica dependencias y arranca la interfaz gráfica.

    Cualquier excepción no controlada se captura aquí y se muestra
    de forma legible antes de salir, evitando que la ventana se cierre
    sola sin dar información al usuario.
    """
    if not verificar_dependencias():
        input("\nPresione Enter para salir...")
        sys.exit(1)

    try:
        # Importar aquí (después de agregar src al path y verificar deps)
        from main_app import main as run_app
        run_app()

    except ImportError as e:
        # Error al importar algún módulo del proyecto
        print("\n" + "=" * 60)
        print("ERROR: No se pudo importar un módulo del proyecto")
        print("=" * 60)
        print(f"\nDetalle: {e}")
        print("\nVerifique que la carpeta 'src/' existe y contiene todos los archivos.")
        print("Archivos requeridos: config.py, database.py, excel_processor.py,")
        print("                     document_generator.py, main_app.py")
        input("\nPresione Enter para salir...")
        sys.exit(1)

    except Exception as e:
        # Cualquier otro error inesperado al iniciar
        import traceback
        print("\n" + "=" * 60)
        print("ERROR INESPERADO al iniciar la aplicación")
        print("=" * 60)
        print(f"\n{e}")
        print("\nDetalle completo del error:")
        traceback.print_exc()
        print("\nSi el problema persiste, revise el archivo:")
        print(f"  {BASE_DIR / 'historial' / 'actas.log'}")
        input("\nPresione Enter para salir...")
        sys.exit(1)


if __name__ == "__main__":
    main()
