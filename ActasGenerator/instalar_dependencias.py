#!/usr/bin/env python3
"""
Script para instalar las dependencias del Generador de Actas
"""

import subprocess
import sys
import os

def instalar_dependencias():
    """Instala las dependencias desde requirements.txt"""
    print("=" * 60)
    print("Instalador de Dependencias - Generador de Actas INTi")
    print("=" * 60)
    print()
    
    # Verificar si pip está disponible
    try:
        subprocess.run([sys.executable, "-m", "pip", "--version"], 
                      check=True, capture_output=True)
    except subprocess.CalledProcessError:
        print("ERROR: No se encontró pip. Por favor instale Python correctamente.")
        return False
    
    # Instalar dependencias
    print("Instalando dependencias...")
    print("Esto puede tomar unos minutos...")
    print()
    
    requirements_file = os.path.join(os.path.dirname(__file__), "requirements.txt")
    
    if not os.path.exists(requirements_file):
        print(f"ERROR: No se encontró el archivo {requirements_file}")
        return False
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", requirements_file],
            capture_output=False,
            text=True
        )
        
        if result.returncode == 0:
            print()
            print("=" * 60)
            print("¡Instalación completada exitosamente!")
            print("=" * 60)
            print()
            print("Ahora puede ejecutar la aplicación con:")
            print("  python iniciar.py")
            print()
            print("O haciendo doble clic en 'Iniciar Generador de Actas.bat'")
            print("=" * 60)
            return True
        else:
            print()
            print("ERROR: La instalación falló")
            return False
            
    except Exception as e:
        print(f"ERROR: {e}")
        return False

if __name__ == "__main__":
    exito = instalar_dependencias()
    input("\nPresione Enter para salir...")
    sys.exit(0 if exito else 1)
