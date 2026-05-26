@echo off
chcp 65001 >nul
title Generador de Actas INTi

REM Cambiar al directorio donde esta este .bat para que las rutas relativas funcionen
cd /d "%~dp0"

REM Verificar que Python este instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH
    echo Instale Python desde https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante la instalacion
    pause
    exit /b 1
)

REM Verificar dependencias criticas e instalar si faltan
python -c "import pandas, docx, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: No se pudieron instalar las dependencias.
        echo Ejecute manualmente: pip install -r requirements.txt
        pause
        exit /b 1
    )
    echo Dependencias instaladas correctamente.
)

REM Iniciar la aplicacion
REM Se usa "start /wait python" para que la ventana no se cierre
REM mientras la app grafica esta abierta
start /wait python iniciar.py

REM Solo mostrar error si iniciar.py termino con codigo de error
if errorlevel 1 (
    echo.
    echo La aplicacion se cerro con errores.
    echo Revise el archivo historial\actas.log para mas detalles.
    pause
)
