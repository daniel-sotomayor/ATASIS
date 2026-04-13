# -*- coding: utf-8 -*-
"""
test_app.py - Suite de pruebas del Generador de Actas INTi

Verifica que todos los módulos importan correctamente y que las
operaciones principales funcionan sin errores:
  1. Importación de módulos
  2. Base de datos (guardar, recuperar, estadísticas, eliminar)
  3. Procesador de Excel (si hay archivo disponible)
  4. Generador de documentos Word

Uso:
    python test_app.py
"""

import sys
import os
from pathlib import Path

# Forzar UTF-8 en la salida estándar para que los símbolos PASS/FAIL
# se muestren correctamente en la consola de Windows (cp1252 no los soporta)
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

# Agregar la carpeta src al path para poder importar los módulos
src_dir = Path(__file__).parent / "src"
sys.path.insert(0, str(src_dir))

# Símbolos de resultado (compatibles con UTF-8)
OK   = "[OK]"
FAIL = "[FAIL]"
WARN = "[WARN]"


# ---------------------------------------------------------------------------
# PRUEBA 1: IMPORTACIÓN DE MÓDULOS
# ---------------------------------------------------------------------------

def test_imports() -> bool:
    """
    Verifica que todos los módulos del proyecto se pueden importar
    sin errores. Es la prueba más básica: si falla aquí, nada más funciona.
    """
    print("=" * 60)
    print("Prueba 1: Importacion de modulos")
    print("=" * 60)

    try:
        from config import APP_NAME, APP_VERSION
        print(f"{OK} config cargado: {APP_NAME} v{APP_VERSION}")
    except Exception as e:
        print(f"{FAIL} Error en config: {e}")
        return False

    try:
        from database import DatabaseManager, db_manager
        print(f"{OK} database cargado")
    except Exception as e:
        print(f"{FAIL} Error en database: {e}")
        return False

    try:
        from excel_processor import ExcelProcessor, PuntoCuenta, ResumenSesion
        print(f"{OK} excel_processor cargado")
    except Exception as e:
        print(f"{FAIL} Error en excel_processor: {e}")
        return False

    try:
        from document_generator import ActaDocumentGenerator
        print(f"{OK} document_generator cargado")
    except Exception as e:
        print(f"{FAIL} Error en document_generator: {e}")
        return False

    print(f"\n{OK} Todos los modulos importados correctamente")
    return True


# ---------------------------------------------------------------------------
# PRUEBA 2: BASE DE DATOS
# ---------------------------------------------------------------------------

def test_database() -> bool:
    """
    Prueba el ciclo completo de la base de datos:
      - Insertar un acta de prueba
      - Recuperarla por número de sesión
      - Consultar estadísticas globales (verifica que total_sesiones existe)
      - Guardar estadística de sesión por estado
      - Eliminar el registro de prueba
    """
    print("\n" + "=" * 60)
    print("Prueba 2: Base de datos")
    print("=" * 60)

    try:
        from database import db_manager

        # Insertar acta de prueba
        acta_id = db_manager.guardar_acta(
            numero_sesion="TEST-001",
            fecha_sesion="2026-03-13",
            total_aprobados=100,
            total_negados=5,
            estados_involucrados=["Anzoátegui", "Apure"],
            tipos_solicitud=["Adjudicación de Tierras"],
            archivo_generado="/test/acta.docx",
            observaciones="Acta de prueba"
        )
        print(f"{OK} Acta guardada con ID: {acta_id}")

        # Recuperar por sesión
        acta = db_manager.obtener_acta_por_sesion("TEST-001")
        if acta:
            print(f"{OK} Acta recuperada: {acta.numero_sesion}")
        else:
            print(f"{FAIL} No se pudo recuperar el acta")
            return False

        # Verificar que las estadísticas globales incluyen total_sesiones
        stats = db_manager.obtener_estadisticas_globales()
        assert "total_sesiones" in stats, "Falta clave total_sesiones"
        print(f"{OK} Estadisticas globales OK - {stats['total_actas']} actas, "
              f"{stats['total_sesiones']} sesiones unicas")

        # Guardar estadística de detalle por estado
        db_manager.guardar_estadistica_sesion(
            numero_sesion="TEST-001",
            estado="Anzoátegui",
            tipo_solicitud="Adjudicación de Tierras",
            cantidad_aprobados=80,
            cantidad_negados=3
        )
        print(f"{OK} Estadistica de sesion guardada")

        # Verificar que aparece en estadísticas por estado
        por_estado = db_manager.obtener_estadisticas_por_estado("TEST-001")
        assert len(por_estado) > 0, "No se encontraron estadísticas por estado"
        print(f"{OK} Estadisticas por estado OK - {len(por_estado)} estado(s)")

        # Eliminar el acta de prueba para no contaminar el historial
        db_manager.eliminar_acta(acta_id)
        print(f"{OK} Acta de prueba eliminada")

        return True

    except Exception as e:
        print(f"{FAIL} Error en base de datos: {e}")
        import traceback
        traceback.print_exc()
        return False


# ---------------------------------------------------------------------------
# PRUEBA 3: PROCESADOR DE EXCEL
# ---------------------------------------------------------------------------

def test_excel_processor() -> bool:
    """
    Prueba la carga y procesamiento del archivo Excel de Cargas Atancha.
    Si no se encuentra el archivo, la prueba se omite (no es un error crítico).
    """
    print("\n" + "=" * 60)
    print("Prueba 3: Procesador de Excel")
    print("=" * 60)

    try:
        from excel_processor import ExcelProcessor

        # Buscar el archivo Excel en ubicaciones conocidas
        candidatos = [
            Path(__file__).parent.parent / "upload" / "Cargas atancha 2026.xlsx",
            Path(__file__).parent / "Cargas atancha 2026.xlsx",
            Path("Cargas atancha 2026.xlsx"),
        ]

        archivo = next((f for f in candidatos if f.exists()), None)

        if not archivo:
            print(f"{WARN} No se encontro archivo Excel de prueba")
            print("  Buscado en:")
            for f in candidatos:
                print(f"    - {f}")
            print("  Coloque el archivo en una de esas rutas para probar esta funcion.")
            return True  # No es un error del código, solo falta el archivo

        print(f"Archivo encontrado: {archivo}")

        processor = ExcelProcessor(str(archivo))

        if not processor.cargar_archivo():
            print(f"{FAIL} Error al cargar el archivo")
            return False
        print(f"{OK} Archivo cargado correctamente")

        sesiones = processor.procesar_datos()
        print(f"{OK} Sesiones encontradas: {len(sesiones)}")

        # Mostrar las primeras 3 sesiones como muestra
        for sesion_id in list(sesiones.keys())[:3]:
            s = sesiones[sesion_id]
            print(f"  - {sesion_id}: {s.total_aprobados} aprobados, {s.total_negados} negados")

        return True

    except Exception as e:
        print(f"{FAIL} Error en ExcelProcessor: {e}")
        import traceback
        traceback.print_exc()
        return False


# ---------------------------------------------------------------------------
# PRUEBA 4: GENERADOR DE DOCUMENTOS
# ---------------------------------------------------------------------------

def test_document_generator() -> bool:
    """
    Prueba la generación de un documento Word completo con datos sintéticos.
    Verifica que el archivo se crea y tiene un tamaño razonable (> 0 bytes).

    También verifica la corrección de la tabla de negados: genera una sesión
    con negados y comprueba que el documento se crea sin errores.
    """
    print("\n" + "=" * 60)
    print("Prueba 4: Generador de Documentos")
    print("=" * 60)

    try:
        from document_generator import ActaDocumentGenerator, generar_nombre_archivo
        from excel_processor import ResumenSesion, PuntoCuenta

        # Datos de prueba con aprobados Y negados para probar ambas tablas
        puntos = [
            PuntoCuenta(
                id_punto="12345",
                tipo_solicitud="Adjudicación de Tierras",
                recomendacion="Si",
                recomendacion_extra="",
                recomendacion_extra2="",
                superficie="10 hectáreas",
                solicitante="Juan Perez",
                cedula_rif="12345678",
                estado="Anzoátegui",
                decision="Punto Cuenta Aprobado",
                observaciones="",
                fecha="2026-03-13",
                sesion="ORD 1697-26"
            ),
            PuntoCuenta(
                id_punto="12346",
                tipo_solicitud="Declaratoria de Garantía de Permanencia",
                recomendacion="Si",
                recomendacion_extra="",
                recomendacion_extra2="",
                superficie="5 hectáreas",
                solicitante="Maria Garcia",
                cedula_rif="87654321",
                estado="Apure",
                decision="Punto Cuenta Aprobado",
                observaciones="",
                fecha="2026-03-13",
                sesion="ORD 1697-26"
            ),
            # Punto negado para probar la tabla de negados (corrección aplicada)
            PuntoCuenta(
                id_punto="12347",
                tipo_solicitud="Adjudicación de Tierras",
                recomendacion="No",
                recomendacion_extra="",
                recomendacion_extra2="",
                superficie="3 hectáreas",
                solicitante="Pedro Lopez",
                cedula_rif="11223344",
                estado="Barinas",
                decision="Punto Cuenta Negado",
                observaciones="Documentación incompleta",
                fecha="2026-03-13",
                sesion="ORD 1697-26"
            ),
        ]

        sesion = ResumenSesion(
            numero_sesion="ORD 1697-26",
            fecha="2026-03-13",
            total_aprobados=2,
            total_negados=1,
            por_estado={
                "Anzoátegui": {"Adjudicación de Tierras_aprobados": 1},
                "Apure":      {"Declaratoria de Garantía de Permanencia_aprobados": 1},
                "Barinas":    {"Adjudicación de Tierras_negados": 1},
            },
            puntos=puntos
        )

        # Generar el documento
        generator = ActaDocumentGenerator()
        generator.generar_acta(sesion)

        # Guardar en output_actas/
        output_dir = Path(__file__).parent / "output_actas"
        output_dir.mkdir(exist_ok=True)

        filename = generar_nombre_archivo("ORD 1697-26", "2026-03-13")
        filepath = output_dir / filename
        generator.guardar(str(filepath))

        print(f"{OK} Documento generado: {filepath.name}")

        if filepath.exists() and filepath.stat().st_size > 0:
            print(f"{OK} Archivo verificado: {filepath.stat().st_size} bytes")
        else:
            print(f"{FAIL} El archivo no existe o está vacío")
            return False

        return True

    except Exception as e:
        print(f"{FAIL} Error en DocumentGenerator: {e}")
        import traceback
        traceback.print_exc()
        return False


# ---------------------------------------------------------------------------
# PUNTO DE ENTRADA
# ---------------------------------------------------------------------------

def main() -> int:
    """
    Ejecuta todas las pruebas en orden y muestra un resumen final.
    Retorna 0 si todas pasaron, 1 si alguna falló.
    """
    print("\n" + "=" * 60)
    print("GENERADOR DE ACTAS INTi - SUITE DE PRUEBAS")
    print("=" * 60)

    resultados = [
        ("Importaciones",       test_imports()),
        ("Base de datos",       test_database()),
        ("Excel Processor",     test_excel_processor()),
        ("Document Generator",  test_document_generator()),
    ]

    print("\n" + "=" * 60)
    print("RESUMEN")
    print("=" * 60)

    for nombre, resultado in resultados:
        estado = f"{OK} PASS" if resultado else f"{FAIL} FAIL"
        print(f"{estado}: {nombre}")

    pasadas = sum(1 for _, r in resultados if r)
    total   = len(resultados)
    print(f"\nTotal: {pasadas}/{total} pruebas pasadas")

    if pasadas == total:
        print(f"\n{OK} TODAS LAS PRUEBAS PASARON")
        return 0
    else:
        print(f"\n{FAIL} ALGUNAS PRUEBAS FALLARON")
        return 1


if __name__ == "__main__":
    exit_code = main()
    input("\nPresione Enter para salir...")
    sys.exit(exit_code)
