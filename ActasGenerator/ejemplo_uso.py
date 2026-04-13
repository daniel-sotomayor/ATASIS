#!/usr/bin/env python3
"""
Ejemplo de uso programático del Generador de Actas

Este script muestra cómo usar los módulos de la aplicación
sin la interfaz gráfica.
"""

import sys
from pathlib import Path

# Agregar src al path
src_dir = Path(__file__).parent / "src"
sys.path.insert(0, str(src_dir))

from excel_processor import ExcelProcessor
from document_generator import ActaDocumentGenerator, generar_nombre_archivo
from database import db_manager
from config import OUTPUT_DIR


def ejemplo_generar_actas():
    """Ejemplo: Generar actas desde un archivo Excel"""
    
    print("=" * 60)
    print("Ejemplo: Generar Actas desde Excel")
    print("=" * 60)
    print()
    
    # 1. Cargar archivo Excel
    archivo_excel = input("Ruta del archivo Excel (o Enter para usar ejemplo): ").strip()
    
    if not archivo_excel:
        # Usar archivo de ejemplo
        archivo_excel = "/mnt/okcomputer/upload/Cargas atancha 2026.xlsx"
        print(f"Usando archivo de ejemplo: {archivo_excel}")
    
    # 2. Crear procesador y cargar datos
    processor = ExcelProcessor(archivo_excel)
    
    if not processor.cargar_archivo():
        print("✗ Error al cargar el archivo")
        return
    
    print("✓ Archivo cargado correctamente")
    
    # 3. Procesar datos
    sesiones = processor.procesar_datos()
    print(f"✓ Sesiones encontradas: {len(sesiones)}")
    
    # 4. Mostrar sesiones disponibles
    print("\nSesiones disponibles:")
    for i, sesion_id in enumerate(sorted(sesiones.keys()), 1):
        sesion = sesiones[sesion_id]
        print(f"  {i}. {sesion_id} - {sesion.total_aprobados} aprobados, {sesion.total_negados} negados")
    
    # 5. Seleccionar sesiones a generar
    seleccion = input("\nIngrese los números de sesión a generar (separados por coma, o 'todas'): ").strip()
    
    if seleccion.lower() == 'todas':
        sesiones_a_generar = list(sesiones.keys())
    else:
        try:
            indices = [int(x.strip()) - 1 for x in seleccion.split(',')]
            sesiones_list = sorted(sesiones.keys())
            sesiones_a_generar = [sesiones_list[i] for i in indices if 0 <= i < len(sesiones_list)]
        except:
            print("✗ Selección inválida")
            return
    
    # 6. Generar actas
    print(f"\nGenerando {len(sesiones_a_generar)} actas...")
    
    for sesion_id in sesiones_a_generar:
        try:
            sesion = sesiones[sesion_id]
            
            # Generar documento
            generator = ActaDocumentGenerator()
            doc = generator.generar_acta(sesion)
            
            # Guardar
            filename = generar_nombre_archivo(sesion_id, sesion.fecha)
            filepath = OUTPUT_DIR / filename
            generator.guardar(str(filepath))
            
            # Guardar en historial
            db_manager.guardar_acta(
                numero_sesion=sesion_id,
                fecha_sesion=sesion.fecha,
                total_aprobados=sesion.total_aprobados,
                total_negados=sesion.total_negados,
                estados_involucrados=sesion.estados_involucrados,
                tipos_solicitud=sesion.tipos_solicitud,
                archivo_generado=str(filepath)
            )
            
            print(f"  ✓ {sesion_id} -> {filename}")
            
        except Exception as e:
            print(f"  ✗ {sesion_id} -> Error: {e}")
    
    print(f"\n✓ Generación completada. Archivos guardados en: {OUTPUT_DIR}")


def ejemplo_ver_estadisticas():
    """Ejemplo: Ver estadísticas del historial"""
    
    print("=" * 60)
    print("Ejemplo: Ver Estadísticas")
    print("=" * 60)
    print()
    
    stats = db_manager.obtener_estadisticas_globales()
    
    print("Estadísticas Globales:")
    print(f"  Total de actas generadas: {stats['total_actas']}")
    print(f"  Total de puntos aprobados: {stats['total_aprobados']}")
    print(f"  Total de puntos negados: {stats['total_negados']}")
    
    print("\nActas por mes:")
    for mes, cantidad in stats['actas_por_mes']:
        print(f"  {mes}: {cantidad} actas")
    
    print("\nEstadísticas por estado:")
    estados_stats = db_manager.obtener_estadisticas_por_estado()
    for estado, aprobados, negados in estados_stats[:10]:
        print(f"  {estado}: {aprobados} aprobados, {negados} negados")


def ejemplo_buscar_historial():
    """Ejemplo: Buscar en el historial"""
    
    print("=" * 60)
    print("Ejemplo: Buscar en Historial")
    print("=" * 60)
    print()
    
    termino = input("Término de búsqueda (sesión, fecha): ").strip()
    
    resultados = db_manager.buscar_actas(termino)
    
    print(f"\nResultados encontrados: {len(resultados)}")
    
    for acta in resultados:
        print(f"\n  Sesión: {acta.numero_sesion}")
        print(f"  Fecha: {acta.fecha_sesion}")
        print(f"  Generado: {acta.fecha_generacion}")
        print(f"  Aprobados: {acta.total_aprobados}, Negados: {acta.total_negados}")
        print(f"  Archivo: {acta.archivo_generado}")


def main():
    """Menú principal de ejemplos"""
    
    print("\n" + "=" * 60)
    print("GENERADOR DE ACTAS INTi - Ejemplos de Uso")
    print("=" * 60)
    print()
    print("1. Generar actas desde Excel")
    print("2. Ver estadísticas")
    print("3. Buscar en historial")
    print("4. Salir")
    print()
    
    opcion = input("Seleccione una opción: ").strip()
    
    if opcion == "1":
        ejemplo_generar_actas()
    elif opcion == "2":
        ejemplo_ver_estadisticas()
    elif opcion == "3":
        ejemplo_buscar_historial()
    elif opcion == "4":
        print("Saliendo...")
    else:
        print("Opción inválida")


if __name__ == "__main__":
    main()
