"""
Final test to verify header appears on all pages.
"""
import sys
import os
from pathlib import Path

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Mock the config values needed for testing
class MockConfig:
    OUTPUT_DIR = Path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "output_actas"))
    ASSETS_DIR = Path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets"))
    MIEMBROS_DIRECTORIO = {
        'presidente': 'David José Hernández Giménez',
        'directores_principales': [
            'Faiez Kassen Castillo',
            'Giomar Gisela Blanco Espinoza',
            'Lusmialit Lliliana Permodo Torrealba',
            'David Arturo Gómez Rodríguez'
        ],
        'directores_suplentes': [
            'Virginia Coromoto Delgado Oropeza',
            'Francisco Ramon Ojeda Barrera',
            'Alfredo José Mora Álvarez',
            'Gisela Ninoska Faraco De Cuellar'
        ],
        'secretaria': 'Viviana Del Carmen Rivas Tovar'
    }
    REUNION_NUMERO = "98"
    HORA_INICIO_DEFAULT = "9:00 a.m."
    HORA_FIN_DEFAULT = "11:30 a.m."

# Mock the config module
sys.modules['config'] = MockConfig

# Mock the excel_processor module
class MockResumenSesion:
    def __init__(self, numero_sesion, fecha, puntos, total_aprobados, total_negados):
        self.numero_sesion = numero_sesion
        self.fecha = fecha
        self.puntos = puntos
        self.total_aprobados = total_aprobados
        self.total_negados = total_negados

sys.modules['excel_processor'] = type(sys)('excel_processor')
sys.modules['excel_processor'].ResumenSesion = MockResumenSesion

from src.document_generator import ActaDocumentGenerator

# Create test data
class TestPunto:
    def __init__(self, tipo_solicitud, estado, decision):
        self.tipo_solicitud = tipo_solicitud
        self.estado = estado
        self.decision = decision

# Create multiple test points to ensure multiple pages
test_puntos = []
for i in range(50):  # Create enough data to span multiple pages
    estado = ["Lara", "Carabobo", "Aragua", "Miranda"][i % 4]
    tipo = ["Adjudicación", "Declaratoria", "Liberación"][i % 3]
    decision = "Punto Cuenta Aprobado" if i % 5 != 0 else "Punto Cuenta Negado"
    test_puntos.append(TestPunto(tipo, estado, decision))

test_sesion = MockResumenSesion(
    numero_sesion="ORD 1697-26",
    fecha="2026-03-13",
    puntos=test_puntos,
    total_aprobados=40,
    total_negados=10
)

# Generate the acta
print("Generando acta de prueba con múltiples páginas...")
generator = ActaDocumentGenerator()
document = generator.generar_acta(test_sesion)

# Save the document
output_path = "test_header_all_pages.docx"
generator.guardar(output_path)
print(f"Acta generada: {output_path}")
print("\nVerificación:")
print("1. Abra el documento en Microsoft Word")
print("2. Verifique que el encabezado con 'ACTA DE DIRECTORIO' aparece en TODAS las páginas")
print("3. El encabezado debe contener:")
print("   - Logo INTi (si existe en assets/logo.jpeg)")
print("   - Título 'ACTA DE DIRECTORIO'")
print("   - Acta Nro.: ORD 1697-26")
print("   - Sesión Nro.: ORD 1697-26")
print("   - Pág. : (sin número)")
print("   - Fecha 13/03/2026")
print("\nEl encabezado ahora es parte de la sección de encabezado de Word,")
print("lo que significa que aparecerá automáticamente en todas las páginas.")

# Clean up test file
import os
if os.path.exists("test_header_output.docx"):
    os.remove("test_header_output.docx")