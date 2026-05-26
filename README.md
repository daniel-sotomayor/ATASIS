# 📋 Resumen del Proyecto - Generador de Actas INTi

## Descripción

Aplicación de escritorio para Windows 10/11 que automatiza la generación de actas del Instituto Nacional de Tierras (INTi) a partir del archivo Excel de Cargas Atancha.

## Estructura del Proyecto

```
ActasGenerator/
│
├── 📁 src/                          # Código fuente Python
│   ├── __init__.py                  # Inicialización del paquete
│   ├── config.py                    # Configuración global
│   ├── database.py                  # Base de datos SQLite
│   ├── excel_processor.py           # Procesador de Excel
│   ├── document_generator.py        # Generador de Word
│   └── main_app.py                  # Aplicación principal (GUI)
│
├── 📁 assets/                       # Recursos (iconos, imágenes)
├── 📁 output_actas/                 # Actas generadas (output)
├── 📁 historial/                    # Base de datos del historial
└── 📁 reports/                      # Reportes exportados
│
├── 📄 Iniciar Generador de Actas.bat    # Ejecutor para Windows
├── 📄 iniciar.py                        # Script de inicio
├── 📄 instalar_dependencias.py          # Instalador de dependencias
├── 📄 crear_ejecutable.py               # Creador de .exe
├── 📄 ejemplo_uso.py                    # Ejemplos de uso
├── 📄 test_app.py                       # Pruebas unitarias
│
├── 📄 requirements.txt              # Dependencias Python
├── 📄 config.json                   # Configuración JSON
│
├── 📄 README.md                     # Documentación general
├── 📄 MANUAL_DE_USO.md              # Manual del usuario
└── 📄 RESUMEN_PROYECTO.md           # Este archivo
```

## Funcionalidades

### ✅ Generación de Actas
- Procesa automáticamente el archivo Excel de Cargas Atancha
- Detecta todas las sesiones disponibles
- Genera actas en formato Word (.docx) con formato oficial
- Tablas resumen por estado y tipo de solicitud
- Manejo de aprobados y negados

### ✅ Dashboard
- Estadísticas globales en tiempo real
- Total de actas generadas
- Total de puntos aprobados y negados
- Estadísticas por estado
- Actas recientes

### ✅ Historial
- Base de datos SQLite persistente
- Búsqueda de actas por sesión o fecha
- Menú contextual con opciones
- Consulta de actas antiguas

### ✅ Interfaz Gráfica
- Diseño moderno con ttkbootstrap
- Tres pestañas: Dashboard, Generar Actas, Historial
- Barra de progreso durante la generación
- Mensajes informativos y de error

## Tecnologías Utilizadas

| Tecnología | Uso |
|------------|-----|
| Python 3.8+ | Lenguaje de programación |
| tkinter | Interfaz gráfica nativa |
| ttkbootstrap | Estilos modernos para tkinter |
| pandas | Procesamiento de datos Excel |
| openpyxl | Lectura de archivos Excel |
| python-docx | Generación de documentos Word |
| SQLite | Base de datos para historial |

## Cómo Usar

### Método 1: Ejecutor Batch (Recomendado)

1. Haga doble clic en `Iniciar Generador de Actas.bat`
2. La primera vez se instalarán las dependencias automáticamente
3. Use la interfaz gráfica para generar actas

### Método 2: Python

```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
python iniciar.py
```

### Método 3: Ejemplo Programático

```bash
python ejemplo_uso.py
```

## Flujo de Trabajo

```
┌─────────────────┐
│  Cargar Excel   │
│ (Cargas Atancha)│
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Procesar Datos  │
│ (Detectar       │
│  sesiones)      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Seleccionar     │
│ Sesiones        │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Generar Actas   │
│ (Word .docx)    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Guardar en      │
│ Historial       │
└─────────────────┘
```

## Archivos de Salida

### Actas Generadas
- **Ubicación:** `output_actas/`
- **Formato:** `ACTA_ORD_XXXX-XX_YYYYMMDD.docx`
- **Ejemplo:** `ACTA_ORD 1697-26_20260411.docx`

### Base de Datos
- **Ubicación:** `historial/actas_historial.db`
- **Formato:** SQLite
- **Tablas:** actas, estadisticas_sesion, configuracion

## Configuración

### Miembros del Directorio
Editar `src/config.py`:
```python
MIEMBROS_DIRECTORIO = {
    "presidente": "Nombre del Presidente",
    "directores_principales": ["Director 1", "Director 2", ...],
    "directores_suplentes": ["Suplente 1", "Suplente 2", ...],
    "secretaria": "Nombre de la Secretaria"
}
```

### Parámetros de Sesión
Editar `config.json`:
```json
{
  "formato_acta": {
    "hora_inicio": "9:00 a.m.",
    "hora_fin": "11:30 a.m.",
    "reunion_numero": "98"
  }
}
```

## Pruebas

Ejecutar pruebas unitarias:
```bash
python test_app.py
```

Las pruebas verifican:
- Importación de módulos
- Funcionamiento de la base de datos
- Procesamiento de Excel
- Generación de documentos

## Crear Ejecutable (.exe)

Para distribuir la aplicación sin Python:

```bash
python crear_ejecutable.py
```

El ejecutable se creará en `dist/Generador de Actas INTi.exe`

## Requisitos del Sistema

- **SO:** Windows 10 o 11 (64 bits)
- **RAM:** 4 GB mínimo
- **Disco:** 500 MB disponibles
- **Python:** 3.8+ (solo para ejecución desde código fuente)

## Dependencias

```
pandas>=2.0.0
openpyxl>=3.1.0
python-docx>=0.8.11
ttkbootstrap>=1.10.0
Pillow>=10.0.0
```

## Estadísticas de Prueba

Con el archivo `Cargas atancha 2026.xlsx`:
- **Total de sesiones:** 16
- **Total de puntos:** 6,631
- **Tiempo de procesamiento:** ~2 segundos
- **Tiempo de generación por acta:** ~1 segundo

## Próximas Mejoras (Roadmap)

- [ ] Exportar historial a Excel/CSV
- [ ] Generar reportes PDF
- [ ] Vista previa de actas antes de generar
- [ ] Edición de actas generadas
- [ ] Soporte para múltiples idiomas
- [ ] Tema oscuro/claro
- [ ] Atajos de teclado
- [ ] Notificaciones de escritorio

## Licencia

Uso exclusivo para el Instituto Nacional de Tierras (INTi).

## Versión

**Versión Actual:** 1.0.0  
**Fecha de Lanzamiento:** Abril 2026  
**Compatibilidad:** Windows 10, Windows 11

---

**Desarrollado para:** INTi - Instituto Nacional de Tierras  
**Tecnología:** Python + tkinter + SQLite
