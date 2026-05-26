# 📋 Generador de Actas INTi

Aplicación de escritorio para Windows 10/11 que automatiza la generación de actas del Instituto Nacional de Tierras (INTi) a partir del archivo Excel de Cargas Atancha.

## ✨ Características

- **📊 Dashboard** con estadísticas en tiempo real
- **📄 Generación automática** de actas en formato Word
- **📁 Procesamiento** del archivo Excel de Cargas Atancha
- **💾 Historial** de actas generadas con base de datos SQLite
- **📈 Reportes** y estadísticas detalladas
- **🎨 Interfaz moderna** e intuitiva
- **🔍 Búsqueda** en el historial de actas

## 🚀 Instalación Rápida

### Requisitos

- Windows 10 o 11 (64 bits)
- Python 3.8 o superior (si se ejecuta desde código fuente)

### Instalación

1. **Descomprima** el archivo ZIP en la ubicación deseada

2. **Ejecute** el archivo `Iniciar Generador de Actas.bat`

3. **La primera vez** se instalarán automáticamente las dependencias necesarias

## 📖 Uso

### Generar Actas

1. Haga clic en **"Cargar Excel"** y seleccione el archivo `Cargas atancha 2026.xlsx`

2. Haga clic en **"Procesar"** para analizar el archivo

3. Seleccione las **sesiones** que desea generar (haciendo clic en la casilla ✓)

4. Haga clic en **"Generar Actas Seleccionadas"**

5. Las actas se guardarán en la carpeta `output_actas/`

### Ver Estadísticas

- Vaya a la pestaña **"Dashboard"** para ver estadísticas globales
- Gráficos de actas por estado
- Totales de aprobados y negados
- Actas recientes generadas

### Consultar Historial

- Vaya a la pestaña **"Historial"**
- Use la barra de búsqueda para encontrar actas específicas
- Haga clic derecho en un registro para más opciones

## 📁 Estructura de Carpetas

```
ActasGenerator/
├── src/                    # Código fuente
│   ├── config.py          # Configuración
│   ├── database.py        # Base de datos SQLite
│   ├── document_generator.py  # Generador de Word
│   ├── excel_processor.py     # Procesador de Excel
│   └── main_app.py        # Aplicación principal
├── assets/                # Recursos (imágenes, etc.)
├── output_actas/          # Actas generadas (output)
├── historial/             # Base de datos del historial
├── reports/               # Reportes exportados
├── iniciar.py             # Script de inicio
├── instalar_dependencias.py  # Instalador de dependencias
├── requirements.txt       # Dependencias Python
└── README.md             # Este archivo
```

## 🔧 Instalación Manual (Desarrolladores)

Si desea ejecutar desde el código fuente:

```bash
# 1. Clonar o descargar el proyecto
cd ActasGenerator

# 2. Crear entorno virtual (opcional pero recomendado)
python -m venv venv
venv\Scripts\activate  # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar
python iniciar.py
```

## 📋 Formato del Excel de Entrada

El archivo Excel debe contener las siguientes columnas:

- `N°` - Número de fila
- `Id punto de cuenta` - Identificador del punto
- `Tipo de solicitud` - Tipo (Adjudicación, Declaratoria, Liberación)
- `Recomendación` - Si/No
- `Superficie` - Superficie del predio
- `Solicitante` - Nombre del solicitante
- `Cédula/Rif/Id-Red` - Identificación
- `Estado` - Estado de Venezuela
- `Decisión` - Punto Cuenta Aprobado/Negado
- `Observaciones` - Observaciones
- `Fecha` - Fecha de la sesión
- `Sesión` - Número de sesión (ej: ORD 1697-26)

## 🛠️ Solución de Problemas

### "Python no está instalado"

Descargue e instale Python desde: https://www.python.org/downloads/

**IMPORTANTE:** Durante la instalación, marque la opción **"Add Python to PATH"**

### "No se pudieron instalar las dependencias"

Ejecute manualmente:
```bash
pip install pandas openpyxl python-docx ttkbootstrap
```

### "Error al procesar el archivo Excel"

Verifique que:
- El archivo no esté abierto en otro programa
- Las columnas tengan los nombres correctos
- El archivo no esté protegido con contraseña

## 📝 Notas

- Las actas generadas siguen el formato oficial del INTi
- El historial se guarda en una base de datos SQLite local
- Las actas se generan en formato `.docx` (Microsoft Word)
- La aplicación no requiere conexión a Internet

## 📞 Soporte

Para reportar problemas o solicitar nuevas funciones:

1. Verifique que tiene la última versión
2. Revista la sección de solución de problemas
3. Contacte al administrador del sistema

## 📄 Licencia

Este software es propiedad del Instituto Nacional de Tierras (INTi).
Uso exclusivo para fines institucionales.

---

**Versión:** 1.0.0  
**Desarrollado para:** INTi - Instituto Nacional de Tierras  
**Compatible con:** Windows 10, Windows 11
