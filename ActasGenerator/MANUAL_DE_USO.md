# 📘 Manual de Usuario - Generador de Actas INTi

## Índice

1. [Introducción](#introducción)
2. [Instalación](#instalación)
3. [Interfaz Principal](#interfaz-principal)
4. [Generar Actas](#generar-actas)
5. [Dashboard](#dashboard)
6. [Historial](#historial)
7. [Solución de Problemas](#solución-de-problemas)

---

## Introducción

El **Generador de Actas INTi** es una aplicación de escritorio que automatiza la creación de actas del Instituto Nacional de Tierras a partir del archivo Excel de Cargas Atancha.

### Características Principales

- ✅ Procesamiento automático de archivos Excel
- ✅ Generación de actas en formato Word (.docx)
- ✅ Dashboard con estadísticas en tiempo real
- ✅ Historial de actas generadas
- ✅ Búsqueda y filtrado de actas
- ✅ Interfaz intuitiva y moderna

---

## Instalación

### Requisitos del Sistema

- **Sistema Operativo:** Windows 10 o Windows 11 (64 bits)
- **Memoria RAM:** 4 GB mínimo (8 GB recomendado)
- **Espacio en disco:** 500 MB disponibles
- **Software:** Python 3.8+ (solo si ejecuta desde código fuente)

### Instalación Rápida

1. **Descargue** el archivo ZIP del Generador de Actas

2. **Descomprima** el archivo en la ubicación deseada (ej: `C:\ActasGenerator\`)

3. **Ejecute** el archivo `Iniciar Generador de Actas.bat`

4. **La primera vez** se instalarán automáticamente las dependencias necesarias

### Instalación Manual (Desarrolladores)

Si desea ejecutar desde el código fuente:

```bash
# 1. Navegue a la carpeta del proyecto
cd ActasGenerator

# 2. Instale las dependencias
pip install -r requirements.txt

# 3. Ejecute la aplicación
python iniciar.py
```

---

## Interfaz Principal

La aplicación tiene tres pestañas principales:

### 1. Dashboard
Muestra estadísticas globales y gráficos de las actas generadas.

### 2. Generar Actas
Permite cargar el archivo Excel y generar las actas.

### 3. Historial
Muestra el historial de todas las actas generadas.

---

## Generar Actas

### Paso 1: Cargar el Archivo Excel

1. Vaya a la pestaña **"Generar Actas"**

2. Haga clic en el botón **"📁 Cargar Excel"**

3. Seleccione el archivo `Cargas atancha 2026.xlsx`

4. El sistema mostrará la ruta del archivo cargado

### Paso 2: Procesar el Archivo

1. Haga clic en el botón **"🔄 Procesar"**

2. El sistema analizará el archivo y mostrará las sesiones encontradas

3. Verá una tabla con:
   - Número de sesión
   - Fecha
   - Cantidad de aprobados
   - Cantidad de negados
   - Estados involucrados

### Paso 3: Seleccionar Sesiones

1. Haga clic en la casilla **☐** de las sesiones que desea generar
   - La casilla cambiará a **☑** cuando esté seleccionada

2. Puede usar los botones:
   - **"✅ Seleccionar Todas"** - Selecciona todas las sesiones
   - **"⬜ Deseleccionar Todas"** - Deselecciona todas las sesiones

### Paso 4: Generar las Actas

1. Haga clic en **"📄 Generar Actas Seleccionadas"**

2. Espere mientras se generan los documentos
   - Verá una barra de progreso
   - El sistema mostrará el estado de cada acta

3. Al finalizar, verá un mensaje con:
   - Cantidad de actas generadas exitosamente
   - Errores (si los hubo)
   - Ubicación de los archivos

### Paso 5: Ubicación de los Archivos

Las actas generadas se guardan en la carpeta:
```
ActasGenerator/output_actas/
```

Los archivos tienen el formato:
```
ACTA_ORD_XXXX-XX_YYYYMMDD.docx
```

---

## Dashboard

El Dashboard muestra estadísticas en tiempo real de todas las actas generadas.

### Tarjetas de Estadísticas

En la parte superior verá 4 tarjetas:

| Tarjeta | Descripción |
|---------|-------------|
| 📄 Total Actas | Cantidad total de actas generadas |
| ✅ Aprobados | Total de puntos de cuenta aprobados |
| ❌ Negados | Total de puntos de cuenta negados |
| 📊 Sesiones | Cantidad de sesiones diferentes |

### Estadísticas por Estado

En el panel izquierdo verá una tabla con:
- **Estado:** Nombre del estado
- **Aprobados:** Puntos aprobados en ese estado
- **Negados:** Puntos negados en ese estado
- **Total:** Suma total

### Actas Recientes

En el panel derecho verá las últimas 10 actas generadas con:
- Número de sesión
- Fecha
- Cantidad de aprobados y negados

### Actualizar Estadísticas

Haga clic en **"🔄 Actualizar Estadísticas"** para refrescar los datos.

---

## Historial

La pestaña de Historial permite consultar y buscar actas generadas anteriormente.

### Ver el Historial

1. Vaya a la pestaña **"Historial"**

2. Verá una tabla con todas las actas generadas:
   - ID
   - Sesión
   - Fecha de sesión
   - Fecha de generación
   - Aprobados
   - Negados
   - Nombre del archivo

### Buscar en el Historial

1. Escriba un término de búsqueda en el campo **"Buscar"**
   - Puede buscar por número de sesión
   - Puede buscar por fecha

2. Haga clic en **"🔍 Buscar"**

3. La tabla mostrará solo los resultados que coincidan

### Menú Contextual

Haga clic derecho sobre cualquier acta para ver opciones:

- **Abrir archivo:** Abre el documento Word
- **Ver detalles:** Muestra información detallada
- **Eliminar registro:** Elimina el registro del historial

### Actualizar Historial

Haga clic en **"🔄 Actualizar"** para refrescar la lista.

---

## Solución de Problemas

### Error: "Python no está instalado"

**Causa:** Python no está instalado o no está en el PATH del sistema.

**Solución:**
1. Descargue Python desde: https://www.python.org/downloads/
2. Durante la instalación, **marque la casilla** "Add Python to PATH"
3. Reinicie la aplicación

### Error: "No se pudieron instalar las dependencias"

**Causa:** Problema de conexión o permisos.

**Solución:**
1. Ejecute como administrador
2. Instale manualmente:
   ```bash
   pip install pandas openpyxl python-docx ttkbootstrap
   ```

### Error: "No se pudo cargar el archivo Excel"

**Causa:** El archivo está abierto en otro programa o tiene formato incorrecto.

**Solución:**
1. Cierre el archivo Excel si está abierto
2. Verifique que el archivo no esté protegido con contraseña
3. Asegúrese de que las columnas tengan los nombres correctos:
   - Id punto de cuenta
   - Tipo de solicitud
   - Estado
   - Decisión
   - Sesión
   - Fecha

### Las actas no se generan

**Causa:** Problema con los permisos de escritura.

**Solución:**
1. Verifique que tenga permisos de escritura en la carpeta `output_actas/`
2. Ejecute la aplicación como administrador
3. Cambie la ubicación de la carpeta a una donde tenga permisos

### La aplicación se cierra inesperadamente

**Causa:** Error en el procesamiento de datos.

**Solución:**
1. Verifique que el archivo Excel tenga el formato correcto
2. Revise que no haya filas vacías en medio de los datos
3. Contacte al administrador del sistema

---

## Preguntas Frecuentes

### ¿Puedo modificar el formato de las actas?

Sí, puede editar el archivo `src/config.py` para cambiar:
- Nombres de los miembros del directorio
- Horarios de las sesiones
- Otros parámetros del formato

### ¿Las actas generadas son editables?

Sí, los archivos se generan en formato `.docx` (Microsoft Word) y pueden editarse.

### ¿Puedo generar actas de sesiones anteriores?

Sí, siempre que el archivo Excel contenga los datos de esas sesiones.

### ¿Se pierde el historial al cerrar la aplicación?

No, el historial se guarda en una base de datos SQLite persistente.

### ¿Puedo exportar el historial?

Actualmente no hay función de exportación, pero puede consultar la base de datos en:
```
historial/actas_historial.db
```

---

## Contacto y Soporte

Para reportar problemas o solicitar nuevas funciones:

1. Verifique que tiene la última versión de la aplicación
2. Revise la sección de Solución de Problemas
3. Contacte al administrador del sistema INTi

---

**Versión del Manual:** 1.0  
**Última actualización:** Abril 2026
