"""
excel_processor.py - Procesador del archivo Excel de Cargas Atancha

Este módulo se encarga de leer el archivo Excel de entrada, limpiar los datos
y transformarlos en objetos Python estructurados que el resto de la aplicación
puede consumir fácilmente.

Clases principales:
  PuntoCuenta   : representa una fila del Excel (un punto de cuenta individual)
  ResumenSesion : agrupa todos los puntos de una misma sesión con sus totales
  ExcelProcessor: orquesta la carga, limpieza y agrupación de los datos

Mejoras implementadas:
  - Validación de columnas requeridas antes de procesar
  - Logging a archivo en lugar de print() para facilitar el soporte
"""

import pandas as pd
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from collections import defaultdict

from config import HISTORIAL_DIR

# ---------------------------------------------------------------------------
# LOGGER
# Comparte el mismo archivo de log que database.py (actas.log)
# ---------------------------------------------------------------------------
logger = logging.getLogger("excel_processor")

# Columnas que DEBEN existir en el Excel para que el procesamiento funcione.
# Si falta alguna, se informa al usuario con un mensaje claro.
COLUMNAS_REQUERIDAS = [
    "Id punto de cuenta",
    "Tipo de solicitud",
    "Estado",
    "Decisión",
    "Sesión",
    "Fecha",
]


# ---------------------------------------------------------------------------
# MODELOS DE DATOS
# ---------------------------------------------------------------------------

@dataclass
class PuntoCuenta:
    """
    Representa un punto de cuenta individual leído desde el Excel.

    Cada fila del archivo Excel corresponde a un PuntoCuenta.
    Los campos reflejan directamente las columnas del archivo.

    Atributos:
        id_punto           : identificador único del punto (columna 'Id punto de cuenta')
        tipo_solicitud     : tipo de trámite (Adjudicación, Declaratoria, Liberación)
        recomendacion      : recomendación principal (Si/No)
        recomendacion_extra: segunda columna de recomendación si existe
        recomendacion_extra2: tercera columna de recomendación si existe
        superficie         : superficie del predio en hectáreas
        solicitante        : nombre completo del solicitante
        cedula_rif         : documento de identidad del solicitante
        estado             : estado de Venezuela donde se ubica el predio
        decision           : resultado del directorio (Aprobado/Negado)
        observaciones      : notas adicionales del punto
        fecha              : fecha de la sesión en formato YYYY-MM-DD
        sesion             : número de sesión (ej. "ORD 1697-26")
    """
    id_punto: str
    tipo_solicitud: str
    recomendacion: str
    recomendacion_extra: str
    recomendacion_extra2: str
    superficie: str
    solicitante: str
    cedula_rif: str
    estado: str
    decision: str
    observaciones: str
    fecha: str
    sesion: str


@dataclass
class ResumenSesion:
    """
    Agrupa todos los puntos de una misma sesión y calcula sus totales.

    Se construye a partir de una lista de PuntoCuenta con el mismo
    número de sesión.

    Atributos:
        numero_sesion       : identificador de la sesión (ej. "ORD 1697-26")
        fecha               : fecha de la sesión
        total_aprobados     : cantidad de puntos con decisión 'Aprobado'
        total_negados       : cantidad de puntos con decisión 'Negado'
        por_estado          : dict anidado {estado: {tipo_aprobados: n, tipo_negados: n}}
        puntos              : lista completa de PuntoCuenta de la sesión
        tipos_solicitud     : lista de tipos únicos presentes (calculado en __post_init__)
        estados_involucrados: lista ordenada de estados presentes (calculado en __post_init__)
    """
    numero_sesion: str
    fecha: str
    total_aprobados: int
    total_negados: int
    por_estado: Dict[str, Dict[str, int]]
    puntos: List[PuntoCuenta]

    def __post_init__(self):
        """
        Calcula campos derivados después de la inicialización del dataclass.
        Se ejecuta automáticamente al crear la instancia.
        """
        # Lista de tipos de solicitud únicos presentes en la sesión
        self.tipos_solicitud = list(set(p.tipo_solicitud for p in self.puntos))

        # Lista de estados ordenada alfabéticamente
        self.estados_involucrados = sorted(list(self.por_estado.keys()))


# ---------------------------------------------------------------------------
# PROCESADOR PRINCIPAL
# ---------------------------------------------------------------------------

class ExcelProcessor:
    """
    Lee y procesa el archivo Excel de Cargas Atancha.

    Flujo de uso:
        processor = ExcelProcessor("ruta/al/archivo.xlsx")
        processor.cargar_archivo()          # carga el DataFrame
        sesiones = processor.procesar_datos()  # agrupa por sesión
        sesion = processor.obtener_resumen_sesion("ORD 1697-26")
    """

    def __init__(self, file_path: str):
        """
        Inicializa el procesador con la ruta al archivo Excel.

        Args:
            file_path: ruta absoluta o relativa al archivo .xlsx o .xls
        """
        self.file_path = Path(file_path)
        self.df = None                          # DataFrame de pandas con los datos crudos
        self.sesiones: Dict[str, ResumenSesion] = {}  # Resultado procesado

    # ------------------------------------------------------------------
    # CARGA DEL ARCHIVO
    # ------------------------------------------------------------------

    def cargar_archivo(self, sheet_name: str = None) -> bool:
        """
        Carga el archivo Excel en un DataFrame de pandas.

        Lógica de selección de hoja:
          1. Si se pasa sheet_name, usa esa hoja.
          2. Si existe una hoja llamada 'Base', la usa (convención del proyecto).
          3. Si no, usa la última hoja disponible.

        Validación aplicada:
          - Verifica que existan todas las columnas requeridas (COLUMNAS_REQUERIDAS)
          - Si falta alguna columna, retorna False con un mensaje claro

        Limpieza aplicada:
          - Elimina espacios en los nombres de columnas
          - Elimina filas donde 'Id punto de cuenta' es vacío o NaN

        Retorna True si la carga fue exitosa, False si ocurrió algún error.
        """
        try:
            excel_file  = pd.ExcelFile(self.file_path)
            sheet_names = excel_file.sheet_names

            if not sheet_name:
                if 'Base' in sheet_names:
                    sheet_name = 'Base'
                elif sheet_names:
                    sheet_name = sheet_names[-1]
                else:
                    sheet_name = 0

            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0)
            self.df.columns = [str(col).strip() for col in self.df.columns]

            # --- Validación de columnas requeridas ---
            # Detecta qué columnas faltan y las reporta antes de continuar
            faltantes = [
                col for col in COLUMNAS_REQUERIDAS
                if col not in self.df.columns
            ]
            if faltantes:
                msg = f"Columnas requeridas no encontradas: {', '.join(faltantes)}"
                logger.error(msg)
                # Guardar el mensaje para que main_app.py pueda mostrarlo al usuario
                self.error_columnas = msg
                return False

            self.error_columnas = None  # Sin errores de columnas

            # Eliminar filas sin identificador de punto de cuenta
            self.df = self.df.dropna(subset=['Id punto de cuenta'], how='all')
            self.df = self.df[
                self.df['Id punto de cuenta'].astype(str).str.strip() != ''
            ]
            self.df = self.df[
                self.df['Id punto de cuenta'].astype(str).str.lower() != 'nan'
            ]
            self.df['Id punto de cuenta'] = self.df['Id punto de cuenta'].astype(str)

            logger.info("Archivo cargado: %s | Hoja: %s | Filas: %d",
                        self.file_path.name, sheet_name, len(self.df))
            return True

        except Exception as e:
            logger.error("Error al cargar el archivo: %s", e)
            return False

    # ------------------------------------------------------------------
    # PROCESAMIENTO DE DATOS
    # ------------------------------------------------------------------

    def procesar_datos(self) -> Dict[str, ResumenSesion]:
        """
        Itera sobre el DataFrame y agrupa los puntos por número de sesión.

        Retorna un diccionario {numero_sesion: ResumenSesion}.
        Lanza ValueError si se llama antes de cargar_archivo().
        """
        if self.df is None:
            raise ValueError("Primero debe llamar a cargar_archivo()")

        # Agrupar filas por número de sesión usando un defaultdict de listas
        sesiones_temp: Dict[str, List[PuntoCuenta]] = defaultdict(list)

        for _, row in self.df.iterrows():
            punto = self._crear_punto_cuenta(row)
            # Solo agregar si el punto es válido y tiene número de sesión
            if punto and punto.sesion:
                sesiones_temp[punto.sesion].append(punto)

        # Construir objetos ResumenSesion a partir de los grupos
        self.sesiones = {}
        for sesion_id, puntos in sesiones_temp.items():
            self.sesiones[sesion_id] = self._crear_resumen_sesion(sesion_id, puntos)

        return self.sesiones

    def _crear_punto_cuenta(self, row: pd.Series) -> Optional[PuntoCuenta]:
        """
        Convierte una fila del DataFrame en un objeto PuntoCuenta.

        Maneja columnas adicionales de recomendación (el Excel puede tener
        varias columnas con ese nombre).

        Retorna None si la fila no tiene un ID de punto válido o si
        ocurre cualquier error de conversión.
        """
        try:
            id_punto = str(row.get('Id punto de cuenta', '')).strip()
            if not id_punto or id_punto == 'nan':
                return None

            tipo_solicitud = str(row.get('Tipo de solicitud', '')).strip()
            recomendacion  = str(row.get('Recomendación', '')).strip()

            recomendacion_extra  = ''
            recomendacion_extra2 = ''
            for col in self.df.columns:
                if 'Recomendación' in col and col != 'Recomendación':
                    if not recomendacion_extra:
                        recomendacion_extra = str(row.get(col, '')).strip()
                    else:
                        recomendacion_extra2 = str(row.get(col, '')).strip()

            return PuntoCuenta(
                id_punto=id_punto,
                tipo_solicitud=tipo_solicitud,
                recomendacion=recomendacion,
                recomendacion_extra=recomendacion_extra,
                recomendacion_extra2=recomendacion_extra2,
                superficie=str(row.get('Superficie', '')).strip(),
                solicitante=str(row.get('Solicitante', '')).strip(),
                cedula_rif=str(row.get('Cédula/Rif/Id-Red', '')).strip(),
                estado=str(row.get('Estado', '')).strip(),
                decision=str(row.get('Decisión', '')).strip(),
                observaciones=str(row.get('Observaciones', '')).strip(),
                fecha=str(row.get('Fecha', '')).strip(),
                sesion=str(row.get('Sesión', '')).strip(),
            )

        except Exception as e:
            logger.warning("Error al procesar fila: %s", e)
            return None

    def _crear_resumen_sesion(
        self, numero_sesion: str, puntos: List[PuntoCuenta]
    ) -> ResumenSesion:
        """
        Construye un ResumenSesion a partir de la lista de puntos de una sesión.

        Calcula:
          - total_aprobados / total_negados
          - por_estado: dict anidado con conteos por estado y tipo de solicitud
        """
        # Contar aprobados y negados buscando la palabra clave en la decisión
        total_aprobados = sum(1 for p in puntos if 'Aprobado' in p.decision)
        total_negados   = sum(1 for p in puntos if 'Negado'   in p.decision)

        # Tomar la fecha del primer punto (todos los de la sesión tienen la misma)
        fecha = puntos[0].fecha if puntos else ""

        # Construir el diccionario anidado por estado y tipo de solicitud
        # Estructura: {estado: {"TipoX_aprobados": n, "TipoX_negados": n, ...}}
        por_estado: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
        for punto in puntos:
            estado = punto.estado
            tipo   = punto.tipo_solicitud
            if 'Aprobado' in punto.decision:
                por_estado[estado][f"{tipo}_aprobados"] += 1
            else:
                por_estado[estado][f"{tipo}_negados"] += 1

        return ResumenSesion(
            numero_sesion=numero_sesion,
            fecha=fecha,
            total_aprobados=total_aprobados,
            total_negados=total_negados,
            por_estado=dict(por_estado),
            puntos=puntos,
        )

    # ------------------------------------------------------------------
    # CONSULTAS
    # ------------------------------------------------------------------

    def obtener_sesiones(self) -> List[str]:
        """Retorna la lista de números de sesión encontrados, ordenada."""
        return sorted(self.sesiones.keys())

    def obtener_resumen_sesion(self, numero_sesion: str) -> Optional[ResumenSesion]:
        """
        Retorna el ResumenSesion para el número de sesión indicado.
        Retorna None si la sesión no existe en los datos procesados.
        """
        return self.sesiones.get(numero_sesion)

    def generar_tabla_resumen(self, numero_sesion: str) -> "pd.DataFrame":
        """
        Genera un DataFrame con la tabla dinámica de aprobados para una sesión.

        Columnas: ESTADO | <tipo1> | <tipo2> | ... | TOTAL GENERAL
        La última fila es el total general por tipo.

        Retorna un DataFrame vacío si la sesión no existe.
        """
        sesion = self.sesiones.get(numero_sesion)
        if not sesion:
            return pd.DataFrame()

        # Obtener todos los tipos de solicitud únicos de la sesión
        tipos = sorted(set(p.tipo_solicitud for p in sesion.puntos))

        data = []
        for estado in sorted(sesion.por_estado.keys()):
            row = {'ESTADO': estado}
            total_estado = 0
            for tipo in tipos:
                aprobados = sesion.por_estado[estado].get(f"{tipo}_aprobados", 0)
                row[tipo] = aprobados
                total_estado += aprobados
            row['TOTAL GENERAL'] = total_estado
            data.append(row)

        # Fila de totales generales
        if data:
            total_row = {'ESTADO': 'Total general'}
            for tipo in tipos:
                total_row[tipo] = sum(r.get(tipo, 0) for r in data)
            total_row['TOTAL GENERAL'] = sum(r['TOTAL GENERAL'] for r in data)
            data.append(total_row)

        return pd.DataFrame(data)

    def generar_tabla_negados(self, numero_sesion: str) -> "pd.DataFrame":
        """
        Genera un DataFrame con el detalle de puntos negados para una sesión.

        Columnas: ESTADO | TIPO DE SOLICITUD | CANTIDAD
        Retorna un DataFrame vacío si no hay negados o la sesión no existe.
        """
        sesion = self.sesiones.get(numero_sesion)
        if not sesion or sesion.total_negados == 0:
            return pd.DataFrame()

        # Agrupar negados por estado y tipo
        negados: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
        for punto in sesion.puntos:
            if 'Negado' in punto.decision:
                negados[punto.estado][punto.tipo_solicitud] += 1

        data = [
            {'ESTADO': estado, 'TIPO DE SOLICITUD': tipo, 'CANTIDAD': cantidad}
            for estado in sorted(negados.keys())
            for tipo, cantidad in negados[estado].items()
        ]

        return pd.DataFrame(data)

    def obtener_estadisticas_globales(self) -> Dict:
        """
        Calcula estadísticas agregadas de todos los datos cargados en memoria.

        Útil para mostrar un resumen rápido del archivo Excel antes de guardar
        en la base de datos.

        Retorna dict con claves:
            total_sesiones, total_puntos, total_aprobados,
            total_negados, puntos_por_estado, puntos_por_tipo
        """
        total_puntos    = 0
        total_aprobados = 0
        total_negados   = 0
        puntos_por_estado: Dict[str, int] = defaultdict(int)
        puntos_por_tipo:   Dict[str, int] = defaultdict(int)

        for sesion in self.sesiones.values():
            total_puntos    += len(sesion.puntos)
            total_aprobados += sesion.total_aprobados
            total_negados   += sesion.total_negados
            for punto in sesion.puntos:
                puntos_por_estado[punto.estado]         += 1
                puntos_por_tipo[punto.tipo_solicitud]   += 1

        return {
            'total_sesiones':   len(self.sesiones),
            'total_puntos':     total_puntos,
            'total_aprobados':  total_aprobados,
            'total_negados':    total_negados,
            'puntos_por_estado': dict(puntos_por_estado),
            'puntos_por_tipo':   dict(puntos_por_tipo),
        }
