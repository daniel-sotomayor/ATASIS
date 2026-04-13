"""
database.py - Gestión de la base de datos SQLite del historial de actas

Este módulo centraliza toda la interacción con la base de datos.
Contiene:
  - ActaRecord : dataclass que representa una fila de la tabla 'actas'
  - DatabaseManager : clase con todos los métodos CRUD y de consulta
  - db_manager : instancia global lista para usar desde cualquier módulo

Tablas:
  actas               → registro principal de cada acta generada
  estadisticas_sesion → detalle por estado/tipo para el dashboard
  configuracion       → pares clave/valor para configuración persistente

Mejoras implementadas:
  - Backup automático de la BD antes de cada generación de actas
  - Logging a archivo en lugar de print() para facilitar el soporte
"""

import sqlite3
import json
import shutil
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

from config import DB_PATH, HISTORIAL_DIR

# ---------------------------------------------------------------------------
# CONFIGURACIÓN DEL LOGGER
# Escribe en historial/actas.log además de mostrar en consola.
# Nivel INFO registra operaciones normales; ERROR registra fallos.
# ---------------------------------------------------------------------------

# Carpeta de backups dentro de historial/
BACKUP_DIR = HISTORIAL_DIR / "backup"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

# Archivo de log
LOG_PATH = HISTORIAL_DIR / "actas.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        # Escribe en archivo (modo append, UTF-8 para caracteres especiales)
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        # También muestra en consola
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger("database")


# ---------------------------------------------------------------------------
# MODELO DE DATOS
# ---------------------------------------------------------------------------

@dataclass
class ActaRecord:
    """
    Representa un registro completo de acta tal como se almacena en la BD.

    Campos:
        id                  : clave primaria autoincremental
        numero_sesion       : identificador de la sesión (ej. "ORD 1697-26")
        fecha_sesion        : fecha de la sesión en formato YYYY-MM-DD
        fecha_generacion    : timestamp de cuando se generó el documento
        total_aprobados     : cantidad de puntos aprobados en la sesión
        total_negados       : cantidad de puntos negados en la sesión
        estados_involucrados: JSON string con lista de estados
        tipos_solicitud     : JSON string con lista de tipos de solicitud
        archivo_generado    : ruta absoluta al archivo .docx generado
        observaciones       : notas adicionales opcionales
    """
    id: int
    numero_sesion: str
    fecha_sesion: str
    fecha_generacion: str
    total_aprobados: int
    total_negados: int
    estados_involucrados: str
    tipos_solicitud: str
    archivo_generado: str
    observaciones: str


# ---------------------------------------------------------------------------
# GESTOR DE BASE DE DATOS
# ---------------------------------------------------------------------------

class DatabaseManager:
    """
    Gestiona todas las operaciones sobre la base de datos SQLite.

    Uso típico:
        from database import db_manager
        db_manager.guardar_acta(...)
        registros = db_manager.obtener_historial()
    """

    def __init__(self, db_path: Path = DB_PATH):
        """
        Inicializa el gestor apuntando al archivo de base de datos indicado.
        Crea las tablas si no existen todavía.
        """
        self.db_path = db_path
        self._init_database()
        logger.info("DatabaseManager iniciado. BD: %s", self.db_path)

    # ------------------------------------------------------------------
    # INICIALIZACIÓN
    # ------------------------------------------------------------------

    def _init_database(self):
        """
        Crea las tres tablas necesarias si aún no existen.
        Se ejecuta automáticamente al instanciar la clase.

        Tablas creadas:
          - actas               : registro principal por sesión
          - estadisticas_sesion : detalle por estado y tipo de solicitud
          - configuracion       : almacén de pares clave/valor
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # Tabla principal: una fila por acta generada
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS actas (
                    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_sesion        TEXT    NOT NULL,
                    fecha_sesion         TEXT    NOT NULL,
                    fecha_generacion     TEXT    NOT NULL,
                    total_aprobados      INTEGER DEFAULT 0,
                    total_negados        INTEGER DEFAULT 0,
                    estados_involucrados TEXT,
                    tipos_solicitud      TEXT,
                    archivo_generado     TEXT,
                    observaciones        TEXT
                )
            ''')

            # Tabla de detalle: una fila por combinación estado+tipo dentro de una sesión
            # Alimenta el panel "Estadísticas por Estado" del dashboard
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS estadisticas_sesion (
                    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_sesion      TEXT    NOT NULL,
                    estado             TEXT    NOT NULL,
                    tipo_solicitud     TEXT    NOT NULL,
                    cantidad_aprobados INTEGER DEFAULT 0,
                    cantidad_negados   INTEGER DEFAULT 0,
                    fecha_registro     TEXT    NOT NULL
                )
            ''')

            # Tabla de configuración: pares clave/valor para uso futuro
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS configuracion (
                    clave TEXT PRIMARY KEY,
                    valor TEXT
                )
            ''')

            conn.commit()

    # ------------------------------------------------------------------
    # ESCRITURA
    # ------------------------------------------------------------------

    def hacer_backup(self):
        """
        Crea una copia de seguridad del archivo de base de datos en
        historial/backup/ con un timestamp en el nombre.

        Formato del nombre: actas_historial_YYYYMMDD_HHMMSS.db

        Se llama automáticamente desde guardar_acta() antes de cada
        inserción, garantizando que siempre hay un backup previo a
        cualquier modificación de datos.
        """
        if not self.db_path.exists():
            return  # No hay nada que respaldar todavía

        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"actas_historial_{timestamp}.db"
        backup_path = BACKUP_DIR / backup_name

        try:
            shutil.copy2(self.db_path, backup_path)
            logger.info("Backup creado: %s", backup_path.name)
        except Exception as e:
            # Un fallo de backup no debe interrumpir la operación principal
            logger.error("No se pudo crear el backup: %s", e)

    def guardar_acta(
        self,
        numero_sesion: str,
        fecha_sesion: str,
        total_aprobados: int,
        total_negados: int,
        estados_involucrados: List[str],
        tipos_solicitud: List[str],
        archivo_generado: str,
        observaciones: str = ""
    ) -> int:
        """
        Inserta un nuevo registro en la tabla 'actas'.

        Antes de insertar, crea un backup automático de la BD para
        proteger los datos existentes.

        Las listas estados_involucrados y tipos_solicitud se serializan
        como JSON para almacenarse en columnas TEXT.

        Retorna el ID autoincremental del registro insertado.
        """
        # Backup automático antes de modificar la BD
        self.hacer_backup()
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Timestamp del momento exacto de generación
            fecha_generacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cursor.execute('''
                INSERT INTO actas
                (numero_sesion, fecha_sesion, fecha_generacion,
                 total_aprobados, total_negados, estados_involucrados,
                 tipos_solicitud, archivo_generado, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                numero_sesion,
                fecha_sesion,
                fecha_generacion,
                total_aprobados,
                total_negados,
                json.dumps(estados_involucrados, ensure_ascii=False),
                json.dumps(tipos_solicitud, ensure_ascii=False),
                archivo_generado,
                observaciones,
            ))

            conn.commit()
            logger.info("Acta guardada - sesion: %s, aprobados: %d, negados: %d",
                        numero_sesion, total_aprobados, total_negados)
            return cursor.lastrowid

    def guardar_estadistica_sesion(
        self,
        numero_sesion: str,
        estado: str,
        tipo_solicitud: str,
        cantidad_aprobados: int,
        cantidad_negados: int
    ):
        """
        Inserta una fila de detalle en 'estadisticas_sesion'.

        Debe llamarse una vez por cada combinación (estado, tipo_solicitud)
        encontrada en la sesión. Estos datos alimentan el panel
        "Estadísticas por Estado" del dashboard.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            fecha_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cursor.execute('''
                INSERT INTO estadisticas_sesion
                (numero_sesion, estado, tipo_solicitud,
                 cantidad_aprobados, cantidad_negados, fecha_registro)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                numero_sesion,
                estado,
                tipo_solicitud,
                cantidad_aprobados,
                cantidad_negados,
                fecha_registro,
            ))

            conn.commit()

    # ------------------------------------------------------------------
    # LECTURA
    # ------------------------------------------------------------------

    def obtener_historial(self, limite: int = 100) -> List[ActaRecord]:
        """
        Retorna las últimas `limite` actas generadas, ordenadas de más
        reciente a más antigua.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, numero_sesion, fecha_sesion, fecha_generacion,
                       total_aprobados, total_negados, estados_involucrados,
                       tipos_solicitud, archivo_generado, observaciones
                FROM actas
                ORDER BY fecha_generacion DESC
                LIMIT ?
            ''', (limite,))

            rows = cursor.fetchall()
            return [ActaRecord(*row) for row in rows]

    def obtener_acta_por_sesion(self, numero_sesion: str) -> Optional[ActaRecord]:
        """
        Busca el acta más reciente con el número de sesión indicado.
        Retorna None si no existe ningún registro.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, numero_sesion, fecha_sesion, fecha_generacion,
                       total_aprobados, total_negados, estados_involucrados,
                       tipos_solicitud, archivo_generado, observaciones
                FROM actas
                WHERE numero_sesion = ?
                ORDER BY fecha_generacion DESC
                LIMIT 1
            ''', (numero_sesion,))

            row = cursor.fetchone()
            return ActaRecord(*row) if row else None

    def obtener_estadisticas_globales(self) -> Dict:
        """
        Calcula y retorna un diccionario con los totales acumulados de
        todas las actas registradas en la base de datos.

        Claves del diccionario retornado:
            total_actas     : número de actas generadas
            total_aprobados : suma de todos los puntos aprobados
            total_negados   : suma de todos los puntos negados
            total_sesiones  : número de sesiones únicas registradas
            actas_por_mes   : lista de tuplas (mes, cantidad) últimos 12 meses
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # Cantidad total de actas en el historial
            cursor.execute('SELECT COUNT(*) FROM actas')
            total_actas = cursor.fetchone()[0]

            # Suma acumulada de aprobados y negados
            cursor.execute('''
                SELECT SUM(total_aprobados), SUM(total_negados)
                FROM actas
            ''')
            row = cursor.fetchone()
            total_aprobados = row[0] or 0
            total_negados   = row[1] or 0

            # Número de sesiones únicas (una sesión puede tener varias actas
            # si se regeneró, pero se cuenta como una sola sesión)
            cursor.execute('SELECT COUNT(DISTINCT numero_sesion) FROM actas')
            total_sesiones = cursor.fetchone()[0]

            # Distribución de actas por mes (para gráficos futuros)
            cursor.execute('''
                SELECT strftime('%Y-%m', fecha_generacion) AS mes,
                       COUNT(*) AS cantidad
                FROM actas
                GROUP BY mes
                ORDER BY mes DESC
                LIMIT 12
            ''')
            actas_por_mes = cursor.fetchall()

            return {
                "total_actas":     total_actas,
                "total_aprobados": total_aprobados,
                "total_negados":   total_negados,
                "total_sesiones":  total_sesiones,
                "actas_por_mes":   actas_por_mes,
            }

    def obtener_estadisticas_por_estado(
        self, numero_sesion: str = None
    ) -> List[Tuple]:
        """
        Retorna estadísticas agrupadas por estado desde la tabla
        'estadisticas_sesion'.

        Si se pasa numero_sesion, filtra solo esa sesión.
        Si no se pasa, agrega todos los registros históricos.

        Retorna lista de tuplas: (estado, total_aprobados, total_negados)
        ordenadas de mayor a menor por aprobados.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            if numero_sesion:
                # Estadísticas de una sesión específica
                cursor.execute('''
                    SELECT estado,
                           SUM(cantidad_aprobados) AS total_aprobados,
                           SUM(cantidad_negados)   AS total_negados
                    FROM estadisticas_sesion
                    WHERE numero_sesion = ?
                    GROUP BY estado
                    ORDER BY total_aprobados DESC
                ''', (numero_sesion,))
            else:
                # Estadísticas globales de todas las sesiones
                cursor.execute('''
                    SELECT estado,
                           SUM(cantidad_aprobados) AS total_aprobados,
                           SUM(cantidad_negados)   AS total_negados
                    FROM estadisticas_sesion
                    GROUP BY estado
                    ORDER BY total_aprobados DESC
                ''')

            return cursor.fetchall()

    def buscar_actas(self, termino: str) -> List[ActaRecord]:
        """
        Busca actas cuyo número de sesión, fecha o observaciones
        contengan el término indicado (búsqueda parcial, sin distinción
        de mayúsculas/minúsculas gracias al LIKE de SQLite).

        Retorna lista de ActaRecord ordenada de más reciente a más antigua.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            patron = f'%{termino}%'
            cursor.execute('''
                SELECT id, numero_sesion, fecha_sesion, fecha_generacion,
                       total_aprobados, total_negados, estados_involucrados,
                       tipos_solicitud, archivo_generado, observaciones
                FROM actas
                WHERE numero_sesion  LIKE ?
                   OR fecha_sesion   LIKE ?
                   OR observaciones  LIKE ?
                ORDER BY fecha_generacion DESC
            ''', (patron, patron, patron))

            rows = cursor.fetchall()
            return [ActaRecord(*row) for row in rows]

    # ------------------------------------------------------------------
    # ELIMINACIÓN
    # ------------------------------------------------------------------

    def eliminar_acta(self, acta_id: int) -> bool:
        """
        Elimina el registro con el ID indicado de la tabla 'actas'.
        Retorna True si se eliminó al menos una fila, False en caso contrario
        o si ocurrió un error.
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM actas WHERE id = ?', (acta_id,))
                conn.commit()
                eliminado = cursor.rowcount > 0
                if eliminado:
                    logger.info("Acta eliminada - ID: %d", acta_id)
                return eliminado
        except Exception as e:
            logger.error("Error al eliminar acta ID %d: %s", acta_id, e)
            return False


# ---------------------------------------------------------------------------
# INSTANCIA GLOBAL
# Se importa directamente desde otros módulos:
#   from database import db_manager
# ---------------------------------------------------------------------------

db_manager = DatabaseManager()
