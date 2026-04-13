"""
document_generator.py - Generador de documentos Word para las Actas del INTi

Construye el documento .docx replicando exactamente el formato oficial:

  1. Encabezado: tabla 2 columnas
       - Izquierda: logo INTi + título 'ACTA DE DIRECTORIO'
       - Derecha  : 4 filas (Acta Nro. | Sesión Nro. | Pág. | Fecha)
  2. Texto narrativo de apertura
  3. Agenda del directorio
  4. Decisiones del directorio
  5. Tabla de aprobados por estado y tipo
  6. Tabla de negados (solo si hay negados)
  7. Lectura y aprobación
  8. Firmas: tabla con presidente centrado arriba, directores principales
     en una fila, directores suplentes en otra fila, secretaria debajo

Clase principal : ActaDocumentGenerator
Función auxiliar: generar_nombre_archivo
"""

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime

from config import OUTPUT_DIR, ASSETS_DIR, MIEMBROS_DIRECTORIO
from config import REUNION_NUMERO, HORA_INICIO_DEFAULT, HORA_FIN_DEFAULT
from excel_processor import ResumenSesion


class ActaDocumentGenerator:
    """
    Genera documentos Word (.docx) con el formato oficial de actas del INTi.

    Uso:
        gen = ActaDocumentGenerator()
        gen.generar_acta(sesion)
        gen.guardar("/ruta/acta.docx")
    """

    def __init__(self):
        self.document = None

    # ------------------------------------------------------------------
    # UTILIDADES DE CELDAS
    # ------------------------------------------------------------------

    def _set_cell_shading(self, cell, color: str):
        """Aplica color de fondo a una celda. color = hex sin '#' (ej. 'D9E1F2')."""
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color)
        cell._tc.get_or_add_tcPr().append(shd)

    def _format_cell_text(self, cell, text: str, bold=False,
                          font_size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                          italic=False):
        """Escribe texto en una celda con formato básico."""
        cell.text = text
        p = cell.paragraphs[0]
        p.alignment = alignment
        for run in p.runs:
            run.font.name   = 'Arial'
            run.font.size   = Pt(font_size)
            run.font.bold   = bold
            run.font.italic = italic

    def _add_run(self, paragraph, text: str, bold=False, italic=True,
                 font_size=11):
        """Agrega un run con formato a un párrafo existente."""
        run = paragraph.add_run(text)
        run.font.name   = 'Arial'
        run.font.size   = Pt(font_size)
        run.font.bold   = bold
        run.font.italic = italic
        return run

    # ------------------------------------------------------------------
    # MÉTODO PRINCIPAL
    # ------------------------------------------------------------------

    def generar_acta(self, sesion: ResumenSesion,
                     fecha_acta: str = None,
                     hora_inicio: str = None,
                     hora_fin: str = None) -> Document:
        """
        Genera el documento Word completo para una sesión.

        Args:
            sesion     : datos de la sesión
            fecha_acta : YYYY-MM-DD (usa sesion.fecha si es None)
            hora_inicio: usa HORA_INICIO_DEFAULT de config.json si es None
            hora_fin   : usa HORA_FIN_DEFAULT de config.json si es None
        """
        self.document = Document()

        hora_inicio = hora_inicio or HORA_INICIO_DEFAULT
        hora_fin    = hora_fin    or HORA_FIN_DEFAULT

        # Márgenes de página
        for section in self.document.sections:
            section.top_margin    = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        self._agregar_encabezado_oficial(sesion, fecha_acta)
        self._agregar_texto_apertura(sesion, fecha_acta, hora_inicio)
        self._agregar_agenda()
        self._agregar_decisiones(sesion)
        self._agregar_tabla_aprobados(sesion)

        if sesion.total_negados > 0:
            self._agregar_tabla_negados(sesion)

        self._agregar_lectura_aprobacion(sesion, hora_fin)
        self._agregar_firmas()

        return self.document

    # ------------------------------------------------------------------
    # SECCIONES DEL DOCUMENTO
    # ------------------------------------------------------------------

    def _agregar_encabezado_oficial(self, sesion: ResumenSesion,
                                    fecha_acta: str = None):
        """
        Encabezado oficial: tabla de 2 columnas.

        Columna izquierda (ancha):
          - Logo del INTi
          - Título 'ACTA DE DIRECTORIO' en negrita cursiva grande

        Columna derecha (angosta), 4 filas separadas con borde:
          - Acta Nro.: ORD XXXX-XX
          - Sesión Nro.: ORD XXXX-XX
          - Pág.: (número de página)
          - Fecha DD/MM/YYYY
        """
        if fecha_acta is None:
            fecha_acta = sesion.fecha if sesion.fecha else datetime.now().strftime("%Y-%m-%d")
        try:
            fecha_display = datetime.strptime(fecha_acta, "%Y-%m-%d").strftime("%d/%m/%Y")
            # Formato corto para el encabezado: "13/3/2026"
            fecha_corta = datetime.strptime(fecha_acta, "%Y-%m-%d").strftime("%-d/%m/%Y")
        except (ValueError, AttributeError):
            fecha_display = fecha_acta
            fecha_corta   = fecha_acta

        # Tabla principal del encabezado: 2 columnas
        table = self.document.add_table(rows=4, cols=2)
        table.style = 'Table Grid'

        # Anchos: izquierda ~72%, derecha ~28%
        table.columns[0].width = Inches(4.8)
        table.columns[1].width = Inches(1.85)

        # --- Columna izquierda: fusionar las 4 filas en una sola celda ---
        celda_izq = table.cell(0, 0)
        for r in range(1, 4):
            celda_izq = celda_izq.merge(table.cell(r, 0))

        # Insertar logo del INTi si existe
        logo_path = ASSETS_DIR / "logo.jpeg"
        p_logo = celda_izq.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if logo_path.exists():
            p_logo.add_run().add_picture(str(logo_path), width=Cm(2.8))

        # Título ACTA DE DIRECTORIO
        p_titulo = celda_izq.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_titulo = p_titulo.add_run("  ACTA DE DIRECTORIO")
        run_titulo.font.name = 'Arial'
        run_titulo.font.size = Pt(22)
        run_titulo.bold      = True
        run_titulo.italic    = True

        # --- Columna derecha: 4 filas con los datos de la sesión ---
        # Fila 0: Acta Nro.
        c = table.cell(0, 1)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self._add_run(p, f"Acta Nro.: {sesion.numero_sesion}", font_size=11)

        # Fila 1: Sesión Nro.
        c = table.cell(1, 1)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self._add_run(p, f"Sesión Nro.: {sesion.numero_sesion}", font_size=11)

        # Fila 2: Pág.
        c = table.cell(2, 1)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self._add_run(p, "Pág. : 1", font_size=11)

        # Fila 3: Fecha
        c = table.cell(3, 1)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self._add_run(p, f"Fecha {fecha_display}", font_size=11)

        # Espacio después del encabezado
        self.document.add_paragraph()

    def _agregar_texto_apertura(self, sesion: ResumenSesion,
                                fecha_acta: str = None,
                                hora_inicio: str = "9:00 a.m."):
        """
        Párrafo narrativo de apertura con los asistentes,
        número de reunión y número de sesión ordinaria.
        """
        if fecha_acta is None:
            fecha_acta = sesion.fecha if sesion.fecha else datetime.now().strftime("%Y-%m-%d")
        try:
            fecha_formateada = datetime.strptime(fecha_acta, "%Y-%m-%d").strftime("%d de %B del %Y")
        except ValueError:
            fecha_formateada = fecha_acta

        texto = (
            f"En el día de hoy {fecha_formateada}, siendo las {hora_inicio}, "
            f"reunidos vía telemática, el Presidente del Instituto Nacional de Tierras, "
            f"Ciudadano {MIEMBROS_DIRECTORIO['presidente']}, quien preside esta sesión "
            f"como lo establece el Artículo 121 de la Ley de Tierras y Desarrollo Agrario, "
            f"{MIEMBROS_DIRECTORIO['directores_principales'][0]}, miembro principal del "
            f"Directorio del Instituto Nacional de Tierras, vía telemática los miembros "
            f"principales del Directorio del Instituto Nacional de Tierras ciudadanos: "
            f"{', '.join(MIEMBROS_DIRECTORIO['directores_principales'][1:])}. "
            f"Acompañados por la funcionaria del Instituto Nacional de Tierras en su "
            f"condición de Coordinadora (E) de Secretaria del Directorio "
            f"{MIEMBROS_DIRECTORIO['secretaria']}, para dar continuidad a la reunión "
            f"N° {REUNION_NUMERO}, en la sesión ordinaria número ORD {sesion.numero_sesion} "
            f"del Directorio del Instituto Nacional de Tierras."
        )

        p = self.document.add_paragraph()
        run = p.add_run(texto)
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.italic    = True
        p.alignment   = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _agregar_agenda(self):
        """Sección Agenda del Directorio con el punto estándar."""
        p = self.document.add_paragraph()
        run = p.add_run("Agenda del Directorio")
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.bold      = True
        run.italic    = True

        p2 = self.document.add_paragraph("1. ", style='List Number')
        run2 = p2.add_run("Revisión de los puntos de cuenta del sistema ATANCHA-OMAKÓN.")
        run2.font.name = 'Arial'
        run2.font.size = Pt(11)
        run2.italic    = True

    def _agregar_decisiones(self, sesion: ResumenSesion):
        """
        Sección Decisiones del Directorio.
        Los números de aprobados y negados se muestran en negrita.
        """
        p = self.document.add_paragraph()
        run = p.add_run("Decisiones del Directorio:")
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.bold      = True
        run.italic    = True

        p2 = self.document.add_paragraph("1. ", style='List Number')
        self._add_run(p2, "Se Aprobaron ")
        self._add_run(p2, str(sesion.total_aprobados), bold=True)
        self._add_run(p2, " y negó ")
        self._add_run(p2, str(sesion.total_negados), bold=True)
        self._add_run(p2, " Puntos de Cuenta generados por el sistema "
                         "Atancha-Omakón según consta en el siguiente cuadro:")

    def _agregar_tabla_aprobados(self, sesion: ResumenSesion):
        """
        Tabla de aprobados por estado y tipo de solicitud.
        Estructura: encabezado fusionado | cols de tipos | TOTAL GENERAL
        """
        tipos = sorted(set(
            p.tipo_solicitud for p in sesion.puntos if 'Aprobado' in p.decision
        ))
        if not tipos:
            return

        try:
            fecha_tabla = datetime.strptime(sesion.fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            fecha_tabla = sesion.fecha

        num_cols = len(tipos) + 2
        table = self.document.add_table(rows=1, cols=num_cols)
        table.style     = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Fila 0: encabezado fusionado
        hcells = table.rows[0].cells
        hcells[0].merge(hcells[-1])
        self._format_cell_text(
            hcells[0],
            f"SESIÓN {sesion.numero_sesion.replace('-', '_')} FECHA {fecha_tabla} APROBADOS",
            bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
        self._set_cell_shading(hcells[0], "D9E1F2")

        # Fila 1: encabezados de columnas
        row = table.add_row()
        self._format_cell_text(row.cells[0], "ESTADOS", bold=True, font_size=9)
        self._set_cell_shading(row.cells[0], "D9E1F2")
        for i, tipo in enumerate(tipos, 1):
            self._format_cell_text(row.cells[i], tipo.upper(), bold=True, font_size=8)
            self._set_cell_shading(row.cells[i], "D9E1F2")
        self._format_cell_text(row.cells[-1], "TOTAL GENERAL", bold=True, font_size=9)
        self._set_cell_shading(row.cells[-1], "D9E1F2")

        # Filas de datos
        estados  = sorted(set(p.estado for p in sesion.puntos if 'Aprobado' in p.decision))
        totales  = {tipo: 0 for tipo in tipos}

        for estado in estados:
            row = table.add_row()
            self._format_cell_text(row.cells[0], estado, font_size=9)
            total_estado = 0
            for i, tipo in enumerate(tipos, 1):
                cant = sum(1 for p in sesion.puntos
                           if p.estado == estado and p.tipo_solicitud == tipo
                           and 'Aprobado' in p.decision)
                self._format_cell_text(
                    row.cells[i], str(cant) if cant > 0 else "",
                    font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
                )
                total_estado  += cant
                totales[tipo] += cant
            self._format_cell_text(
                row.cells[-1], str(total_estado),
                bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
            )

        # Fila de totales
        row = table.add_row()
        self._format_cell_text(row.cells[0], "Total general", bold=True, font_size=9)
        self._set_cell_shading(row.cells[0], "D9E1F2")
        for i, tipo in enumerate(tipos, 1):
            self._format_cell_text(
                row.cells[i], str(totales[tipo]),
                bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
            )
            self._set_cell_shading(row.cells[i], "D9E1F2")
        self._format_cell_text(
            row.cells[-1], str(sum(totales.values())),
            bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
        self._set_cell_shading(row.cells[-1], "D9E1F2")

        self.document.add_paragraph()

    def _agregar_tabla_negados(self, sesion: ResumenSesion):
        """
        Tabla de negados por estado y tipo de solicitud.
        Misma estructura que la tabla de aprobados.
        """
        tipos = sorted(set(
            p.tipo_solicitud for p in sesion.puntos if 'Negado' in p.decision
        ))
        if not tipos:
            return

        try:
            fecha_tabla = datetime.strptime(sesion.fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            fecha_tabla = sesion.fecha

        num_cols = len(tipos) + 2
        table = self.document.add_table(rows=1, cols=num_cols)
        table.style     = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hcells = table.rows[0].cells
        hcells[0].merge(hcells[-1])
        self._format_cell_text(
            hcells[0],
            f"SESIÓN {sesion.numero_sesion.replace('-', '_')} FECHA {fecha_tabla} NEGADOS",
            bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
        self._set_cell_shading(hcells[0], "D9E1F2")

        row = table.add_row()
        self._format_cell_text(row.cells[0], "ESTADOS", bold=True, font_size=9)
        self._set_cell_shading(row.cells[0], "D9E1F2")
        for i, tipo in enumerate(tipos, 1):
            self._format_cell_text(row.cells[i], tipo.upper(), bold=True, font_size=8)
            self._set_cell_shading(row.cells[i], "D9E1F2")
        self._format_cell_text(row.cells[-1], "TOTAL GENERAL", bold=True, font_size=9)
        self._set_cell_shading(row.cells[-1], "D9E1F2")

        estados = sorted(set(p.estado for p in sesion.puntos if 'Negado' in p.decision))
        totales = {tipo: 0 for tipo in tipos}

        for estado in estados:
            row = table.add_row()
            self._format_cell_text(row.cells[0], estado, font_size=9)
            total_estado = 0
            for i, tipo in enumerate(tipos, 1):
                cant = sum(1 for p in sesion.puntos
                           if p.estado == estado and p.tipo_solicitud == tipo
                           and 'Negado' in p.decision)
                self._format_cell_text(
                    row.cells[i], str(cant) if cant > 0 else "",
                    font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
                )
                total_estado  += cant
                totales[tipo] += cant
            self._format_cell_text(
                row.cells[-1], str(total_estado),
                bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
            )

        row = table.add_row()
        self._format_cell_text(row.cells[0], "Total general", bold=True, font_size=9)
        self._set_cell_shading(row.cells[0], "D9E1F2")
        for i, tipo in enumerate(tipos, 1):
            self._format_cell_text(
                row.cells[i], str(totales[tipo]),
                bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
            )
            self._set_cell_shading(row.cells[i], "D9E1F2")
        self._format_cell_text(
            row.cells[-1], str(sum(totales.values())),
            bold=True, font_size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
        self._set_cell_shading(row.cells[-1], "D9E1F2")

        self.document.add_paragraph()

    def _agregar_lectura_aprobacion(self, sesion: ResumenSesion, hora_fin: str):
        """Sección de lectura y aprobación del acta con hora de cierre."""
        p = self.document.add_paragraph()
        run = p.add_run("Lectura y aprobación del Acta.")
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.bold      = True
        run.italic    = True

        p2 = self.document.add_paragraph()
        run2 = p2.add_run(
            f"Se procedió a dar lectura del Acta Nº {sesion.numero_sesion} de la "
            f"Sesión Ordinaria de Directorio para su aprobación. Se decidió aprobar "
            f"y firmar el Acta. Siendo las {hora_fin}, se dio por concluida la presente Sesión."
        )
        run2.font.name = 'Arial'
        run2.font.size = Pt(11)
        run2.italic    = True

    def _agregar_firmas(self):
        """
        Bloque de firmas replicando exactamente el formato oficial:

          Tabla de 4 columnas:
            Fila 0: espacio en blanco (línea de firma del presidente)
            Fila 1: nombre del Presidente centrado (columnas fusionadas)
            Fila 2: etiqueta 'Presidente' centrada
            Fila 3: espacio separador
            Fila 4: 4 directores principales (uno por columna)
            Fila 5: etiqueta 'Directores Principales' fusionada
            Fila 6: espacio separador
            Fila 7: 4 directores suplentes (uno por columna)
            Fila 8: etiqueta 'Directores Suplentes' fusionada

          Debajo de la tabla:
            Línea de firma + nombre + cargo de la Secretaria del Directorio
        """
        self.document.add_paragraph()

        # Tabla de 4 columnas (una por director)
        table = self.document.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Fila 0: espacio para la firma del presidente
        # La fusión debe hacerse ANTES de escribir en la celda
        c0 = table.cell(0, 0)
        c0.merge(table.cell(0, 3))
        c0.paragraphs[0].add_run("\n\n")

        # Fila 1: nombre del Presidente centrado (fusionar antes de escribir)
        row = table.add_row()
        merged = row.cells[0].merge(row.cells[3])
        self._format_cell_text(
            merged, MIEMBROS_DIRECTORIO['presidente'],
            bold=True, font_size=10, italic=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER
        )

        # Fila 2: etiqueta Presidente
        row = table.add_row()
        merged = row.cells[0].merge(row.cells[3])
        self._format_cell_text(
            merged, "Presidente",
            bold=True, font_size=10, italic=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER
        )

        # Fila 3: espacio separador
        row = table.add_row()
        row.cells[0].merge(row.cells[3]).paragraphs[0].add_run(" ")

        # Fila 4: 4 directores principales (uno por columna)
        row = table.add_row()
        for i, nombre in enumerate(MIEMBROS_DIRECTORIO['directores_principales'][:4]):
            self._format_cell_text(
                row.cells[i], nombre,
                bold=True, font_size=9, italic=True,
                alignment=WD_ALIGN_PARAGRAPH.CENTER
            )

        # Fila 5: etiqueta Directores Principales
        row = table.add_row()
        merged = row.cells[0].merge(row.cells[3])
        self._format_cell_text(
            merged, "Directores Principales",
            bold=True, font_size=9, italic=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER
        )

        # Fila 6: espacio separador
        row = table.add_row()
        row.cells[0].merge(row.cells[3]).paragraphs[0].add_run(" ")

        # Fila 7: 4 directores suplentes (uno por columna)
        row = table.add_row()
        for i, nombre in enumerate(MIEMBROS_DIRECTORIO['directores_suplentes'][:4]):
            self._format_cell_text(
                row.cells[i], nombre,
                bold=True, font_size=9, italic=True,
                alignment=WD_ALIGN_PARAGRAPH.CENTER
            )

        # Fila 8: etiqueta Directores Suplentes
        row = table.add_row()
        merged = row.cells[0].merge(row.cells[3])
        self._format_cell_text(
            merged, "Directores Suplentes",
            bold=True, font_size=9, italic=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER
        )

        # --- Secretaria del Directorio (fuera de la tabla, centrada) ---
        self.document.add_paragraph()

        p_linea = self.document.add_paragraph()
        p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linea.add_run("______________________").font.size = Pt(11)

        p_sec = self.document.add_paragraph()
        p_sec.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_sec.add_run(
            f"{MIEMBROS_DIRECTORIO['secretaria']}\nSecretaria del Directorio"
        )
        run.font.name   = 'Arial'
        run.font.size   = Pt(11)
        run.italic      = True

    # ------------------------------------------------------------------
    # GUARDAR
    # ------------------------------------------------------------------

    def guardar(self, filepath: str) -> str:
        """Guarda el documento en la ruta indicada y retorna la ruta."""
        self.document.save(filepath)
        return filepath


# ---------------------------------------------------------------------------
# FUNCIÓN AUXILIAR
# ---------------------------------------------------------------------------

def generar_nombre_archivo(numero_sesion: str, fecha: str = None) -> str:
    """
    Genera el nombre estándar del archivo .docx.
    Formato: ACTA_<numero_sesion>_<YYYYMMDD>.docx
    """
    if fecha:
        try:
            fecha_str = datetime.strptime(fecha, "%Y-%m-%d").strftime("%Y%m%d")
        except ValueError:
            fecha_str = datetime.now().strftime("%Y%m%d")
    else:
        fecha_str = datetime.now().strftime("%Y%m%d")

    return f"ACTA_{numero_sesion}_{fecha_str}.docx"
