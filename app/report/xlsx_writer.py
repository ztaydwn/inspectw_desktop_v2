from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import openpyxl.drawing.image
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from typing import Dict
from app.core.processing import Grupo
from app.utils.nlg_utils import agrupa_y_redacta
from PIL import Image, ImageOps
import io, math, os, re
import unicodedata

def read_project_info(path: str) -> Dict[str, str]:
    """Lee pares clave:valor desde ``path`` y los retorna en un diccionario.
    El archivo es opcional; si no existe se devuelve un diccionario vacío
    para que el proceso continúe sin errores.
    """
    info: Dict[str, str] = {}
    if not path:
        return info
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                if ":" in line:
                    key, value = line.split(":", 1)
                    info[key.strip().lower()] = value.strip()
    except FileNotFoundError:
        pass
    return info

def parse_project_info_text(text: str) -> Dict[str, str]:
    """Convierte texto con líneas 'clave: valor' en un diccionario.
    Claves se normalizan a minúsculas conservando tildes.
    """
    info: Dict[str, str] = {}
    if not text:
        return info
    for line in text.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            info[key.strip().lower()] = value.strip()
    return info
def add_intro_sheets(wb: Workbook, info_path: str | Dict[str, str], logo_path: str = None) -> None:
    """Agrega hojas iniciales independientes al ``Workbook``. Los valores se obtienen del archivo ``infoproyect.txt`` con formato ``clave: valor`` por línea. Si el archivo no existe, las celdas quedarán vacías y el resto del proceso no se verá afectado."""
    # Permite pasar un dict ya parseado o una ruta a archivo
    if isinstance(info_path, dict):
        info = info_path
    else:
        info = read_project_info(info_path)
    # Helper para obtener valores tolerando variaciones de clave
    def _norm_key(s: str) -> str:
        s = unicodedata.normalize('NFD', s or '')
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
        s = s.lower()
        s = re.sub(r'[^a-z0-9]+', ' ', s).strip()
        return s
    info_idx = {_norm_key(k): v for k, v in info.items()}
    def iget(*names: str) -> str:
        for n in names:
            val = info_idx.get(_norm_key(n))
            if val:
                return val
        return ""
    # Preparar estilos locales (bordes y rellenos) para evitar referencias a
    # variables externas como `gray_fill` o `thin_border` que sólo existen en
    # otros contextos. Estos se usan para tablas en la hoja de desarrollo.
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    thick_top_border = Border(top=Side(style='medium'))

    # --------------------------------------------------------------------------
    # Hoja de portada
    # --------------------------------------------------------------------------
    portada = wb.create_sheet(title="PORTADA", index=0)
    portada.page_setup.orientation = portada.ORIENTATION_PORTRAIT
    portada.page_setup.paperSize = portada.PAPERSIZE_A4
    portada.page_setup.fitToWidth = 1
    portada.page_setup.fitToHeight = 1
    try:
        portada.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    portada.page_margins.left = 0.25
    portada.page_margins.right = 0.25
    portada.page_margins.top = 0.25
    portada.page_margins.bottom = 0.25

    for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        portada.column_dimensions[col].width = 12

    light_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    for row in range(1, 51): # Increased row range for A4 feel
        for col_idx in range(1, 9):
            cell = portada.cell(row=row, column=col_idx)
            cell.fill = light_fill

    # --- 1. Logo ---
    portada.row_dimensions[1].height = 25
    portada.row_dimensions[2].height = 25
    portada.row_dimensions[3].height = 25
    if logo_path and os.path.exists(logo_path):
        logo_img = OpenpyxlImage(logo_path)
        logo_img.width = 210
        logo_img.height = 75
        
        total_width_px = sum([(portada.column_dimensions[c].width * 7) + 5 for c in ["A", "B", "C", "D", "E", "F", "G", "H"]])
        x_offset_px = max(0, (total_width_px - logo_img.width) / 2)
        y_offset_px = pixels_to_EMU(15) # Small top margin

        x_offset_emu = pixels_to_EMU(x_offset_px)
        width_emu = pixels_to_EMU(logo_img.width)
        height_emu = pixels_to_EMU(logo_img.height)

        pos = XDRPoint2D(x_offset_emu, y_offset_px)
        size = XDRPositiveSize2D(width_emu, height_emu)
        logo_img.anchor = AbsoluteAnchor(pos=pos, ext=size)
        portada.add_image(logo_img)

    # --- 2. Separator line after logo ---
    portada.row_dimensions[4].height = 15
    portada.merge_cells("C4:F4")
    portada["C4"].border = thick_top_border

    # --- 3. Main Title ---
    portada.row_dimensions[5].height = 30
    portada.row_dimensions[6].height = 30
    portada.merge_cells("A5:H6")
    title_cell = portada["A5"]
    main_title = iget("titulo") or "INFORME DE SIMULACRO DE INSPECCION DE DEFENSA CIVIL EN EDIFICACIONES"
    set_cell_style(
        title_cell,
        main_title,
        bold=True,
        size=16,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    )

    # --- 4. Separator line after title ---
    portada.row_dimensions[7].height = 15
    portada.merge_cells("C7:F7")
    portada["C7"].border = thick_top_border
    
    # --- 5. Space for image ---
    portada.row_dimensions[8].height = 15
    portada.merge_cells("B9:G18")
    image_placeholder_cell = portada["B9"]
    set_cell_style(
        image_placeholder_cell,
        "Espacio para imagen",
        size=12,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    image_placeholder_cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    for row in range(9, 19):
        portada.row_dimensions[row].height = 20


    # --- 6. Project Details (better spaced) ---
    detail_rows = [
        ("NOMBRE DEL ESTABLECIMIENTO:", iget("nombre del establecimiento", "nombre", "establecimiento")),
        ("PROPIETARIO:", iget("propietario", "propietaria")),
        ("DIRECCIÓN:", iget("direccion", "dirección")),
        ("FECHA:", iget("fecha", "dia de la inspeccion")),
        ("INSPECTORES:", iget("inspectores", "profesionales designados")),
    ]
    
    start_row = 22 # Start details lower on the page
    
    # Distribute remaining space
    available_rows = 48 - start_row
    num_details = len(detail_rows)
    row_increment = available_rows // (num_details + 1) if num_details > 0 else 2

    for i, (label, value) in enumerate(detail_rows):
        current_row = start_row + (i * row_increment)
        portada.row_dimensions[current_row].height = 40

        # Label
        portada.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
        cell_label = portada.cell(row=current_row, column=2)
        set_cell_style(
            cell_label,
            label,
            bold=True,
            size=12,
            alignment=Alignment(horizontal="right", vertical="center")
        )
        # Value
        portada.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=8)
        cell_val = portada.cell(row=current_row, column=5)
        set_cell_style(
            cell_val,
            value,
            size=12,
            alignment=Alignment(horizontal="left", vertical="center", wrap_text=True)
        )

    # --- Footer ---
    footer_row = 49
    portada.row_dimensions[footer_row].height = 30
    portada.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=8)
    footer_cell = portada.cell(row=footer_row, column=1)
    set_cell_style(
        footer_cell,
        "LIMA-2025", # This could also be dynamic from info
        bold=False,
        size=12,
        alignment=Alignment(horizontal="center", vertical="center")
    )

    # --------------------------------------------------------------------------
    # Hoja de datos generales
    # --------------------------------------------------------------------------
    # Esta hoja reproduce la segunda página del informe donde se consignan los
    # datos básicos de la inspección y antecedentes. Se organizan los textos en
    # filas numeradas de acuerdo al formato.
    datos = wb.create_sheet(title="DATOS GENERALES", index=1)
    # Configuración de página A4
    datos.page_setup.orientation = datos.ORIENTATION_PORTRAIT
    datos.page_setup.paperSize = datos.PAPERSIZE_A4
    datos.page_setup.fitToWidth = 1
    datos.page_setup.fitToHeight = 1
    try:
        datos.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    datos.page_margins.left = 0.25
    datos.page_margins.right = 0.25
    datos.page_margins.top = 0.25
    datos.page_margins.bottom = 0.25
    # Definir anchos de columna
    datos.column_dimensions["A"].width = 28
    datos.column_dimensions["B"].width = 45
    datos.column_dimensions["C"].width = 5
    datos.column_dimensions["D"].width = 5
    # Encabezado principal
    datos.merge_cells("A1:D1")
    header_cell = datos["A1"]
    set_cell_style(
        header_cell,
        "INFORME DE INSPECCIÓN SIMULACRO",
        bold=True,
        size=14,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    datos.row_dimensions[1].height = 35
    # Reemplazar encabezado con el Título del infoproyecto si está disponible
    try:
        datos["A1"].value = iget("titulo") or datos["A1"].value
    except Exception:
        pass
    # Sección 1: Datos generales
    datos.merge_cells("A3:D3")
    set_cell_style(
        datos["A3"],
        "1. DATOS GENERALES",
        bold=True,
        size=12,
        alignment=Alignment(horizontal="left", vertical="center")
    )
    datos.row_dimensions[3].height = 25
    # Fila por cada subapartado
    sec1 = [
        ("1.1 PROPIETARIO:", info.get("propietario", "")),
        ("1.2 NOMBRE DE ESTABLECIMIENTO INSPECCIONADO:", info.get("nombre", "")),
        ("1.3 DIRECCIÓN DE LOCAL INSPECCIONADO:", info.get("direccion", "")),
        ("1.4 DÍA DE LA INSPECCIÓN:", ""),
        ("1.5 ESPECIALIDAD:", ""),
        ("1.6 PROFESIONALES DESIGNADOS:", ""),
        ("1.7 PERSONAL DE ACOMPAÑAMIENTO INNOVA:", ""),
        ("1.8 COMENTARIOS DEL PROCESO DE INSPECCIÓN:", ""),
    ]
    row_ptr = 4
    for label, value in sec1:
        # Etiqueta
        datos[f"A{row_ptr}"].value = label
        datos[f"A{row_ptr}"].font = Font(bold=True, size=10)
        datos[f"A{row_ptr}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # Para campos que pueden ocupar varias líneas, se fusionan varias filas
        if label.startswith("1.6") or label.startswith("1.7") or label.startswith("1.8"):
            # Reservar dos filas para estos campos
            datos.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr + 1, end_column=4)
            cell_val = datos.cell(row=row_ptr, column=2)
            set_cell_style(
                cell_val,
                value,
                alignment=Alignment(horizontal="left", vertical="top", wrap_text=True)
            )
            datos.row_dimensions[row_ptr].height = 30
            datos.row_dimensions[row_ptr + 1].height = 30
            row_ptr += 2
        else:
            datos.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr, end_column=4)
            cell_val = datos.cell(row=row_ptr, column=2)
            set_cell_style(
                cell_val,
                value,
                alignment=Alignment(horizontal="left", vertical="top", wrap_text=True)
            )
            datos.row_dimensions[row_ptr].height = 20
            row_ptr += 1
    # Sección 2: Antecedentes
    datos.merge_cells(start_row=row_ptr, start_column=1, end_row=row_ptr, end_column=4)
    set_cell_style(
        datos.cell(row=row_ptr, column=1),
        "2. ANTECEDENTES",
        bold=True,
        size=12,
        alignment=Alignment(horizontal="left", vertical="center")
    )
    datos.row_dimensions[row_ptr].height = 25
    row_ptr += 1
    # Subapartados de antecedentes
    antecedentes = [
        ("2.1 FUNCIÓN DEL ESTABLECIMIENTO:", ""),
        ("2.2 ÁREA OCUPADA:", ""),
        ("2.3 CANTIDAD DE PISOS:", ""),
        ("2.4 RIESGO:", ""),
        ("2.5 SITUACIÓN FORMAL:", ""),
    ]
    for label, value in antecedentes:
        datos[f"A{row_ptr}"].value = label
        datos[f"A{row_ptr}"].font = Font(bold=True, size=10)
        datos[f"A{row_ptr}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # Fusionar celdas para el valor
        datos.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr, end_column=4)
        set_cell_style(
            datos.cell(row=row_ptr, column=2),
            value,
            alignment=Alignment(horizontal="left", vertical="top", wrap_text=True)
        )
        datos.row_dimensions[row_ptr].height = 20
        row_ptr += 1

    # --------------------------------------------------------------------------
    # Hoja de desarrollo del simulacro
    # --------------------------------------------------------------------------
    # Esta hoja reproduce el desglose de la visita técnica y las observaciones.
    desarrollo = wb.create_sheet(title="DESARROLLO", index=2)
    # Configuración de página A4
    desarrollo.page_setup.orientation = desarrollo.ORIENTATION_PORTRAIT
    desarrollo.page_setup.paperSize = desarrollo.PAPERSIZE_A4
    desarrollo.page_setup.fitToWidth = 1
    desarrollo.page_setup.fitToHeight = 1
    try:
        desarrollo.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    desarrollo.page_margins.left = 0.25
    desarrollo.page_margins.right = 0.25
    desarrollo.page_margins.top = 0.25
    desarrollo.page_margins.bottom = 0.25
    desarrollo.column_dimensions["A"].width = 50
    desarrollo.column_dimensions["B"].width = 15
    desarrollo.column_dimensions["C"].width = 10
    desarrollo.column_dimensions["D"].width = 15
    # Encabezado principal de la sección 3
    desarrollo.merge_cells("A1:D1")
    set_cell_style(
        desarrollo["A1"],
        "3. DESARROLLO DEL SIMULACRO:",
        bold=True,
        size=12,
        alignment=Alignment(horizontal="left", vertical="center")
    )
    desarrollo.row_dimensions[1].height = 25
    # Párrafo descriptivo
    desarrollo.merge_cells("A2:D3")
    descriptive_text = (
        "Se programó una visita técnica al local en referencia, donde participaron los profesionales designados. "
        "Se realizó el recorrido por todas las instalaciones, anotando las observaciones en el acta del anexo 7A, "
        "aprobado por Reglamento de Inspecciones Técnicas de Seguridad en Edificaciones (D.S. 002-2018-PCM)."
    )
    set_cell_style(
        desarrollo["A2"],
        descriptive_text,
        size=10,
        alignment=Alignment(horizontal="justify", vertical="top", wrap_text=True)
    )
    desarrollo.row_dimensions[2].height = 40
    desarrollo.row_dimensions[3].height = 40
    # Observaciones especiales
    desarrollo["A5"].value = "Observaciones especiales:"
    desarrollo["A5"].font = Font(bold=True, size=10)
    desarrollo["A5"].alignment = Alignment(horizontal="left", vertical="center")
    # Área para observaciones (filas 5‑7, columnas B‑D)
    desarrollo.merge_cells(start_row=5, start_column=2, end_row=7, end_column=4)
    obs_cell = desarrollo.cell(row=5, column=2)
    obs_cell.border = thin_border
    obs_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    obs_cell.value = ""  # área en blanco para completar
    desarrollo.row_dimensions[5].height = 25
    desarrollo.row_dimensions[6].height = 25
    desarrollo.row_dimensions[7].height = 25
    # Tabla de condiciones sobre la edificación
    start_table_row = 9
    # Título de la tabla
    desarrollo.merge_cells(start_row=start_table_row, start_column=1, end_row=start_table_row, end_column=4)
    set_cell_style(
        desarrollo.cell(row=start_table_row, column=1),
        "SOBRE LA EDIFICACIÓN:",
        bold=True,
        size=10,
        fill=gray_fill,
        border=thin_border,
        alignment=Alignment(horizontal="left", vertical="center")
    )
    desarrollo.row_dimensions[start_table_row].height = 20
    # Fila de encabezados de la tabla (después del título)
    header_row = start_table_row + 1
    # Primera columna: descripción general con varias líneas
    desarrollo.merge_cells(start_row=header_row, start_column=1, end_row=header_row, end_column=2)
    set_cell_style(
        desarrollo.cell(row=header_row, column=1),
        "CONDICIÓN DE SEGURIDAD OBSERVADA\n(Según tabla de D.S. 007-2018-PCM – Anexo 7A)",
        bold=True,
        size=9,
        fill=gray_fill,
        border=thin_border,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    )
    # Segunda columna: Sí / No
    set_cell_style(
        desarrollo.cell(row=header_row, column=3),
        "Sí / No",
        bold=True,
        size=9,
        fill=gray_fill,
        border=thin_border,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    )
    # Tercera columna: No corresponde
    set_cell_style(
        desarrollo.cell(row=header_row, column=4),
        "No corresponde",
        bold=True,
        size=9,
        fill=gray_fill,
        border=thin_border,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    )
    desarrollo.row_dimensions[header_row].height = 25
    # Detalle de cada condición (filas 11‑14)
    condiciones = [
        (
            "1. No se encuentra en proceso de construcción según lo establecido en el artículo único de la Norma G.040 "
            "Definiciones del Reglamento Nacional de Edificaciones",
            "SI",
            ""
        ),
        (
            "2. Cuenta con servicios de agua, electricidad, y los que resulten esenciales para el desarrollo de sus "
            "actividades, debidamente instalados e implementados.",
            "SI",
            ""
        ),
        (
            "3. Cuenta con mobiliario básico e instalado para el desarrollo de la actividad.",
            "SI",
            ""
        ),
        (
            "4. Tiene los equipos o artefactos debidamente instalados o ubicados, respectivamente, en los lugares de uso "
            "habitual o permanente.",
            "SI",
            ""
        ),
    ]
    current = header_row + 1
    for descripcion, si_no, no_corresponde in condiciones:
        # Descripción ocupa columnas A–B
        desarrollo.merge_cells(start_row=current, start_column=1, end_row=current, end_column=2)
        set_cell_style(
            desarrollo.cell(row=current, column=1),
            descripcion,
            size=9,
            border=thin_border,
            alignment=Alignment(horizontal="left", vertical="top", wrap_text=True)
        )
        # Columna Sí/No
        set_cell_style(
            desarrollo.cell(row=current, column=3),
            si_no,
            size=9,
            border=thin_border,
            alignment=Alignment(horizontal="center", vertical="center")
        )
        # Columna No corresponde
        set_cell_style(
            desarrollo.cell(row=current, column=4),
            no_corresponde,
            size=9,
            border=thin_border,
            alignment=Alignment(horizontal="center", vertical="center")
        )
        desarrollo.row_dimensions[current].height = 35
        current += 1
    # Comentarios adicionales
    comentarios_row = current + 1
    desarrollo[f"A{comentarios_row}"] = "Comentarios adicionales al respecto:"
    desarrollo[f"A{comentarios_row}"].font = Font(bold=True, size=10)
    desarrollo[f"A{comentarios_row}"].alignment = Alignment(horizontal="left", vertical="center")
    # Área para comentarios (celdas B–D varias filas)
    desarrollo.merge_cells(start_row=comentarios_row, start_column=2, end_row=comentarios_row + 2, end_column=4)
    comentarios_cell = desarrollo.cell(row=comentarios_row, column=2)
    comentarios_cell.border = thin_border
    comentarios_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    comentarios_cell.value = ""
    desarrollo.row_dimensions[comentarios_row].height = 25
    desarrollo.row_dimensions[comentarios_row + 1].height = 25
    desarrollo.row_dimensions[comentarios_row + 2].height = 25

    # ------------------------------------------------------------------
    # Completar DATOS GENERALES con valores del infoproyecto si existen
    # ------------------------------------------------------------------
    try:
        datos["B4"].value = iget("propietario", "propietaria") or datos["B4"].value
        datos["B5"].value = iget("nombre del establecimiento", "nombre", "establecimiento") or datos["B5"].value
        datos["B6"].value = iget("direccion", "dirección") or datos["B6"].value
        datos["B7"].value = iget("fecha", "dia de la inspeccion", "día de la inspección") or datos["B7"].value
        datos["B8"].value = iget("especialidad") or datos["B8"].value
        datos["B9"].value = iget("inspectores", "profesionales designados") or datos["B9"].value
        datos["B11"].value = iget("acompañamiento", "acompanamiento", "personal de acompañamiento") or datos["B11"].value
        datos["B13"].value = iget("comentarios", "comentarios del proceso") or datos["B13"].value
    except Exception:
        pass


def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s[0])]

def apply_border_to_range(ws, start_cell, end_cell, border_style='thin'):
    """Aplica bordes a un rango de celdas."""
    border = Border(
        left=Side(style=border_style),
        right=Side(style=border_style),
        top=Side(style=border_style),
        bottom=Side(style=border_style)
    )
    
    # Convertir referencias de celda a coordenadas
    start_coord = coordinate_from_string(start_cell)
    end_coord = coordinate_from_string(end_cell)
    start_col = column_index_from_string(start_coord[0])
    end_col = column_index_from_string(end_coord[0])
    
    for row in range(start_coord[1], end_coord[1] + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row=row, column=col).border = border

def set_cell_style(cell, text, bold=False, size=11, alignment=None, fill=None, border=None):
    cell.value = text
    cell.font = Font(bold=bold, size=size)
    if alignment:
        cell.alignment = alignment
    if fill:
        cell.fill = fill
    if border:
        cell.border = border

def estimate_visual_lines(text: str, chars_per_line: int) -> int:
    """Estima el número de líneas visuales que ocupará un texto con word-wrap."""
    if not text or chars_per_line <= 0:
        return 1
    
    total_lines = 0
    for line_segment in text.split('\n'):
        total_lines += math.ceil(len(line_segment) / chars_per_line) if line_segment else 1
    return total_lines

def export_groups_to_xlsx_report(
    grupos: Dict[str, Grupo],
    archivos: Dict[str, bytes],
    output_xlsx_path: str,
    progress_callback=None,
    info_path: str = os.path.join("datos", "infoproyect.txt"),
    control_documents=None,
    conclusiones: list[str] | None = None,
    ) -> None:
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Agregar hojas independientes iniciales
    # Intentar leer 'infoproyect.txt' desde los archivos cargados (ZIP/carpeta)
    info_from_archivos = None
    try:
        for k, v in archivos.items():
            base = os.path.basename(k).lower()
            if base.startswith('infoproyect') and base.endswith('.txt'):
                try:
                    info_from_archivos = parse_project_info_text(v.decode('utf-8', errors='ignore'))
                    break
                except Exception:
                    pass
    except Exception:
        pass
    add_intro_sheets(wb, info_from_archivos if info_from_archivos else info_path, logo_path="datos/portadat.png")

    # Define styles
    header_font = Font(bold=True, size=12)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    red_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Natural sort for group names
    sorted_grupos = sorted(grupos.items(), key=natural_sort_key)
    total_grupos = len(sorted_grupos)

    for idx, (gname, grupo) in enumerate(sorted_grupos):
        # Replace invalid characters for sheet titles
        invalid_chars = ['/', '\\', '?', '*', '[', ']']
        sanitized_gname = gname
        for char in invalid_chars:
            sanitized_gname = sanitized_gname.replace(char, '-')
        
        sheet_name = sanitized_gname[:31]  # Sheet name limit is 31 chars
        ws = wb.create_sheet(title=sheet_name)

        # --- Configuración de Página para Impresión ---
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0 # Permite que se extienda a varias páginas de alto

        # Establecer márgenes estrechos (en pulgadas)
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75

        # --- Header banner (condición de seguridad) ---
        banner_text = (
            "CONDICIÓN DE SEGURIDAD OBSERVADA:\n"
            "SEGÚN TABLA DE D.S. 007-2018-PCM (ANEXO 7A)"
        )
        ws.merge_cells('A1:C1')
        banner_cell = ws['A1']
        set_cell_style(
            banner_cell,
            banner_text,
            bold=True,
            size=11,
            alignment=Alignment(horizontal='left', vertical='center', wrap_text=True),
            fill=gray_fill,
            border=thin_border,
        )
        apply_border_to_range(ws, 'A1', 'C1')
        banner_lines = banner_text.count('\n') + 1
        ws.row_dimensions[1].height = max(28, banner_lines * 18)

        # --- Title ---
        ws.merge_cells('A2:C2')
        title_cell = ws['A2']
        set_cell_style(title_cell, gname, bold=True, size=12)
        
        title_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        
        # Ajustar cálculo de líneas al nuevo ancho total (aprox 90 chars)
        text_lines = len(gname) // 90 + 1
        ws.row_dimensions[2].height = max(25, text_lines * 20)
        
        apply_border_to_range(ws, 'A2', 'C2')

        # --- Photo section header ---
        ws.merge_cells('A3:C3')
        set_cell_style(ws['A3'], "FOTOGRAFÍAS:", bold=True, size=12, alignment=Alignment(horizontal='left', vertical='center'))
        ws['A3'].font = header_font
        apply_border_to_range(ws, 'A3', 'C3')
        
        # --- Photo file names ---
        cols, rows = 3, 2
        per_page = cols * rows
        num_fotos = len(grupo.fotos)
        pages = math.ceil(num_fotos / per_page) if per_page else 0

        # Ancho de celda aprox 230px con el nuevo ancho de columna
        image_cell_height_px = 240
        
        current_row = 4
        for page in range(pages):
            chunk = grupo.fotos[page * per_page:(page + 1) * per_page]
            
            # Procesar cada fila de fotos y añadir una fila de etiquetas debajo
            for r in range(rows):
                photo_row_idx = current_row + (r * 2)
                label_row_idx = photo_row_idx + 1

                ws.row_dimensions[photo_row_idx].height = image_cell_height_px * 0.75 # 180
                ws.row_dimensions[label_row_idx].height = 20

                for c in range(cols):
                    chunk_idx = r * cols + c
                    if chunk_idx >= len(chunk):
                        break # No hay más fotos en esta página

                    idx_global = page * per_page + chunk_idx + 1
                    foto = chunk[chunk_idx]
                    cell_pos = f"{get_column_letter(c + 1)}{photo_row_idx}"
                    
                    possible_paths = [
                        f"{foto.carpeta}/{foto.filename}",
                        f"{foto.carpeta}\\{foto.filename}",
                        foto.filename,
                        f"{foto.carpeta.replace('/', '')}\\{foto.filename}"
                    ]
                    
                    img_data = None
                    for path in possible_paths:
                        img_data = archivos.get(path)
                        if img_data:
                            break
                    
                    if img_data:
                        try:
                            img = Image.open(io.BytesIO(img_data))
                            img = ImageOps.exif_transpose(img)
                            if img.mode in ("RGBA", "LA", "P"): img = img.convert("RGB")
                            
                            # No reducir la resolución. Insertar original y ajustar tamaño de visualización.
                            cell_w_px = 229 # Ancho de celda (32 unidades) en píxeles
                            cell_h_px = 240 # Alto de celda (180 pt) en píxeles
                            
                            # Calcular dimensiones de visualización manteniendo el aspect ratio
                            # Dejar un pequeño margen para evitar desbordes
                            margin = 4 
                            ratio = min((cell_w_px - margin) / img.width, (cell_h_px - margin) / img.height)
                            display_width, display_height = int(img.width * ratio), int(img.height * ratio)
                            
                            img_bytes = io.BytesIO()
                            img.save(img_bytes, format='PNG') # Guardar original en buffer
                            img_bytes.seek(0)
                            img_excel = openpyxl.drawing.image.Image(img_bytes)

                            # Asignar tamaño de visualización y anclar a la celda.
                            # Este método es más compatible con versiones antiguas de openpyxl.
                            img_excel.width = display_width
                            img_excel.height = display_height
                            img_excel.anchor = cell_pos
                            ws.add_image(img_excel)

                            # Añadir etiqueta [Foto x] en la celda de abajo
                            label_cell_coord = f"{get_column_letter(c + 1)}{label_row_idx}"
                            set_cell_style(ws[label_cell_coord], f"[Foto {idx_global}]", size=9, alignment=Alignment(horizontal='center', vertical='center'), border=thin_border)

                        except Exception as e:
                            print(f"Error procesando imagen {foto.filename}: {str(e)}")
                            ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"
                    else:
                        ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"
            
            # Incrementar el puntero de fila para la siguiente página
            current_row += rows * 2 # 2 filas por cada fila de fotos (foto + etiqueta)
            if page < pages - 1:
                ws.row_dimensions[current_row].height = 15 # Espacio entre páginas
                current_row += 1
        current_row += 1

        # --- Details Header ---
        details_header_cell = ws[f'A{current_row}']
        set_cell_style(details_header_cell, "UBICACIÓN Y DETALLE:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        current_row += 1

        # --- Details Content ---
        entradas = []
        for i, foto in enumerate(grupo.fotos, start=1):
            full_detail = foto.specific_detail
            detail_after_plus = full_detail.split('+', 1)[1].strip() if '+' in full_detail else full_detail
            entradas.append((detail_after_plus, f"{foto.carpeta} [Foto {i}]")) # Usar el índice global
            
        oraciones = agrupa_y_redacta(entradas, umbral_similitud=0.8)
        details_text = "\n".join(f"{i}. {sentencia}" for i, sentencia in enumerate(oraciones, start=1))

        chars_per_line_details = 90 # Ancho de 3 columnas
        details_lines_visual = estimate_visual_lines(details_text, chars_per_line_details)
        needed_rows_details = max(4, details_lines_visual)

        details_content_cell = ws[f'A{current_row}']
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + needed_rows_details - 1, end_column=3)
        for i in range(needed_rows_details):
            ws.row_dimensions[current_row + i].height = 16
        set_cell_style(details_content_cell, details_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'))
        apply_border_to_range(
            ws,
            f'A{current_row}',
            f'C{current_row + needed_rows_details - 1}'
        )
        current_row += needed_rows_details

        # --- Recommendations Header ---
        rec_header_cell = ws[f'A{current_row}']
        set_cell_style(rec_header_cell, "RECOMENDACIONES:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        current_row += 1

        # --- Recommendations Content ---
        recs = getattr(grupo, "recomendaciones", None) or []
        rec_text = "\n".join(f"• {r}" for r in recs) if recs else "—"
        chars_per_line_recs = 90
        rec_lines_visual = estimate_visual_lines(rec_text, chars_per_line_recs)
        needed_rows_recs = max(4, rec_lines_visual)

        rec_content_cell = ws[f'A{current_row}']
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + needed_rows_recs - 1, end_column=3)
        for i in range(needed_rows_recs):
            ws.row_dimensions[current_row + i].height = 16
        set_cell_style(rec_content_cell, rec_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'), fill=green_fill)
        apply_border_to_range(
            ws,
            f'A{current_row}',
            f'C{current_row + needed_rows_recs - 1}'
        )

        # Adjust column widths
        ws.column_dimensions['A'].width = 32
        ws.column_dimensions['B'].width = 32
        ws.column_dimensions['C'].width = 32

        if progress_callback:
            progress_percentage = int(((idx + 1) / total_grupos) * 100)
            progress_callback.emit(progress_percentage)

    # ------------------------------------------------------------------
    # 5. CONTROL DE DOCUMENTACIÓN DE SEGURIDAD (Hojas finales opcionales)
    # ------------------------------------------------------------------
    def _add_control_docs_sheet(wb: Workbook, page_title: str, items_slice: list[tuple[int, str, str]]):
        ws = wb.create_sheet(title=page_title)
        # Configuración de página: A4, orientación vertical, 1 página de ancho y alto, márgenes estrechos
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        # Forzar el uso de FitToPage en algunos visores
        try:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:
            pass
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25
        # Centrar ligeramente para mejor presentación
        ws.print_options.horizontalCentered = True
        # Anchos de columna similares a la maqueta
        ws.column_dimensions['A'].width = 5
        # Reducimos el ancho para garantizar 1 página de ancho
        ws.column_dimensions['B'].width = 60
        ws.column_dimensions['C'].width = 32

        # Título
        ws.merge_cells('A1:C1')
        set_cell_style(
            ws['A1'],
            '5. CONTROL DE DOCUMENTACIÓN DE SEGURIDAD',
            bold=True,
            size=12,
            alignment=Alignment(horizontal='left', vertical='center')
        )
        ws.row_dimensions[1].height = 25

        # Encabezados
        ws['A3'].value = 'N°'
        ws['B3'].value = 'CERTIFICADOS, CONSTANCIAS Y/O PROTOCOLO'
        ws['C3'].value = 'SITUACION'
        for col in ['A', 'B', 'C']:
            cell = ws[f'{col}3']
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.fill = gray_fill
            cell.border = thin_border
        ws.row_dimensions[3].height = 22

        # Filas
        row = 4
        for num, descripcion, situacion in items_slice:
            set_cell_style(ws[f'A{row}'], str(num), size=10, alignment=Alignment(horizontal='center', vertical='top'), border=thin_border)

            desc_cell = ws[f'B{row}']
            set_cell_style(desc_cell, descripcion, size=10, alignment=Alignment(wrap_text=True, vertical='top'), border=thin_border)

            sit_cell = ws[f'C{row}']
            # Determinar color según el contenido de situacion
            sit_text = situacion or ''
            low = sit_text.lower()
            fill = None
            if 'no aplica' in low:
                fill = gray_fill
                # Normalizamos el texto para que al menos diga NO APLICA
                if not sit_text.strip():
                    sit_text = 'NO APLICA'
            elif 'observado' in low or 'observación' in low or 'observacion' in low:
                fill = red_fill
            elif 'correcto' in low or 'cumple' in low:
                fill = green_fill
            set_cell_style(sit_cell, sit_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'), fill=fill, border=thin_border)

            # Altura de fila estimada
            # Estimar con chars_per_line acordes a los nuevos anchos
            est = max(2, estimate_visual_lines(descripcion, 55), estimate_visual_lines(sit_text, 28))
            ws.row_dimensions[row].height = 18 * est
            row += 1

        # Bordes de tabla ya asignados celda a celda con border=thin_border
        # Limitar el área de impresión exactamente a la tabla construida
        ws.print_area = f"A1:C{row-1}"
        return ws

    # Si se proporcionó control_documents, construir hojas
    if control_documents:
        # Descripciones fijas de los 22 ítems (según el formato mostrado)
        items_descriptions = [
            "Certificado vigente de medición de resistencia del sistema de puesta a tierra: De conformidad con el Código Nacional de Electricidad, el valor de la medición de resistencia del sistema de puesta a tierra no debe exceder los 25 ohmios. El certificado de dicha medición debe encontrarse vigente (la medición de la resistencia del pozo a tierra debe realizarse anualmente) y estar firmado por un ingeniero electricista o mecánico electricista, colegiado y habilitado.",
            "Certificado de sistema de detección y alarma de incendios: Debe indicar la cantidad y ubicación de detectores del sistema de detección y alarma de incendios centralizada con que cuenta el Establecimiento, incluye el protocolo de pruebas de operatividad y/o mantenimiento del sistema. Se debe considerar lo señalado en Art. 52 al 65 de la Norma A.130 del RNE, y la inspección, prueba y mantenimiento según Cap. 14 de la NFPA 72.",
            "Certificado de extintores: Debe indicar la cantidad, ubicación, numeración, tipo y peso de los extintores instalados en el Establecimiento, incluye los protocolos de pruebas de operatividad y/o mantenimiento de los extintores. Considerar lo señalado en art. 163 al 165 de la Norma A.130 RNE y NTP 350.043-1.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Rociadores: Su elaboración según el literal A) del art. 102 de la Norma A.130 RNE; la inspección, prueba y mantenimiento según estándar NFPA 25 según lo establecido en el articulo 27.1 de la NFPA 13.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Rociadores especiales tipo Spray: Su elaboración según el literal B) del art. 102 de la Norma A.130 RNE; la inspección, prueba y mantenimiento según estándar NFPA 25 según lo establecido en el articulo 11.1.1 de la NFPA 15.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Redes Principales de Protección Contra Incendios enterradas (casos de fabricas, almacenes, otros): Su elaboración según el literal C) del art. 102 de la Norma A.130 RNE; la inspección, prueba y mantenimiento según estándar NFPA 25 según lo establecido en el articulo 14.1 de la NFPA 24.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Montantes y Gabinetes de Agua Contra Incendio: Su elaboración según el literal H) del art. 102 de la Norma A.130 RNE; la inspección, prueba y mantenimiento según estándar NFPA 25 según lo establecido en el articulo 13.1 de la NFPA 14.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento de las Bombas de Agua Contra Incendio: Su elaboración según el art. 152 de la Norma A.130 RNE; la inspección, prueba y mantenimiento según estándar NFPA 25 según lo establecido en el articulo 14.4 de la NFPA 20. Incluyen las pruebas de presión hidrostática.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de las luces de emergencia: Su elaboración según la Sección 010-010 (3) del Código Nacional de Electricidad – Normas de Utilización. Mantenimiento según manual del fabricante.",
            "Protocolo de pruebas de operatividad y/o las puertas cortafuego y sus dispositivos como marcos, bisagras cierrapuertas, manija, cerradura o barra antipánico: Su certificación para uso cortafuego, según los artículos 10 y 11 de la Norma A.130 RNE. Mantenimiento según el manual del fabricante.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema de administración de humos: Su elaboración según literal b) del Art. 94 de la Norma A.130 del RNE; la inspección, prueba y mantenimiento según Capítulo 8 del estándar NFPA 92 según lo establecido en la Guía NFPA 92B.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema de Presurización de Escaleras de Evacuación: Su elaboración según Sub Capitulo IV. Requisitos de los Sistemas de Presurización de Escaleras de la Norma A.130 del RNE; la inspección, prueba y mantenimiento según artículo 7.3 del Capítulo 4.6 y capítulo 8 de la NFPA 92.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema Mecánico de Extracción de Monóxido de Carbono: Su elaboración según el art.69 de la Norma A.010. Condiciones Generales del Diseño del RNE.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del Teléfono de Emergencia en Ascensor: Su elaboración según los literales C) y D) del art.30 de la Norma A.010. Condiciones Generales del Diseño del RNE; art. 19 de la Norma A.130. Requisitos de Seguridad del RNE.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del Teléfono de Bomberos: Según la NFPA 72. Para la elaboración de las memorias o protocolos de pruebas de operatividad y mantenimiento de los equipos de seguridad y protección contraincendios, se debe cumplir con los requerimientos mínimos establecidos en la normatividad señalada en los párrafos precedentes, en las especificaciones técnicas de los fabricantes, estándares y otras que resulten aplicables, para tales efectos puede hacer uso de los formatos sugeridos por las normas NFPA u otros aplicables.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de Ascensor, Montacarga, Escaleras mecánicas y equipos de elevación eléctrica, firmado por ing. mecánico, electricista o mecánico electricista colegiado y habilitado.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de Equipos de Aire Acondicionado.",
            "Certificado de vidrios templados expedido por el fabricante.",
            "Certificado de laminado de vidrios y/o espejos.",
            "Constancia de registro de hidrocarburos emitido por  OSINERGMIN, además de la constancia de Operatividad y mantenimiento de la red de interna de GLP y/o líquido combustible, emitido por empresa o profesional especializado.  NTP 321.121",
            "Certificado de pintura ignífuga en maderas.",
            "OTROS (por ejemplo: Protocolo de aislamiento de tableros).",
        ]

        # Normalizar diferentes estructuras de entrada
        # Acepta: {1: 'texto'}, [{'numero':1,'situacion':'...'}], [('1','texto')], etc.
        norm: dict[int, str] = {}
        if isinstance(control_documents, dict):
            for k, v in control_documents.items():
                try:
                    num = int(k)
                    norm[num] = str(v) if v is not None else ''
                except Exception:
                    continue
        elif isinstance(control_documents, (list, tuple)):
            for item in control_documents:
                if isinstance(item, dict):
                    num = item.get('numero') or item.get('num') or item.get('id')
                    if num is None:
                        continue
                    try:
                        num = int(num)
                    except Exception:
                        continue
                    norm[num] = str(item.get('situacion', ''))
                elif isinstance(item, (list, tuple)) and len(item) >= 2:
                    try:
                        num = int(item[0])
                    except Exception:
                        continue
                    norm[num] = str(item[1])

        # Construir lista total de (n, descripcion, situacion)
        full_items = []
        for i, desc in enumerate(items_descriptions, start=1):
            full_items.append((i, desc, norm.get(i, 'NO APLICA')))

        # Rebanar en páginas (como en las imágenes: 1-8, 9-16, 17-22)
        pages = [
            ('CONTROL DOC. (1)', full_items[0:8]),
            ('CONTROL DOC. (2)', full_items[8:16]),
            ('CONTROL DOC. (3)', full_items[16:22]),
        ]
        created = []
        for title, slice_items in pages:
            if slice_items:
                created.append(_add_control_docs_sheet(wb, title, slice_items))

        # Agregar conclusiones (opcional) en la última hoja
        if conclusiones and created:
            ws = created[-1]
            # Buscar primera fila libre
            last_row = ws.max_row + 2
            ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=3)
            set_cell_style(ws.cell(row=last_row, column=1), '6. CONCLUSIONES:', bold=True, size=12, alignment=Alignment(horizontal='left', vertical='center'))
            ws.row_dimensions[last_row].height = 24
            row = last_row + 1
            for i, txt in enumerate(conclusiones, start=1):
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
                set_cell_style(ws.cell(row=row, column=1), f"{i}. {txt}", size=10, alignment=Alignment(wrap_text=True, vertical='top'))
                ws.row_dimensions[row].height = 18 * max(2, estimate_visual_lines(txt, 90))
                row += 1

    wb.save(output_xlsx_path)
