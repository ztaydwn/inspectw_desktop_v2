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
def add_intro_sheets(wb: Workbook, info_path: str, logo_path: str = None) -> None:
    """Agrega hojas iniciales independientes al ``Workbook``.
    Los valores se obtienen del archivo ``infoproyect.txt`` con formato
    ``clave: valor`` por línea. Si el archivo no existe, las celdas quedarán
    vacías y el resto del proceso no se verá afectado.
    """
    info = read_project_info(info_path)
    # Preparar estilos locales (bordes y rellenos) para evitar referencias a
    # variables externas como `gray_fill` o `thin_border` que sólo existen en
    # otros contextos. Estos se usan para tablas en la hoja de desarrollo.
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # --------------------------------------------------------------------------
    # Hoja de portada
    # --------------------------------------------------------------------------
    # La portada se diseña como una plantilla con un aspecto similar al modelo
    # proporcionado. Se definen anchos de columna amplios, alturas de fila y
    # combinaciones de celdas para lograr una disposición agradable. Las celdas
    # usadas se rellenan con un tono gris claro para simular el fondo del
    # informe.
    portada = wb.create_sheet(title="PORTADA", index=0)
    # Definir el ancho de las columnas (A–H) para dar un margen amplio en A4
    for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        portada.column_dimensions[col].width = 11
    # Colocar un color de fondo claro en el área de trabajo
    light_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    for row in range(1, 22):
        for col_idx in range(1, 9):
            cell = portada.cell(row=row, column=col_idx)
            cell.fill = light_fill
    
    portada.row_dimensions[1].height = 30
    portada.row_dimensions[2].height = 30
    portada.row_dimensions[3].height = 30

        # Insertar logo (fila 1–3)
    if logo_path:
        logo_img = OpenpyxlImage(logo_path)
        logo_img.width = 140  # Ajusta según necesidades
        logo_img.height = 90
        
        # --- Centrar logo en el área A1:H3 ---
        # 1. Calcular dimensiones del área en píxeles
        # Ancho (aprox): (caracteres * 7) + 5
        total_width_px = sum([(portada.column_dimensions[c].width * 7) + 5 for c in ["A", "B", "C", "D", "E", "F", "G", "H"]])
        # Alto: puntos * 4/3
        total_height_px = sum([portada.row_dimensions[r].height * 4/3 for r in [1, 2, 3]])

        # 2. Calcular offset para centrar
        x_offset_px = max(0, (total_width_px - logo_img.width) / 2)
        y_offset_px = max(0, (total_height_px - logo_img.height) / 2)

        # --- AJUSTE MANUAL DE CENTRADO ---
        # Modifique este valor para ajustar la posición horizontal del logo.
        # Un valor positivo lo mueve a la derecha, un valor negativo a la izquierda.
        manual_adjustment_px = -20  # Cambie este valor, por ejemplo a 20 o -30
        x_offset_px += manual_adjustment_px

        # 3. Convertir a EMUs y crear ancla absoluta
        x_offset_emu = pixels_to_EMU(x_offset_px)
        y_offset_emu = pixels_to_EMU(y_offset_px)
        width_emu = pixels_to_EMU(logo_img.width)
        height_emu = pixels_to_EMU(logo_img.height)

        pos = XDRPoint2D(x_offset_emu, y_offset_emu)
        size = XDRPositiveSize2D(width_emu, height_emu)
        logo_img.anchor = AbsoluteAnchor(pos=pos, ext=size)
        
        portada.add_image(logo_img)
        
    # Espacio reservado para el logo en las filas 1‑3
    portada.merge_cells("A1:H3")
    logo_cell = portada["A1"]
    set_cell_style(
        logo_cell,
        "",  # Dejar sin texto; se podría insertar una imagen aquí si estuviera disponible
        alignment=Alignment(horizontal="center", vertical="center")
    )
    # Título principal del informe en filas 5‑6
    portada.merge_cells("A5:H6")
    title_cell = portada["A5"]
    set_cell_style(
        title_cell,
        "INFORME DE SIMULACRO DE INSPECCIÓN DE DEFENSA CIVIL EN EDIFICACIONES",
        bold=True,
        size=14,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    )
    portada.row_dimensions[5].height = 45
    portada.row_dimensions[6].height = 45
    # Datos del proyecto: etiquetas y valores
    detail_rows = [
        ("NOMBRE DEL ESTABLECIMIENTO:", info.get("nombre", "")),
        ("PROPIETARIO:", info.get("propietario", "")),
        ("DIRECCIÓN:", info.get("direccion", "")),
    ]
    start_row = 9
    for label, value in detail_rows:
        # Etiqueta en columnas B-D, alineada a la derecha
        portada.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=4)
        cell_label = portada.cell(row=start_row, column=2)
        set_cell_style(
            cell_label,
            label,
            bold=True,
            alignment=Alignment(horizontal="right", vertical="center")
        )
        # Valor en columnas E-G, alineado a la izquierda
        portada.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=7)
        cell_val = portada.cell(row=start_row, column=5)
        set_cell_style(
            cell_val,
            value,
            alignment=Alignment(horizontal="left", vertical="center", wrap_text=True)
        )
        # Ajustar altura de fila
        portada.row_dimensions[start_row].height = 30
        start_row += 2
    # Texto al pie de página con la ubicación y año (p. ej. LIMA‑2025)
    footer_row = 19
    portada.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=8)
    footer_cell = portada.cell(row=footer_row, column=1)
    set_cell_style(
        footer_cell,
        "LIMA-2025",
        bold=False,
        size=12,
        alignment=Alignment(horizontal="center", vertical="center")
    )
    portada.row_dimensions[footer_row].height = 30

    # --------------------------------------------------------------------------
    # Hoja de datos generales
    # --------------------------------------------------------------------------
    # Esta hoja reproduce la segunda página del informe donde se consignan los
    # datos básicos de la inspección y antecedentes. Se organizan los textos en
    # filas numeradas de acuerdo al formato.
    datos = wb.create_sheet(title="DATOS GENERALES", index=1)
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
    add_intro_sheets(wb, info_path, logo_path="datos/portadat.png")

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

        # --- Title ---
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        set_cell_style(title_cell, gname, bold=True, size=12)
        
        title_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        
        text_lines = len(gname) // 50 + 1
        ws.row_dimensions[1].height = max(25, text_lines * 20)
        
        apply_border_to_range(ws, 'A1', 'D1')

        # --- Photo section header ---
        ws['A3'] = "FOTOGRAFÍAS:"
        ws['A3'].font = header_font
        apply_border_to_range(ws, 'A3', 'D3')
        
        # --- Photo file names ---
        cols, rows = 4, 3
        per_page = cols * rows
        num_fotos = len(grupo.fotos)
        pages = math.ceil(num_fotos / per_page) if per_page else 0

        image_height = 180
        
        current_row = 4
        for page in range(pages):
            chunk = grupo.fotos[page * per_page:(page + 1) * per_page]
            
            for r in range(rows):
                ws.row_dimensions[current_row + r].height = image_height * 0.75
                for c in range(cols):
                    cell_coord = f"{get_column_letter(c + 1)}{current_row + r}"
                    apply_border_to_range(ws, cell_coord, cell_coord)
                
            if page < pages - 1:
                ws.row_dimensions[current_row + rows].height = 15
            
            for idx_foto, foto in enumerate(chunk):
                r = idx_foto // cols
                c = idx_foto % cols
                cell_pos = f"{get_column_letter(c + 1)}{current_row + r}"
                
                posible_paths = [
                    f"{foto.carpeta}/{foto.filename}",
                    f"{foto.carpeta}\\{foto.filename}",
                    foto.filename,
                    f"{foto.carpeta.replace('/', '')}\\{foto.filename}"
                ]
                
                img_data = None
                for path in posible_paths:
                    img_data = archivos.get(path)
                    if img_data:
                        break
                
                if img_data:
                    try:
                        img = Image.open(io.BytesIO(img_data))
                        img = ImageOps.exif_transpose(img)
                        if img.mode in ("RGBA", "LA", "P"): img = img.convert("RGB")
                        ratio = min(220 / img.width, image_height / img.height)
                        width, height = int(img.width * ratio), int(img.height * ratio)
                        img = img.resize((width, height), Image.BICUBIC)
                        img_bytes = io.BytesIO()
                        img.save(img_bytes, format='PNG')
                        img_bytes.seek(0)
                        img_excel = openpyxl.drawing.image.Image(img_bytes)
                        img_excel.anchor = cell_pos
                        ws.add_image(img_excel)
                    except Exception as e:
                        print(f"Error procesando imagen {foto.filename}: {str(e)}")
                        ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"
                else:
                    ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"
                    
            current_row += rows + 1

        current_row += 1

        # --- Details and Recommendations Headers ---
        details_header_cell = ws[f'A{current_row}']
        set_cell_style(details_header_cell, "UBICACIÓN Y DETALLE:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)

        rec_header_cell = ws[f'C{current_row}']
        set_cell_style(rec_header_cell, "RECOMENDACIONES:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
        
        current_row += 1

        # --- Details and Recommendations Content ---
        entradas = []
        for i, foto in enumerate(grupo.fotos, start=1):
            full_detail = foto.specific_detail
            detail_after_plus = full_detail.split('+', 1)[1].strip() if '+' in full_detail else full_detail
            entradas.append((detail_after_plus, f"{foto.carpeta} [Foto {i}]"))
            
        oraciones = agrupa_y_redacta(entradas, umbral_similitud=0.8)
        details_text = "\n".join(f"{i}. {sentencia}" for i, sentencia in enumerate(oraciones, start=1))
        
        recs = getattr(grupo, "recomendaciones", None) or []
        rec_text = "\n".join(f"• {r}" for r in recs) if recs else "—"
        
        # --- Nueva Lógica de Cálculo de Altura ---
        chars_per_line = 70
        details_lines_visual = estimate_visual_lines(details_text, chars_per_line)
        rec_lines_visual = estimate_visual_lines(rec_text, chars_per_line)
        
        needed_rows = max(8, max(details_lines_visual, rec_lines_visual))
        
        details_content_cell = ws[f'A{current_row}']
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + needed_rows - 1, end_column=2)
        apply_border_to_range(ws, f'A{current_row}', f'B{current_row + needed_rows - 1}')
        
        rec_content_cell = ws[f'C{current_row}']
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row + needed_rows - 1, end_column=4)
        apply_border_to_range(ws, f'C{current_row}', f'D{current_row + needed_rows - 1}')
        
        for i in range(needed_rows):
            ws.row_dimensions[current_row + i].height = 16

        set_cell_style(details_content_cell, details_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'))
        details_content_cell.border = thin_border
        
        set_cell_style(rec_content_cell, rec_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'), fill=green_fill)
        rec_content_cell.border = thin_border

        # Adjust column widths
        ws.column_dimensions['A'].width = 23
        ws.column_dimensions['B'].width = 23
        ws.column_dimensions['C'].width = 23
        ws.column_dimensions['D'].width = 23

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
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25
        # Anchos de columna similares a la maqueta
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 90
        ws.column_dimensions['C'].width = 45

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
            est = max(2, estimate_visual_lines(descripcion, 70), estimate_visual_lines(sit_text, 35))
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
            "Certificado vigente de medición de resistencia del sistema de puesta a tierra: ... (valor no debe exceder los 25 ohmios; firmado por profesional colegiado y habilitado).",
            "Certificado de sistema de detección y alarma de incendios: cantidad y ubicación de detectores; incluye protocolo de pruebas de operatividad y/o mantenimiento; considerar NFPA 72 y Norma A.130 REN.",
            "Certificado de extintores: cantidad, ubicación, numeración, tipo y peso de los extintores instalados; incluye protocolos de operatividad y/o mantenimiento; Norma A.130 RNE y NTP 350.043-1.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Rociadores (literal A) art. 102 Norma A.130 RNE; NFPA 13.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Rociadores especiales tipo Spray (literal B) art. 102 Norma A.130 RNE; NFPA 15.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Redes Principales de Protección Contra Incendios enterradas (literal C) art. 102 Norma A.130 RNE; NFPA 24.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento del Sistema de Montantes y Gabinetes de Agua Contra Incendio (literal H) art. 102 Norma A.130 RNE; NFPA 14.",
            "Protocolos de Pruebas de Operatividad y/o Mantenimiento de las Bombas de Agua Contra Incendio (art. 152 Norma A.130 RNE); NFPA 20; incluye pruebas de presión hidrostatica.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de las luces de emergencia según Código Nacional de Electricidad – Normas de Utilización y manual del fabricante.",
            "Protocolo de pruebas de operatividad y/o las puertas cortafuego y sus dispositivos; certificación para uso cortafuego; Norma A.130 RNE; manual del fabricante.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema de administración de humos (literal b) Art. 94 de la Norma A.130 del RNE; Guía NFPA 92B.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema de Presurización de Escaleras de Evacuación; Norma A.130 del RNE; NFPA 92.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del sistema Mecánico de Extracción de Monóxido de Carbono; art.69 Norma A.010; Condiciones Generales del Diseño del RNE.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del Teléfono de Emergencia en Ascensor; art.30 Norma A.010; art.19 Norma A.130.",
            "Protocolo de pruebas de operatividad y/o mantenimiento del Teléfono de Bomberos; NFPA 72.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de Ascensor, Montacarga, Escaleras mecánicas y equipos de elevación eléctrica; firmado por profesional colegiado y habilitado.",
            "Protocolo de pruebas de operatividad y/o mantenimiento de Equipos de Aire Acondicionado.",
            "Certificado de vidrios templados expedido por el fabricante.",
            "Certificado de laminado de vidrios y/o espejos.",
            "Constancia de registro de hidrocarburos emitido por OSINERGMIN y constancia de operatividad y mantenimiento de la red interna de GLP o líquido combustible. NTP 321.121.",
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
