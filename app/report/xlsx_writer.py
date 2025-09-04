from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import openpyxl.drawing.image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.utils.units import pixels_to_EMU
from typing import Dict
from app.core.processing import Grupo
from app.utils.nlg_utils import agrupa_y_redacta
from PIL import Image, ImageOps
import io, math, os, re

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

def export_groups_to_xlsx_report(grupos: Dict[str, Grupo], archivos: Dict[str, bytes], output_xlsx_path: str, progress_callback=None) -> None:
    wb = Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Define styles
    header_font = Font(bold=True, size=12)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
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

    wb.save(output_xlsx_path)

