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
import io, math, os

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

def export_groups_to_xlsx_report(grupos: Dict[str, Grupo], archivos: Dict[str, bytes], output_xlsx_path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Define styles
    title_font = Font(bold=True, size=16)
    header_font = Font(bold=True, size=12)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for gname, grupo in grupos.items():
        # Replace invalid characters for sheet titles
        invalid_chars = ['/', '\\', '?', '*', '[', ']']
        sanitized_gname = gname
        for char in invalid_chars:
            sanitized_gname = sanitized_gname.replace(char, '-')
        
        sheet_name = sanitized_gname[:31]  # Sheet name limit is 31 chars
        ws = wb.create_sheet(title=sheet_name)

        # --- Title ---
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        set_cell_style(title_cell, gname, bold=True, size=12)
        
        # Configurar el ajuste de texto y alineación para el título
        title_cell.alignment = Alignment(wrap_text=True, 
                                       vertical='center',
                                       horizontal='left')
        
        # Ajustar altura de fila automáticamente según el contenido
        text_lines = len(gname) // 50 + 1  # Aproximadamente 50 caracteres por línea
        ws.row_dimensions[1].height = max(25, text_lines * 20)
        
        # Añadir borde al título
        apply_border_to_range(ws, 'A1', 'D1')

        # --- Photo section header ---
        ws['A3'] = "FOTOGRAFÍAS:"
        ws['A3'].font = header_font
        apply_border_to_range(ws, 'A3', 'D3')
        
        # --- Photo file names ---
        # This is different from pptx, we'll list the files.
        # We can have a layout similar to the pptx one.
        cols, rows = 4, 3
        per_page = cols * rows
        num_fotos = len(grupo.fotos)
        pages = math.ceil(num_fotos / per_page) if per_page else 0

        # Ajustar altura de fila para las imágenes
        image_height = 180  # altura en píxeles aumentada
        
        current_row = 4
        for page in range(pages):
            chunk = grupo.fotos[page * per_page:(page + 1) * per_page]
            
            # Ajustar altura de las filas para las imágenes
            for r in range(rows):
                ws.row_dimensions[current_row + r].height = image_height * 0.75  # Convertir píxeles a puntos
                # Añadir bordes a las celdas de las fotos
                for c in range(cols):
                    cell_coord = f"{get_column_letter(c + 1)}{current_row + r}"
                    apply_border_to_range(ws, cell_coord, cell_coord)
                
            # Reducir el espacio entre grupos de fotos
            if page < pages - 1:  # No añadir espacio extra después del último grupo
                ws.row_dimensions[current_row + rows].height = 15  # Espacio pequeño entre grupos
            
            for idx, foto in enumerate(chunk):
                r = idx // cols
                c = idx % cols
                cell_pos = f"{get_column_letter(c + 1)}{current_row + r}"
                
                # Intentar diferentes variaciones de la ruta
                posible_paths = [
                    f"{foto.carpeta}/{foto.filename}",
                    f"{foto.carpeta}\\{foto.filename}",
                    foto.filename,
                    f"{foto.carpeta.replace('/', '\\')}\\{foto.filename}"
                ]
                
                img_data = None
                for path in posible_paths:
                    img_data = archivos.get(path)
                    if img_data:
                        break
                
                if img_data:
                    try:
                        # Procesar la imagen
                        img = Image.open(io.BytesIO(img_data))
                        
                        # Convertir a RGB si es necesario
                        if img.mode in ("RGBA", "LA", "P"):
                            img = img.convert("RGB")
                        
                        # Calcular el nuevo tamaño manteniendo la proporción
                        max_width = 220  # ancho máximo en píxeles
                        
                        # Calcular ratio para mantener proporción
                        ratio = min(max_width / img.width, image_height / img.height)
                        width = int(img.width * ratio)
                        height = int(img.height * ratio)
                        
                        # Redimensionar usando alta calidad
                        img = img.resize((width, height), Image.LANCZOS)
                        
                        # Guardar en memoria
                        img_bytes = io.BytesIO()
                        img.save(img_bytes, format='PNG')
                        img_bytes.seek(0)
                        
                        # Añadir imagen a Excel
                        img_excel = openpyxl.drawing.image.Image(img_bytes)
                        
                        # Coordenadas de la celda
                        coord = coordinate_from_string(cell_pos)
                        col_idx = column_index_from_string(coord[0])
                        
                        # Centrado simple
                        x_offset = 20  # Pequeño offset fijo para centrar aproximadamente
                        
                        # Ajustar la posición de la imagen
                        img_excel.anchor = cell_pos
                        ws.add_image(img_excel)
                            
                    except Exception as e:
                        print(f"Error procesando imagen {foto.filename}: {str(e)}")
                        ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"  # Fallback a texto
                        
                    except Exception as e:
                        print(f"Error procesando imagen {foto.filename}: {str(e)}")
                        ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"  # Fallback a texto
                else:
                    ws[cell_pos] = f"{foto.carpeta}/{foto.filename}"  # Si no se encuentra la imagen
                    
            current_row += rows + 1  # Add some space between pages of photos

        current_row += 1 # Space before the next section

        # --- Details and Recommendations Headers ---
        details_header_cell = ws[f'A{current_row}']
        set_cell_style(details_header_cell, "UBICACIÓN Y DETALLE:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)

        rec_header_cell = ws[f'C{current_row}']
        set_cell_style(rec_header_cell, "RECOMENDACIONES:", bold=True, size=11, fill=gray_fill, border=thin_border)
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
        
        current_row += 1

        # --- Details and Recommendations Content ---
        # Preparar el texto primero para calcular el número de líneas
        entradas = []
        for idx, foto in enumerate(grupo.fotos, start=1):
            foto_num = idx
            full_detail = foto.specific_detail
            detail_after_plus = full_detail.split('+', 1)[1].strip() if '+' in full_detail else full_detail
            entradas.append((detail_after_plus, f"{foto.carpeta} [Foto {foto_num}]"))
            
        oraciones = agrupa_y_redacta(entradas, umbral_similitud=0.8)
        details_text = "\n".join(f"{idx}. {sentencia}" for idx, sentencia in enumerate(oraciones, start=1))
        
        # Calcular número de líneas para detalles
        details_lines = len(details_text.split('\n'))
        
        # Calcular número de líneas para recomendaciones
        recs = getattr(grupo, "recomendaciones", None) or []
        rec_text = "\n".join(f"• {r}" for r in recs) if recs else "—"
        rec_lines = len(rec_text.split('\n'))
        
        # Calcular altura necesaria (mínimo 6 filas, más si el contenido lo requiere)
        # Asumimos aproximadamente 1.2 líneas de texto por fila de Excel para mejor legibilidad
        needed_rows = max(6, math.ceil(max(details_lines, rec_lines) * 1.2))
        
        details_content_cell = ws[f'A{current_row}']
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + needed_rows - 1, end_column=2)
        # Aplicar borde a la sección de detalles
        apply_border_to_range(ws, f'A{current_row}', f'B{current_row + needed_rows - 1}')
        
        rec_content_cell = ws[f'C{current_row}']
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row + needed_rows - 1, end_column=4)
        # Aplicar borde a la sección de recomendaciones
        apply_border_to_range(ws, f'C{current_row}', f'D{current_row + needed_rows - 1}')
        
        # Ajustar la altura de las filas para el contenido
        for i in range(needed_rows):
            ws.row_dimensions[current_row + i].height = 20  # altura en puntos

        # --- Set Content ---
        set_cell_style(details_content_cell, details_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'))
        details_content_cell.border = thin_border
        
        set_cell_style(rec_content_cell, rec_text, size=10, alignment=Alignment(wrap_text=True, vertical='top'), fill=green_fill)
        rec_content_cell.border = thin_border

        # Adjust column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 25

    wb.save(output_xlsx_path)
