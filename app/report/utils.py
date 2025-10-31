"""
Shared utility functions for Excel report generation.
"""

import unicodedata
import re
from typing import Dict
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl import Workbook


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


def normalize_key(s: str) -> str:
    """Helper para normalizar claves de información del proyecto."""
    s = unicodedata.normalize('NFD', s or '')
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = s.lower()
    s = re.sub(r'[^a-z0-9]+', ' ', s).strip()
    return s


def create_info_index(info: Dict[str, str]) -> Dict[str, str]:
    """Crea un índice normalizado de información del proyecto."""
    return {normalize_key(k): v for k, v in info.items()}


def get_info_value(info_idx: Dict[str, str], *names: str) -> str:
    """Obtiene valores de información tolerando variaciones de clave."""
    for n in names:
        val = info_idx.get(normalize_key(n))
        if val:
            return val
    return ""


def apply_border_to_range(ws, start_cell, end_cell, border_style='thin'):
    """Aplica bordes a un rango de celdas."""
    border = Border(
        left=Side(style=border_style),
        right=Side(style=border_style),
        top=Side(style=border_style),
        bottom=Side(style=border_style)
    )

    # Convertir referencias de celda a coordenadas
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    start_coord = coordinate_from_string(start_cell)
    end_coord = coordinate_from_string(end_cell)
    start_col = column_index_from_string(start_coord[0])
    end_col = column_index_from_string(end_coord[0])

    for row in range(start_coord[1], end_coord[1] + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row=row, column=col).border = border


def set_cell_style(cell, text, bold=False, size=11, alignment=None, fill=None, border=None):
    """Establece el estilo de una celda."""
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


def natural_sort_key(s):
    """Key function for natural sorting of strings with numbers."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s[0])]


# Import math here to avoid circular imports
import math