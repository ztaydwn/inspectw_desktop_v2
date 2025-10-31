"""
Main Excel report writer module that orchestrates all sheet generation.
"""

from typing import Dict
import os

from app.core.processing import Grupo
from app.utils.nlg_utils import agrupa_y_redacta

from .utils import read_project_info, parse_project_info_text
from .portada_sheet import add_portada_sheet
from .datos_generales_sheet import add_datos_generales_sheet
from .desarrollo_sheet import add_desarrollo_sheet
from .group_sheets import add_group_sheets
from .control_documents_sheet import add_control_documents_sheets


def add_intro_sheets(wb, info_path: str | Dict[str, str], logo_path: str | None = None) -> None:
    """Agrega hojas iniciales independientes al Workbook.

    Los valores se obtienen del archivo infoproyect.txt con formato clave: valor por línea.
    Si el archivo no existe, las celdas quedarán vacías y el resto del proceso no se verá afectado.
    """
    # Permite pasar un dict ya parseado o una ruta a archivo
    if isinstance(info_path, dict):
        info = info_path
    else:
        info = read_project_info(info_path)

    add_portada_sheet(wb, info, logo_path)
    add_datos_generales_sheet(wb, info)
    add_desarrollo_sheet(wb)


def export_groups_to_xlsx_report(
    grupos: Dict[str, Grupo],
    archivos: Dict[str, bytes],
    output_xlsx_path: str,
    progress_callback=None,
    info_path: str = os.path.join("datos", "infoproyect.txt"),
    control_documents=None,
    conclusiones: list[str] | None = None,
) -> None:
    """Exporta grupos a un reporte Excel con múltiples hojas."""
    from openpyxl import Workbook

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

    # Agregar hojas de grupos
    add_group_sheets(wb, grupos, archivos, progress_callback)

    # Agregar hojas de control de documentos
    add_control_documents_sheets(wb, control_documents, conclusiones)

    wb.save(output_xlsx_path)
