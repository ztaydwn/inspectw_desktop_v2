from pathlib import Path
import sys

def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", None)  # carpeta temporal de PyInstaller
    if base:
        return str(Path(base) / rel)
    # raíz del repo: .../inspectw_desktop/
    return str((Path(__file__).resolve().parents[2] / rel))