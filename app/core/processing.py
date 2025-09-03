import zipfile, re
from dataclasses import dataclass, field
from typing import Dict
from pathlib import Path
try:
    from app.core.paths import resource_path
except Exception:
    # Fallback por si el import falla en el .exe (PyInstaller)
    from pathlib import Path
    import sys
    def resource_path(rel: str) -> str:
        base = getattr(sys, "_MEIPASS", None)  # carpeta temporal del .exe
        if base:
            return str(Path(base) / rel)
        # raíz del proyecto en desarrollo
        return str((Path(__file__).resolve().parents[2] / rel))
    
from app.core.recommend import load_engine
from app.core.recommend import load_engine, RecommendationEngine

# Ruta opcional al histórico; si es None se usa el valor por defecto
HIST_DEFAULT = resource_path("datos/historico.csv")

@dataclass
class Foto:
    filename: str
    group_name: str      # El nombre oficial del grupo desde grupos.txt
    specific_detail: str # La pregunta/detalle específico de descriptions.txt
    carpeta: str
    path_tmp: str | None = None

@dataclass
class Grupo:
    descripcion: str # Este será el group_name
    fotos: list[Foto] = field(default_factory=list)
    recomendaciones: list[str] = field(default_factory=list)

def cargar_zip(path_zip: str) -> Dict[str, bytes]:
    out = {}
    with zipfile.ZipFile(path_zip, "r") as zf:
        for n in zf.namelist():
            out[n] = zf.read(n)
    return out

def _create_group_lookup(txt_grupos: str) -> Dict[str, str]:
    """Convierte el texto de grupos.txt en un diccionario para búsqueda rápida."""
    lookup = {}
    for line in txt_grupos.splitlines():
        full_line = line.strip()
        if not full_line:
            continue
        
        # La clave sigue siendo solo el código, para una búsqueda limpia
        parts = re.split(r'\s+', full_line, 1)
        if len(parts) == 2:
            key = parts[0].strip()
            # El valor ahora es la línea completa, para no perder la numeración
            lookup[key] = full_line
    return lookup

def _parse_descriptions(txt_descriptions: str, group_lookup: Dict[str, str]) -> list[Foto]:
    """Parsea descriptions.txt y usa el lookup de grupos para asignar el nombre oficial."""
    fotos = []
    bloque = {}

    for line in txt_descriptions.splitlines():
        line = line.strip()
        # Asume que la info de la foto sigue viniendo en este formato
        m = re.match(r'\[(.+?)\]\s+(\S+\.jpg)', line, flags=re.I)
        if m:
            bloque = {"carpeta": m.group(1), "filename": m.group(2)}
        elif line.lower().startswith("description:") and bloque:
            desc_content = line.split(":", 1)[1].strip()
            
            # Extraer el código numérico y el detalle específico
            desc_parts = re.split(r'\s+', desc_content, 1)
            numbering_code = desc_parts[0].strip()
            specific_detail = desc_parts[1].strip() if len(desc_parts) > 1 else ''
            
            # Buscar el nombre oficial del grupo usando el código
            official_group_name = group_lookup.get(numbering_code, f"Grupo no encontrado para '{numbering_code}'")

            fotos.append(Foto(
                filename=bloque["filename"],
                group_name=official_group_name,
                specific_detail=specific_detail,
                carpeta=bloque["carpeta"],
            ))
            bloque = {}
    return fotos
def asignar_recomendaciones(grupos: Dict[str, Grupo], engine: RecommendationEngine, top_k: int = 1):
    """Rellena grupo.recomendaciones usando el motor."""
    for g in grupos.values():
        # Usa descripción base + agregación de detalles/ubicaciones para contextualizar la consulta
        extra = ", ".join(sorted({f"{f.carpeta} {f.specific_detail}".strip() for f in g.fotos if f.specific_detail or f.carpeta}))[:400]
        sugerencias = engine.suggest(query=g.descripcion, extra_text=extra, top_k=top_k)
        g.recomendaciones = [rec for _, rec in sugerencias] or g.recomendaciones

def procesar_zip(archivos: Dict[str, bytes], hist_path: str | None = None) -> tuple[Dict[str, Grupo], str | None]:
    # Leer ambos archivos de texto del zip
    txt_descriptions = archivos.get("descriptions.txt", b"").decode("utf-8", errors="ignore")
    txt_grupos = archivos.get("grupos.txt", b"").decode("utf-8", errors="ignore")
    
    # Crear el mapa de búsqueda desde grupos.txt
    group_lookup = _create_group_lookup(txt_grupos)
    
    # Parsear las descripciones usando el mapa
    fotos = _parse_descriptions(txt_descriptions, group_lookup)
    
    grupos: Dict[str, Grupo] = {}
    for f in fotos:
        # Agrupar por el nombre oficial del grupo
        g = grupos.setdefault(f.group_name, Grupo(descripcion=f.group_name))
        g.fotos.append(f)
    
    # Cargar histórico y asignar recomendaciones    
    error_msg = None
    try:
        hp = hist_path or HIST_DEFAULT
        engine = load_engine(hp)
        asignar_recomendaciones(grupos, engine, top_k=2)
    except Exception as e:
        print(f"[WARN] No se pudo cargar histórico: {e}")
        error_msg = str(e)
    
    return grupos, error_msg

def reaplicar_recomendaciones(grupos: Dict[str, Grupo], hist_path: str) -> str | None:
    """Toma grupos existentes y aplica/re-aplica recomendaciones desde un archivo."""
    error_msg = None
    if not hist_path:
        return "No se proporcionó una ruta al archivo histórico."
    try:
        engine = load_engine(hist_path)
        asignar_recomendaciones(grupos, engine, top_k=2)
    except Exception as e:
        error_msg = str(e)
    return error_msg