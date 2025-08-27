import zipfile, re
from dataclasses import dataclass, field
from typing import Dict

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
        if not line.strip():
            continue
        # Dividir por el primer tabulador que encuentre
        parts = line.strip().split('\t', 1)
        if len(parts) == 2:
            key = parts[0].strip()
            value = parts[1].strip()
            lookup[key] = value
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
            desc_parts = desc_content.split('\t', 1)
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

def procesar_zip(archivos: Dict[str, bytes]) -> Dict[str, Grupo]:
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
        
    # TODO: cargar historico.csv, asignar recomendaciones
    return grupos
