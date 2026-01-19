import zipfile, re, os
from dataclasses import dataclass, field, asdict
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
    recommendation: str = ""  # Recomendación específica de descriptions.txt
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
            # Normalizar separadores de ruta a '/' para consistencia
            normalized_name = n.replace('\\', '/')
            out[normalized_name] = zf.read(n)
    return out

def cargar_directorio(path_dir: str) -> Dict[str, bytes]:
    """
    Lee todos los archivos de un directorio y sus subdirectorios y los retorna
    en un diccionario similar al de cargar_zip.
    """
    out = {}
    base_path = Path(path_dir)
    for root, _, files in os.walk(base_path):
        for name in files:
            file_path = Path(root) / name
            # La clave es la ruta relativa al directorio base, usando '/' como separador
            relative_path = file_path.relative_to(base_path).as_posix()
            out[relative_path] = file_path.read_bytes()
    return out

def _find_image_data(archivos: Dict[str, bytes], foto: Foto) -> bytes | None:
    """
    Busca los datos de una imagen en el diccionario de archivos de forma robusta.
    Intenta varias combinaciones de rutas para maximizar la compatibilidad.
    """
    # 1. La ruta ideal y más común (normalizada)
    path1 = f"{foto.carpeta}/{foto.filename}"
    if path1 in archivos:
        return archivos[path1]

    # 2. Ruta con separadores de Windows (por si acaso)
    path2 = f"{foto.carpeta}\\{foto.filename}".replace('/', '\\')
    if path2 in archivos:
        return archivos[path2]

    # 3. Solo el nombre del archivo (si está en la raíz)
    if foto.filename in archivos:
        return archivos[foto.filename]

    return None

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
            lookup[key] = key + " " + parts[1].strip()
    return lookup

def _parse_descriptions(txt_descriptions: str, group_lookup: Dict[str, str], archivos: Dict[str, bytes]) -> tuple[list[Foto], list[str]]:
    """
    Parsea descriptions.txt, empareja fotos de forma robusta y asigna el nombre oficial del grupo.

    Retorna una tupla: (lista de fotos encontradas, lista de advertencias).
    """
    fotos = []
    warnings = []
    current_block = None

    lines = txt_descriptions.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        line_num = i + 1
        m = re.match(r'\[(.+?)\]\s+(\S+\.jpg)', line, flags=re.I)
        if m:
            # Si había un bloque anterior, procesarlo
            if current_block and 'description' in current_block:
                foto_obj = _create_foto_from_block(current_block, group_lookup, archivos, warnings, line_num)
                if foto_obj:
                    fotos.append(foto_obj)
            current_block = {"carpeta": m.group(1), "filename": m.group(2), "description": "", "recommendation": ""}
        elif line.lower().startswith("description:") and current_block:
            desc_content = line.split(":", 1)[1].strip()
            desc_parts = re.split(r'\s+', desc_content, 1)
            numbering_code = desc_parts[0].strip()
            specific_detail = desc_parts[1].strip() if len(desc_parts) > 1 else ''
            current_block["description"] = specific_detail
            current_block["numbering_code"] = numbering_code
        elif line.lower().startswith("recommendation:") and current_block:
            rec_text = line.split(":", 1)[1].strip()
            if rec_text and not current_block.get("recommendation"):
                current_block["recommendation"] = rec_text
        i += 1

    # Process last block
    if current_block and 'description' in current_block:
        foto_obj = _create_foto_from_block(current_block, group_lookup, archivos, warnings, len(lines))
        if foto_obj:
            fotos.append(foto_obj)

    return fotos, warnings

def _create_foto_from_block(block: dict, group_lookup: Dict[str, str], archivos: Dict[str, bytes], warnings: list[str], line_num: int) -> Foto | None:
    """Helper to create Foto from parsed block."""
    filename_from_desc = block["filename"]
    carpeta_from_desc = block["carpeta"]
    numbering_code = block.get("numbering_code", "")
    specific_detail = block["description"]
    recommendation = block["recommendation"]

    official_group_name = group_lookup.get(numbering_code, f"Grupo no encontrado para '{numbering_code}'")

    # --- Lógica de emparejamiento robusto ---
    # Intento 1: Ruta ideal (carpeta/archivo.jpg)
    ideal_path = f"{carpeta_from_desc}/{filename_from_desc}"
    if ideal_path not in archivos:
        # Intento 2: Fallback - buscar solo por nombre de archivo
        matches = [path for path in archivos if os.path.basename(path) == filename_from_desc]
        if len(matches) == 1:
            # Éxito: se encontró una única coincidencia. Se corrige la carpeta.
            carpeta_from_desc = os.path.dirname(matches[0]).replace('\\', '/')
        elif len(matches) > 1:
            warnings.append(f"Línea {line_num}: Nombre de archivo '{filename_from_desc}' es ambiguo (encontrado en {len(matches)} ubicaciones). Se omitió la foto.")
            return None
        else:
            warnings.append(f"Línea {line_num}: No se encontró la foto '{filename_from_desc}' en ninguna carpeta.")
            return None

    return Foto(
        filename=filename_from_desc,
        group_name=official_group_name,
        specific_detail=specific_detail,
        carpeta=carpeta_from_desc,
        recommendation=recommendation
    )

def asignar_recomendaciones(grupos: Dict[str, Grupo], engine: RecommendationEngine, top_k: int = 1):
    """Rellena grupo.recomendaciones usando el motor, solo si no hay recomendaciones de descriptions.txt."""
    for g in grupos.values():
        if not g.recomendaciones:  # Solo asignar si no hay recomendaciones de descriptions.txt
            # Usa descripción base + agregación de detalles/ubicaciones para contextualizar la consulta
            extra = ", ".join(sorted({f"{f.carpeta} {f.specific_detail}".strip() for f in g.fotos if f.specific_detail or f.carpeta}))[:400]
            sugerencias = engine.suggest(query=g.descripcion, extra_text=extra, top_k=top_k)
            g.recomendaciones = [rec for _, rec in sugerencias] or g.recomendaciones

def procesar_zip(archivos: Dict[str, bytes], hist_path: str | None = None) -> tuple[Dict[str, Grupo], str | None]:
    # Leer ambos archivos de texto del zip
    txt_descriptions = archivos.get("descriptions.txt", b"").decode("utf-8", errors="ignore")
    txt_grupos = archivos.get("grupos.txt", b"").decode("utf-8", errors="ignore")
    
    if not txt_descriptions or not txt_grupos:
        return {}, "Faltan 'descriptions.txt' o 'grupos.txt' en el origen de datos."
        
    group_lookup = _create_group_lookup(txt_grupos)
    
    fotos, parsing_warnings = _parse_descriptions(txt_descriptions, group_lookup, archivos)
    
    grupos: Dict[str, Grupo] = {}
    for f in fotos:
        # Agrupar por el nombre oficial del grupo
        g = grupos.setdefault(f.group_name, Grupo(descripcion=f.group_name))
        g.fotos.append(f)
        # Collect recommendations from descriptions.txt
        if f.recommendation:
            g.recomendaciones.append(f.recommendation)

    # Cargar histórico y asignar recomendaciones solo si no hay de descriptions.txt
    error_msg = None
    try:
        hp = hist_path or HIST_DEFAULT
        engine = load_engine(hp)
        asignar_recomendaciones(grupos, engine, top_k=2)
    except Exception as e:
        print(f"[WARN] No se pudo cargar histórico: {e}")
        error_msg = f"Error cargando recomendaciones: {e}"
    
    if parsing_warnings:
        error_msg = (error_msg + "\n\n" if error_msg else "") + "\n".join(parsing_warnings)
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
