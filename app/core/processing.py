import zipfile, re
from dataclasses import dataclass, field
from typing import Dict, List

@dataclass
class Foto:
    filename: str
    description_base: str
    detalle: str
    carpeta: str
    path_tmp: str | None = None  # opcional si extraes a disco

@dataclass
class Grupo:
    descripcion: str
    fotos: List[Foto] = field(default_factory=list)
    recomendaciones: List[str] = field(default_factory=list)

def cargar_zip(path_zip: str) -> Dict[str, bytes]:
    out = {}
    with zipfile.ZipFile(path_zip, "r") as zf:
        for n in zf.namelist():
            out[n] = zf.read(n)
    return out

def _parse_txt(txt: str) -> List[Foto]:
    fotos = []
    bloque = {}
    for line in txt.splitlines():
        line = line.strip()
        m = re.match(r'\[(.+?)\]\s+(\S+\.jpg)', line, flags=re.I)
        if m:
            bloque = {"carpeta": m.group(1), "filename": m.group(2)}
        elif line.lower().startswith("description:") and bloque:
            desc = line.split(":", 1)[1].strip()
            base, detalle = (x.strip() for x in desc.split("+", 1)) if "+" in desc else (desc, "")
            fotos.append(Foto(
                filename=bloque["filename"],
                description_base=base,
                detalle=detalle,
                carpeta=bloque["carpeta"],
            ))
            bloque = {}
    return fotos

def procesar_zip(archivos: Dict[str, bytes]) -> Dict[str, Grupo]:
    txt = archivos.get("descriptions.txt", b"").decode("utf-8", errors="ignore")
    fotos = _parse_txt(txt)
    grupos: Dict[str, Grupo] = {}
    for f in fotos:
        g = grupos.setdefault(f.description_base, Grupo(descripcion=f.description_base))
        g.fotos.append(f)
    # TODO: cargar historico.csv, asignar recomendaciones
    return grupos
