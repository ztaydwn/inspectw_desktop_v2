from dataclasses import dataclass
from typing import Iterable, List, Tuple, Dict
from difflib import SequenceMatcher
from __future__ import annotations
import re, unicodedata, pandas as pd

_WORD_RE = re.compile(r"[a-zA-ZáéíóúñüÁÉÍÓÚÑÜ0-9]+")

def _norm(s:str)->str:
    s=s.lower()
    s=unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if unicodedata.category(ch)!="Mn")

def _tokens(s:str)->List[str]:
    return _WORD_RE.findall(_norm(s))

@dataclass
class RecConfig:
    w_jaccard: float = 0.65
    w_diff: float = 0.35
    min_score: float = 0.18
    top_k: int = 3
    stopwords: set[str] = None
    tag_weight: float = 0.7   # peso del campo TAG
    obs_weight: float = 0.3   # peso del campo OBSERVACION
    keyword_boost: Dict[str,float] = None  # {"fisura":0.05,"humedad":0.04}

class RecommendationEngine:
    def __init__(self, df: pd.DataFrame, cfg: RecConfig = RecConfig()):
        self.cfg = cfg
        cols = {c.upper(): c for c in df.columns}
        self.c_tag = cols.get("TAG")
        self.c_obs = cols.get("OBSERVACION") or cols.get("OBSERVACIÓN")
        self.c_rec = cols.get("RECOMENDACIÓN") or cols.get("RECOMENDACION")
        self.c_src = cols.get("FUENTE", None)
        if not (self.c_tag and self.c_rec):
            raise ValueError("Se requieren columnas TAG y RECOMENDACIÓN")

        def toks(s:str)->set:
            ts=set(_tokens(s))
            if self.cfg.stopwords: ts={t for t in ts if t not in self.cfg.stopwords}
            return ts

        self.rows=[]
        for _, r in df.iterrows():
            tag=str(r[self.c_tag]) if pd.notna(r[self.c_tag]) else ""
            obs=str(r[self.c_obs]) if self.c_obs and pd.notna(r[self.c_obs]) else ""
            rec=str(r[self.c_rec]) if pd.notna(r[self.c_rec]) else ""
            src=str(r[self.c_src]) if self.c_src and pd.notna(r[self.c_src]) else ""
            tok_tag=toks(tag); tok_obs=toks(obs)
            self.rows.append({"tag":tag,"obs":obs,"rec":rec,"src":src,
                              "tok_tag":tok_tag,"tok_obs":tok_obs})

    def _score(self, query:str, q_tokens:Iterable[str])->float:
        return 0.0  # placeholder (se puntúa por fila abajo)

    def suggest(self, query:str, extra_text:str="", top_k:int=None, min_score:float=None):
        cfg=self.cfg
        if top_k is None: top_k=cfg.top_k
        if min_score is None: min_score=cfg.min_score

        q = f"{query} {extra_text}".strip()
        q_tokens = set(_tokens(q))
        if cfg.stopwords:
            q_tokens={t for t in q_tokens if t not in cfg.stopwords}

        out=[]
        for row in self.rows:
            # Jaccard ponderado por campo
            inter_tag=len(q_tokens & row["tok_tag"])
            union_tag=len(q_tokens | row["tok_tag"]) or 1
            j_tag=inter_tag/union_tag

            inter_obs=len(q_tokens & row["tok_obs"])
            union_obs=len(q_tokens | row["tok_obs"]) or 1
            j_obs=inter_obs/union_obs

            jaccard = cfg.tag_weight*j_tag + cfg.obs_weight*j_obs

            # difflib sobre TAG principal
            diff = SequenceMatcher(None, _norm(query), _norm(row["tag"])).ratio()

            score = cfg.w_jaccard*jaccard + cfg.w_diff*diff

            # boosts por palabra clave
            if cfg.keyword_boost:
                for kw,bonus in cfg.keyword_boost.items():
                    if kw in q_tokens: score += bonus

            out.append((score, row["rec"]))

        out.sort(key=lambda x: x[0], reverse=True)
        return [(s, r) for s, r in out[:top_k] if s >= min_score]

def load_engine(csv_path:str, cfg:RecConfig=RecConfig())->RecommendationEngine:
    df = pd.read_csv(csv_path, sep=";", encoding="latin1")
    return RecommendationEngine(df, cfg)
