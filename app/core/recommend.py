from __future__ import annotations
from dataclasses import dataclass
from typing import Iterable, List, Tuple, Dict
from difflib import SequenceMatcher
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
    min_score: float = 0.10
    top_k: int = 3
    stopwords: set[str] = None
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

        q_tokens_for_tag = set(_tokens(query))
        if cfg.stopwords:
            q_tokens_for_tag={t for t in q_tokens_for_tag if t not in cfg.stopwords}

        # Step 1: Find the best matching TAG
        best_tag_score = -1.0
        best_tag = ""
        
        unique_tags = {row["tag"]: row["tok_tag"] for row in self.rows}

        for tag, tok_tag in unique_tags.items():
            diff = SequenceMatcher(None, _norm(query), _norm(tag)).ratio()
            
            inter_tag = len(q_tokens_for_tag & tok_tag)
            union_tag = len(q_tokens_for_tag | tok_tag) or 1
            j_tag = inter_tag / union_tag
            
            tag_score = 0.7 * diff + 0.3 * j_tag
            
            if tag_score > best_tag_score:
                best_tag_score = tag_score
                best_tag = tag

        TAG_MATCH_THRESHOLD = 0.35 
        if best_tag_score < TAG_MATCH_THRESHOLD:
            return []

        # Step 2: Rank recommendations within the selected TAG group
        q_tokens_for_obs = set(_tokens(f"{query} {extra_text}".strip()))
        if cfg.stopwords:
            q_tokens_for_obs={t for t in q_tokens_for_obs if t not in cfg.stopwords}
            
        out = []
        for row in self.rows:
            if row["tag"] == best_tag:
                if not row["obs"]:
                    score = 0.1
                else:
                    inter_obs = len(q_tokens_for_obs & row["tok_obs"])
                    union_obs = len(q_tokens_for_obs | row["tok_obs"]) or 1
                    j_obs = inter_obs / union_obs

                    obs_text_for_diff = extra_text if extra_text else query
                    diff_obs = SequenceMatcher(None, _norm(obs_text_for_diff), _norm(row["obs"])).ratio()

                    score = 0.6 * j_obs + 0.4 * diff_obs

                if cfg.keyword_boost:
                    for kw, bonus in cfg.keyword_boost.items():
                        if kw in q_tokens_for_obs:
                            score += bonus
                
                if row["rec"] and score >= min_score:
                    out.append((score, row["rec"]))

        seen_recs = {}
        for score, rec in sorted(out, key=lambda x: x[0], reverse=True):
            if rec not in seen_recs:
                seen_recs[rec] = score
                
        final_recs = sorted(seen_recs.items(), key=lambda item: item[1], reverse=True)

        return [(s, r) for r, s in final_recs[:top_k]]

def load_engine(csv_path:str, cfg:RecConfig=RecConfig())->RecommendationEngine:
    df = pd.read_csv(csv_path, sep=";", encoding="latin1")
    return RecommendationEngine(df, cfg)
