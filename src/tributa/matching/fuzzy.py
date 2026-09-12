"""
Motor de fuzzy matching para desempate de NCMs por CEST.
Usa rapidfuzz + intersecao de palavras-chave + detecao de volumes.
"""
import os
import re
import json

from rapidfuzz import fuzz as rfuzz

from ..config import console, STOPWORDS


def load_manual_overrides():
    """
    Carrega overrides manuais de data/manual_overrides.json.
    Retorna dict { ncm: cest_code }.
    """
    override_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.dirname(__file__)))), "data", "manual_overrides.json")

    FALLBACK_MAP = {
        "2203": "0302100",
        "22071000": "0600101",
        "22072000": "0600100",
        "22072010": "0600100",
        "21069010": "0301000",
        "22011000": "0300100",
        "22029000": "0301100",
        "22029900": "0301100",
    }

    if os.path.exists(override_path):
        try:
            with open(override_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            result = {ncm: info["cest"] for ncm, info in raw.items() if "cest" in info}
            console.print(f"[dim]  Overrides manuais carregados: {len(result)} NCMs mapeados[/dim]")
            return result
        except Exception as e:
            console.print(f"[yellow]![/yellow] Erro ao ler overrides: {e}. Usando fallback.")

    return FALLBACK_MAP


def find_best_cest(sped_desc, ncm, cest_db, manual_map=None):
    """
    Compara descricao SPED com registros CEST do CONFAZ.
    Retorna tupla (cest_code, score).
    """
    if manual_map and ncm in manual_map:
        return manual_map[ncm], 1.0

    def tokens(text):
        words = set(re.findall(
            r"[a-z\u00e1\u00e0\u00e2\u00e3\u00e9\u00e8\u00ea\u00ed\u00ef\u00f3\u00f4\u00f5\u00fa\u00fc\u00e7]+",
            text.lower(),
        ))
        return words - STOPWORDS

    def extract_vol(text):
        return set(re.findall(r"(\d+)\s*(?:ml|litros|litro|l)\b", text.lower()))

    def combined_score(desc1, desc2):
        fuzzy = rfuzz.token_sort_ratio(desc1.lower(), desc2.lower()) / 100.0
        t1, t2 = tokens(desc1), tokens(desc2)
        kw = len(t1 & t2) / min(len(t1), len(t2)) if t1 and t2 else 0.0
        v1, v2 = extract_vol(desc1), extract_vol(desc2)
        vol_score = 0.0
        if v1 and v2:
            vol_score = 1.0 if v1 & v2 else -0.5
        return max(0.0, 0.4 * fuzzy + 0.6 * kw + vol_score)

    candidates = []

    # 1. Busca por NCM exato ou prefixo
    for confaz_ncm, entries in cest_db.items():
        if confaz_ncm == ncm or confaz_ncm.startswith(ncm) or ncm.startswith(confaz_ncm):
            for e in entries:
                score = combined_score(sped_desc, e["descricao"])
                candidates.append((score, e["cest"], e["descricao"]))

    # 2. Fallback: busca global
    if not candidates:
        for confaz_ncm, entries in cest_db.items():
            for e in entries:
                score = combined_score(sped_desc, e["descricao"])
                if score >= 0.4:
                    candidates.append((score, e["cest"], e["descricao"]))

    if not candidates:
        return "", 0.0

    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][0]