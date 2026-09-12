"""
Parser de NCMs da Tabela SPED 4.3.10.
Extrai NCMs, Excecoes (EX) e intervalos de blocos de texto.
"""
import re


def expand_range(start_ncm, end_ncm):
    """Expande um intervalo de NCMs (ex: '22.01 a 22.03') em lista individual."""
    start = int(start_ncm.replace(".", ""))
    end = int(end_ncm.replace(".", ""))
    return [str(i).zfill(len(start_ncm.replace(".", ""))) for i in range(start, end + 1)]


def parse_ncms(text_block):
    """
    Extrai todos os NCMs e suas Excecoes (EX) de um bloco de texto da tabela SPED.
    Retorna lista de dicts: [{'ncm': '22011000', 'ex': None}, ...]
    """
    results = []
    seen = set()

    # --- SANITIZACAO DE TEXTO ---
    clean_text = re.sub(r"\b\d{2}[\./]\d{2}[\./]\d{4}\b", " ", text_block)
    clean_text = re.sub(
        r"(?:Lei|Decreto|MP|LC|Art\.|Inciso)\s*n?o?\s*[\d\./-]+",
        " ", clean_text, flags=re.IGNORECASE,
    )
    clean_text = re.sub(r"\b20[1-3]\d\b", " ", clean_text)
    clean_text = re.sub(
        r"(\b\d{2,4}(?:\.\d{2})*\.)\s+(\d{2}(?:\.\d{2})*\b)", r"\1\2", clean_text,
    )

    # --- Passo 1: NCMs com EX explicitos ---
    ex_pattern = re.compile(r"(\d{2,4}(?:\.\d{2}){1,3})\s*Ex\s*(\d{1,2})", re.IGNORECASE)
    for match in ex_pattern.finditer(clean_text):
        ncm_raw = match.group(1).replace(".", "")
        ex_num = match.group(2).zfill(2)
        ncm = _normalize_ncm(ncm_raw)
        if ncm is None:
            continue
        key = (ncm, ex_num)
        if key not in seen:
            seen.add(key)
            results.append({"ncm": ncm, "ex": ex_num})

    # --- Passo 2: Remover trechos 'exceto ...' ---
    clean_block = re.sub(r"exceto.*?(?:;|\s-\s|\n|$)", " ", clean_text, flags=re.IGNORECASE)
    clean_block = ex_pattern.sub("", clean_block)

    # --- Passo 3: Intervalos "XX.XX a YY.YY" ---
    ranges = re.findall(r"(\d{2}\.\d{2})\s+a\s+(\d{2}\.\d{2})", clean_block)
    for start, end in ranges:
        try:
            for ncm in expand_range(start, end):
                key = (ncm, None)
                if key not in seen:
                    seen.add(key)
                    results.append({"ncm": ncm, "ex": None})
            clean_block = clean_block.replace(f"{start} a {end}", "")
        except Exception:
            pass

    # --- Passo 4: NCMs isolados (sem EX) ---
    singles = re.findall(r"\b(?:\d{2}|\d{4})(?:\.(?:\d{2}|\d{4}))+\b", clean_block)
    for s in singles:
        ncm = _normalize_ncm(s.replace(".", ""))
        if ncm is None:
            continue
        key = (ncm, None)
        if key not in seen:
            seen.add(key)
            results.append({"ncm": ncm, "ex": None})

    return results


def _normalize_ncm(ncm_raw):
    """Normaliza NCM para 4 ou 8 digitos. Retorna None se invalido."""
    if len(ncm_raw) == 4 or len(ncm_raw) == 8:
        return ncm_raw
    elif 4 < len(ncm_raw) < 8:
        return ncm_raw.ljust(8, "0")
    return None