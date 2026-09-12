"""
Scraper da tabela CEST do CONFAZ (Convenio ICMS 52/17).
"""
import os
import re
import json
import urllib.request
from datetime import datetime

from bs4 import BeautifulSoup

from ..config import console, CONFAZ_URL, CACHE_MAX_AGE


def download_confaz_cest():
    """
    Extrai a tabela CEST do portal do CONFAZ.
    Retorna dict: { 'ncm_limpo': [ {'cest': '0300100', 'descricao': '...'} ] }
    Usa cache local de 30 dias.
    """
    cache_path = "confaz_cest_cache.json"

    if os.path.exists(cache_path):
        cache_age = datetime.now().timestamp() - os.path.getmtime(cache_path)
        if cache_age < CACHE_MAX_AGE:
            console.print("[dim]  Usando cache local da tabela CEST do CONFAZ...[/dim]")
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)

    try:
        req = urllib.request.Request(CONFAZ_URL, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        html = urllib.request.urlopen(req, timeout=30).read().decode("utf-8")
        soup = BeautifulSoup(html, "html.parser")

        cest_db = {}
        tables = soup.find_all("table")

        for table in tables:
            rows = table.find_all("tr")
            if len(rows) < 2:
                continue

            header_cells = [c.get_text(strip=True).upper() for c in rows[0].find_all(["th", "td"])]
            if "CEST" not in header_cells or "NCM/SH" not in header_cells:
                continue

            idx_cest = header_cells.index("CEST")
            idx_ncm = header_cells.index("NCM/SH")
            idx_desc = None
            for desc_label in ["DESCRI\u00c7\u00c3O", "DESCRICAO"]:
                if desc_label in header_cells:
                    idx_desc = header_cells.index(desc_label)
                    break
            if idx_desc is None:
                idx_desc = len(header_cells) - 1

            for row in rows[1:]:
                cells = row.find_all(["th", "td"])
                if len(cells) <= max(idx_cest, idx_ncm, idx_desc):
                    continue

                cest_raw = cells[idx_cest].get_text(strip=True)
                ncm_raw = cells[idx_ncm].get_text(strip=True)
                desc = cells[idx_desc].get_text(strip=True) if idx_desc < len(cells) else ""

                if not re.match(r"\d{2}\.\d{3}\.\d{2}", cest_raw):
                    continue

                cest_limpo = cest_raw.replace(".", "")
                ncm_candidates = re.findall(r"\d{4}(?:\.\d{2}(?:\.\d{2})?)?", ncm_raw)
                for ncm_str in ncm_candidates:
                    ncm_clean = ncm_str.replace(".", "")
                    if len(ncm_clean) == 4 or len(ncm_clean) == 8:
                        pass
                    elif 4 < len(ncm_clean) < 8:
                        ncm_clean = ncm_clean.ljust(8, "0")
                    else:
                        continue

                    if ncm_clean not in cest_db:
                        cest_db[ncm_clean] = []
                    if not any(e["cest"] == cest_limpo for e in cest_db[ncm_clean]):
                        cest_db[ncm_clean].append({"cest": cest_limpo, "descricao": desc})

        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(cest_db, f, indent=2, ensure_ascii=False)

        console.print(f"[green]+[/green] Tabela CEST extraida: [bold]{len(cest_db)}[/bold] NCMs mapeados.\n")
        return cest_db

    except Exception as e:
        console.print(f"[red]x[/red] Erro ao baixar tabela CEST: {e}.")
        if os.path.exists(cache_path):
            console.print("[yellow]![/yellow] Usando cache antigo como fallback...\n")
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
        console.print("[red]x[/red] Nenhum cache disponivel.\n")
        return {}