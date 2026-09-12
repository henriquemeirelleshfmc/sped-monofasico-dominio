"""
Scraper/Downloader da TIPI (Tabela de Incidencia do IPI).
"""
import os
import re
import io
import json
import urllib.request
from datetime import datetime

import pandas as pd

from ..config import console, TIPI_URL, CACHE_MAX_AGE


def download_tipi_excel():
    """
    Baixa a TIPI em Excel, le com pandas e monta dicionario:
    { '21069010': { '02': 'Descricao do EX...' } }
    Usa cache local de 30 dias.
    """
    cache_path = "tipi_cache.json"

    if os.path.exists(cache_path):
        cache_age = datetime.now().timestamp() - os.path.getmtime(cache_path)
        if cache_age < CACHE_MAX_AGE:
            console.print("[dim]  Usando cache local da tabela TIPI...[/dim]")
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)

    tipi_db = {}
    try:
        req = urllib.request.Request(TIPI_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as response:
            excel_data = response.read()

        df = pd.read_excel(io.BytesIO(excel_data), dtype=str, engine="openpyxl")

        for _, row in df.iterrows():
            ncm_raw = str(row.iloc[0]).strip()
            ex_raw = str(row.iloc[1]).strip()
            desc = str(row.iloc[2]).strip()

            ncm_clean = ncm_raw.replace(".", "")
            if not ncm_clean.isdigit() or len(ncm_clean) != 8:
                continue

            if pd.notna(ex_raw) and ex_raw.lower() != "nan":
                ex_match = re.search(r"(\d{1,2})", ex_raw)
                if ex_match:
                    ex_clean = ex_match.group(1).zfill(2)
                    if ncm_clean not in tipi_db:
                        tipi_db[ncm_clean] = {}
                    tipi_db[ncm_clean][ex_clean] = desc

        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(tipi_db, f, indent=2, ensure_ascii=False)

        console.print(f"[green]+[/green] Tabela TIPI: [bold]{len(tipi_db)}[/bold] NCMs com Excecoes.\n")
        return tipi_db

    except Exception as e:
        console.print(f"[red]x[/red] Erro ao baixar/processar TIPI: {e}")
        if os.path.exists(cache_path):
            console.print("[yellow]![/yellow] Usando cache antigo da TIPI...\n")
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
        console.print("[red]x[/red] Nenhum cache disponivel.\n")
        return {}