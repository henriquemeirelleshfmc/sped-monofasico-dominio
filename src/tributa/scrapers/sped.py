"""
Scraper da Tabela SPED 4.3.10 do portal da Receita Federal.
"""
import os
import re
import urllib.request

from bs4 import BeautifulSoup

from ..config import console, SPED_URL


def download_latest_sped_table():
    """
    Baixa a versao mais recente da Tabela 4.3.10 do portal do SPED.
    Retorna nome do arquivo .doc baixado, ou None em caso de falha.
    """
    try:
        req = urllib.request.Request(SPED_URL, headers={"User-Agent": "Mozilla/5.0"})
        html = urllib.request.urlopen(req).read().decode("utf-8")
        soup = BeautifulSoup(html, "html.parser")

        target_id = None
        for a in soup.find_all("a"):
            if a.text and "4.3.10" in a.text:
                href = a.get("href", "")
                match = re.search(r"/item/show/(\d+)", href)
                if match:
                    target_id = match.group(1)
                    console.print(f"[green]+[/green] Encontrada no site: [cyan]{a.text.strip()}[/cyan]")
                    break

        if target_id:
            download_url = f"http://sped.rfb.gov.br/arquivo/download/{target_id}"
            for f in os.listdir("."):
                if f.startswith("Tabela_SPED_4.3.10") and f.endswith(".doc"):
                    try:
                        os.remove(f)
                    except OSError:
                        pass

            filename = f"Tabela_SPED_4.3.10_v{target_id}.doc"
            req_dl = urllib.request.Request(download_url, headers={"User-Agent": "Mozilla/5.0"})
            resp = urllib.request.urlopen(req_dl)
            with open(filename, "wb") as f:
                f.write(resp.read())
            console.print(f"[green]+[/green] Download concluido: [bold]{filename}[/bold]\n")
            return filename
        else:
            console.print("[yellow]![/yellow] Nao foi possivel encontrar a tabela 4.3.10 no site.\n")
            return None
    except Exception as e:
        console.print(f"[yellow]![/yellow] Erro ao baixar a tabela: {e}. Tentando arquivo local...\n")
        return None