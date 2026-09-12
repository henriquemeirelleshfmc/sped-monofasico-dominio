"""
Tributa NF-e - Ponto de entrada principal.
Orquestra scrapers, parser, matching e exporter.
"""
import os
import re
import json
from datetime import datetime

from .config import console, REGIME_MAP, Panel, Table, Prompt, Progress, SpinnerColumn, TextColumn, box
from .parsers.ncm import parse_ncms
from .parsers.doc import extract_txt_from_doc
from .scrapers.sped import download_latest_sped_table
from .scrapers.confaz import download_confaz_cest
from .scrapers.tipi import download_tipi_excel
from .matching.fuzzy import find_best_cest, load_manual_overrides
from .exporters.dominio import gerar_txt_dominio


def main():
    """Fluxo principal do Tributa NF-e."""
    # --- Banner ---
    console.print()
    console.print(Panel.fit(
        "[bold bright_cyan]TRIBUTA NF-e[/bold bright_cyan]\n"
        "[dim]Gerador Automatico de NCMs para Dominio Sistemas[/dim]\n"
        "[dim]Tabela SPED 4.3.10 - Monofasica (Bebidas Frias)[/dim]",
        border_style="bright_blue", padding=(1, 4),
    ))
    console.print()

    # --- Etapa 1: Parametros ---
    vigencia_input = Prompt.ask("[bold cyan]Data da vigencia inicial[/bold cyan]", default="01/01/2026")
    regime_input = Prompt.ask("[bold cyan]Regime da empresa[/bold cyan]", choices=["SN", "LR", "LP"], default="LR").upper()

    tipo_contrib, regime_nome = REGIME_MAP.get(regime_input, ("", regime_input))
    console.print(f"\n[dim]  Regime selecionado: [bold]{regime_nome}[/bold][/dim]")

    # --- Etapa 2: Tabela SPED ---
    console.print()
    console.rule("[bold bright_cyan]Etapa 1 - Obtencao da Tabela SPED[/bold bright_cyan]")

    with Progress(SpinnerColumn("dots"), TextColumn("[progress.description]{task.description}"), console=console, transient=True) as progress:
        task = progress.add_task("Procurando tabela 4.3.10 no portal do SPED...", total=None)
        txt_path = "tabela_temp.txt"
        baixado = download_latest_sped_table()
        progress.remove_task(task)

    if baixado:
        with Progress(SpinnerColumn("dots"), TextColumn("[progress.description]{task.description}"), console=console, transient=True) as progress:
            task = progress.add_task(f"Convertendo {baixado} para TXT...", total=None)
            txt_path = extract_txt_from_doc(baixado)
            progress.remove_task(task)
    else:
        if not os.path.exists(txt_path):
            doc_files = [f for f in os.listdir(".") if f.lower().endswith(".doc") and "tabela" in f.lower()]
            if doc_files:
                console.print(f"[dim]  Encontrado DOC local: {doc_files[0]}[/dim]")
                with Progress(SpinnerColumn("dots"), TextColumn("[progress.description]{task.description}"), console=console, transient=True) as progress:
                    task = progress.add_task("Convertendo DOC para TXT...", total=None)
                    txt_path = extract_txt_from_doc(doc_files[0])
                    progress.remove_task(task)
            else:
                txt_files = [f for f in os.listdir(".") if f.lower().endswith(".txt") and "tabela" in f.lower()]
                if txt_files:
                    txt_path = txt_files[0]
                else:
                    console.print("[red]x ERRO: Nenhuma tabela SPED encontrada na pasta.[/red]")
                    return

    if not txt_path or not os.path.exists(txt_path):
        console.print("[red]x ERRO: Falha ao obter o arquivo de texto.[/red]")
        return

    # --- Etapa 3: Extracao ---
    console.print()
    console.rule("[bold bright_cyan]Etapa 2 - Extracao de NCMs[/bold bright_cyan]")

    for enc in ["utf-8", "latin1"]:
        try:
            with open(txt_path, "r", encoding=enc) as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        with open(txt_path, "r", encoding="latin1", errors="replace") as f:
            text = f.read()

    blocks = re.split(r"(?m)^([1-9]\d{2})$", text)
    data = {}

    for i in range(1, len(blocks), 2):
        code = blocks[i]
        block_text = blocks[i + 1].strip()
        lines = [l.strip() for l in block_text.split("\n") if l.strip()]
        if not lines:
            continue
        desc = lines[0]

        ncm_entries = parse_ncms(block_text)
        if ncm_entries:
            if code not in data:
                data[code] = {"descricao": desc, "ncms": []}
            existing_keys = {(e["ncm"], e["ex"]) for e in data[code]["ncms"]}
            for entry in ncm_entries:
                key = (entry["ncm"], entry["ex"])
                if key not in existing_keys:
                    existing_keys.add(key)
                    data[code]["ncms"].append(entry)

    for code in data:
        data[code]["ncms"] = sorted(data[code]["ncms"], key=lambda x: (x["ncm"], x["ex"] or ""))

    json_path = "tabela_sped_ncms.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    console.print(f"[green]+[/green] {len(data)} Naturezas da Receita extraidas da tabela SPED")

    # --- Etapa 4: Desempate ---
    console.print()
    console.rule("[bold bright_cyan]Etapa 3 - Desempate e Validacao[/bold bright_cyan]")

    ncm_to_codes = {}
    for code, info in data.items():
        for entry in info["ncms"]:
            ncm_to_codes.setdefault(entry["ncm"], set()).add(code)
    duplicated_ncms = {ncm for ncm, codes in ncm_to_codes.items() if len(codes) > 1}

    manual_map = load_manual_overrides()
    cest_db = {}
    tipi_db = {}
    if duplicated_ncms:
        console.print(f"[yellow]![/yellow] {len(duplicated_ncms)} NCMs duplicados encontrados. Iniciando desempate...")
        cest_db = download_confaz_cest()
        tipi_db = download_tipi_excel()
    else:
        console.print("[green]+[/green] Nenhum NCM duplicado. Desempate por CEST nao e necessario.")

    log_path = "ncm_sem_cest_revisao.log"
    log_entries = []
    seen_keys = set()

    for code in sorted(data.keys(), key=lambda x: int(x)):
        valid_entries = []
        for entry in data[code]["ncms"]:
            ncm, ex = entry["ncm"], entry["ex"]

            synthetic_in_code = {e["ncm"] for e in data[code]["ncms"] if len(e["ncm"]) == 4}
            if len(ncm) == 8 and ncm[:4] in synthetic_in_code:
                continue

            if ncm not in duplicated_ncms:
                key = (ncm, "")
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({"ncm": ncm, "ex": ex, "cest": ""})
                continue

            if ex is not None:
                tipi_desc = tipi_db.get(ncm, {}).get(ex)
                if cest_db:
                    sped_desc = data[code]["descricao"]
                    desc_combinada = f"{sped_desc} | Excecao TIPI: {tipi_desc}" if tipi_desc else sped_desc
                    best_cest, score = find_best_cest(desc_combinada, ncm, cest_db, manual_map)
                    if score >= 0.40:
                        key = (ncm, best_cest)
                        if key not in seen_keys:
                            seen_keys.add(key)
                            valid_entries.append({"ncm": ncm, "ex": ex, "cest": best_cest})
                        continue
                key = (ncm, "")
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({"ncm": ncm, "ex": ex, "cest": ""})
                    log_entries.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Nat.Receita={code} | NCM={ncm} EX {ex} | CEST nao resolvido.")
                continue

            if cest_db:
                sped_desc = data[code]["descricao"]
                best_cest, score = find_best_cest(sped_desc, ncm, cest_db, manual_map)
                if score >= 0.35:
                    key = (ncm, best_cest)
                    if key not in seen_keys:
                        seen_keys.add(key)
                        valid_entries.append({"ncm": ncm, "ex": ex, "cest": best_cest})
                else:
                    key = (ncm, "")
                    if key not in seen_keys:
                        seen_keys.add(key)
                        valid_entries.append({"ncm": ncm, "ex": ex, "cest": ""})
                        log_entries.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Nat.Receita={code} | NCM={ncm} | Score={score:.2f} | Desc: {sped_desc[:80]}")
            else:
                key = (ncm, "")
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({"ncm": ncm, "ex": ex, "cest": ""})
                    log_entries.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Nat.Receita={code} | NCM={ncm} | CEST nao encontrado | Desc: {data[code]['descricao'][:80]}")

        data[code]["ncms_final"] = valid_entries

    if log_entries:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("=" * 70 + "\n  NCMs SEM CEST - REVISAO MANUAL NECESSARIA\n" + "=" * 70 + "\n\n")
            f.write("\n".join(log_entries) + "\n")

    # --- Etapa 5: Geracao do arquivo ---
    console.print()
    console.rule("[bold bright_cyan]Etapa 4 - Geracao do Arquivo[/bold bright_cyan]")

    total_ncms = gerar_txt_dominio(data, vigencia_input, tipo_contrib)

    # --- Resumo ---
    total_items = sum(1 for info in data.values() if info.get("ncms_final"))
    console.print()
    summary = Table(title="Resumo da Exportacao", box=box.ROUNDED, title_style="bold bright_cyan", show_header=True, header_style="bold")
    summary.add_column("Metrica", style="cyan", min_width=30)
    summary.add_column("Valor", style="bold white", justify="right", min_width=10)
    summary.add_row("Naturezas de Receita", str(total_items))
    summary.add_row("NCMs Exportados", str(total_ncms))
    summary.add_row("Vigencia Inicial", vigencia_input)
    summary.add_row("Regime Tributario", regime_nome)
    if log_entries:
        summary.add_row("NCMs para Revisao", f"[yellow]{len(log_entries)}[/yellow]")
        summary.add_row("Arquivo de Log", f"[dim]{log_path}[/dim]")
    else:
        summary.add_row("NCMs para Revisao", "[green]0 - Tudo resolvido![/green]")
    console.print(summary)
    console.print()