"""
Exportador de arquivo TXT para importacao no ERP Dominio Sistemas.
"""
from ..config import console


def gerar_txt_dominio(data, vigencia, tipo_contrib, output_path="gerardadosporncm_COMPLETO.txt"):
    """
    Gera o arquivo TXT no layout esperado pela Dominio Sistemas.
    Retorna o numero total de NCMs exportados.
    """
    csv_lines = [f"Vigencia;{vigencia};Vigencia_Inicial"]

    for code, info in sorted(data.items(), key=lambda x: int(x[0])):
        if int(code) == 400:
            continue

        ncm_entries = info.get("ncms_final", [])
        if not ncm_entries:
            continue

        desc = info["descricao"].replace(";", ",").replace("\n", " ")
        desc_com_codigo = f"{code} - {desc}"
        csv_lines.append(f"Item;{desc_com_codigo};{code}")

        impostos_row = [
            "Impostos", tipo_contrib,
            "70", "01", "1", "0", "0",
            "04", str(code), "0", "0",
            "", "", "", "", "", "", "S", "",
        ]
        csv_lines.append(";".join(impostos_row))

        for ncm_dict in ncm_entries:
            cest = ncm_dict.get("cest", "")
            csv_lines.append(f'NCM;{ncm_dict["ncm"]};{cest};')

    with open(output_path, "w", encoding="latin1", errors="ignore") as f:
        f.write("\n".join(csv_lines) + "\n")

    console.print(f"[green]+[/green] Arquivo gerado: [bold]{output_path}[/bold]")

    total_ncms = sum(len(info.get("ncms_final", [])) for info in data.values())
    return total_ncms