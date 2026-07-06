# ============================================================================
# SPED Monofásico → Domínio Sistemas
# Gerador automático de arquivo de importação de NCMs e Impostos
#
# Autor : Henrique Meirelles (github.com/henriquemeirelleshfmc)
# Repo  : https://github.com/henriquemeirelleshfmc/sped-monofasico-dominio
# Tabela: 4.3.10 – Produtos Sujeitos a Alíquotas Diferenciadas (Monofásica)
# ============================================================================

import os
import re
import json
import urllib.request
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------------
#  Funções auxiliares de extração de NCMs
# ---------------------------------------------------------------------------

def expand_range(start_ncm, end_ncm):
    """Expande um intervalo de NCMs (ex: '22.01 a 22.03') em uma lista de NCMs individuais."""
    start = int(start_ncm.replace('.', ''))
    end = int(end_ncm.replace('.', ''))
    return [str(i).zfill(len(start_ncm.replace('.', ''))) for i in range(start, end + 1)]


def parse_ncms(text_block):
    """
    Extrai todos os NCMs encontrados em um bloco de texto da tabela SPED.
    
    Trata três cenários:
      - Intervalos no formato 'XX.XX a YY.YY'
      - NCMs isolados com pontos (ex: 22.02, 2710.11.59)
      - NCMs com 5-7 dígitos, completados com zeros à direita até 8 dígitos
    
    Ignora trechos que contêm a palavra 'exceto' para evitar falsos positivos.
    """
    # Remove textos de exceção que atrapalham a extração
    text_block = re.sub(r'exceto.*?(;|\.|$|\))', '', text_block, flags=re.IGNORECASE)
    ncms = set()
    
    # Busca intervalos de NCM no formato "XX.XX a YY.YY"
    ranges = re.findall(r'(\d{2}\.\d{2})\s+a\s+(\d{2}\.\d{2})', text_block)
    for start, end in ranges:
        try:
            ncms.update(expand_range(start, end))
            text_block = text_block.replace(f'{start} a {end}', '')
        except:
            pass
        
    # Busca NCMs isolados (2, 4, 6 ou 8 dígitos separados por ponto)
    singles = re.findall(r'\b(?:\d{1,4}\.){1,3}\d{1,4}\b', text_block)
    for s in singles:
        clean = s.replace('.', '')
        if len(clean) == 4 or len(clean) == 8:
            # A Domínio aceita 4 dígitos (sintético) ou 8 dígitos (analítico)
            ncms.add(clean)
        elif 4 < len(clean) < 8:
            # Preenche com zeros à direita até formar 8 dígitos (padrão Domínio)
            ncms.add(clean.ljust(8, '0'))
            
    return list(ncms)


# ---------------------------------------------------------------------------
#  Conversão de DOC para TXT (via automação COM do Microsoft Word)
# ---------------------------------------------------------------------------

def extract_txt_from_doc(doc_path):
    """
    Abre um arquivo .doc usando o Microsoft Word (COM/Win32)
    e salva como texto puro (.txt) para facilitar a leitura pelo script.
    """
    import win32com.client as win32
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        txt_path = os.path.abspath('tabela_temp.txt')
        doc.SaveAs(txt_path, FileFormat=2)  # FileFormat=2 → wdFormatText
        doc.Close()
        return txt_path
    except Exception as e:
        print(f'Erro ao converter o arquivo DOC: {e}')
        return None
    finally:
        try:
            word.Quit()
        except:
            pass


# ---------------------------------------------------------------------------
#  Download automático da Tabela 4.3.10 do portal do SPED
# ---------------------------------------------------------------------------

def download_latest_sped_table():
    """
    Acessa o portal do SPED da Receita Federal (sped.rfb.gov.br),
    localiza a versão mais recente da Tabela 4.3.10 e faz o download
    automático do arquivo .doc.
    
    Retorna o nome do arquivo baixado, ou None em caso de falha.
    Se a conexão falhar, o script segue usando um arquivo local como fallback.
    """
    print('Procurando a versão mais recente da Tabela 4.3.10 no site do SPED...')
    url = 'http://sped.rfb.gov.br/pasta/show/1616'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        # Procura o link que contém "4.3.10" no texto do <a>
        target_id = None
        for a in soup.find_all('a'):
            if a.text and '4.3.10' in a.text:
                href = a.get('href', '')
                match = re.search(r'/item/show/(\d+)', href)
                if match:
                    target_id = match.group(1)
                    print(f'Encontrada no site: {a.text.strip()}')
                    break
                    
        if target_id:
            download_url = f'http://sped.rfb.gov.br/arquivo/download/{target_id}'
            print('Baixando automaticamente a tabela atualizada...')
            
            # Remove versões anteriores baixadas pelo script para não acumular lixo
            for f in os.listdir('.'):
                if f.startswith('Tabela_SPED_4.3.10') and f.endswith('.doc'):
                    try:
                        os.remove(f)
                    except:
                        pass
                    
            filename = f'Tabela_SPED_4.3.10_v{target_id}.doc'
            req_dl = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
            resp = urllib.request.urlopen(req_dl)
            with open(filename, 'wb') as f:
                f.write(resp.read())
            print(f'Download concluído: {filename}\n')
            return filename
        else:
            print('Não foi possível encontrar a tabela 4.3.10 no site. Tentando usar arquivo local...\n')
            return None
    except Exception as e:
        print(f'Erro ao baixar a tabela: {e}. Tentando usar arquivo local...\n')
        return None


# ---------------------------------------------------------------------------
#  Função principal
# ---------------------------------------------------------------------------

def main():
    """
    Fluxo principal do script:
      1. Coleta os parâmetros do usuário (vigência e regime tributário)
      2. Baixa ou localiza a tabela SPED 4.3.10
      3. Extrai os códigos e NCMs da tabela
      4. Aplica regras de validação da Domínio (conflitos sintético/analítico, duplicidades)
      5. Gera o arquivo .txt pronto para importação no ERP Domínio Sistemas
    """
    print("=" * 60)
    print(" GERADOR DE ARQUIVO DE NCMs PARA A DOMÍNIO SISTEMAS ")
    print("=" * 60)
    
    # --- Etapa 1: Coleta de parâmetros do usuário ---
    vigencia_input = input("Digite a data da vigência inicial (ex: 01/01/2026): ").strip()
    if not vigencia_input:
        vigencia_input = "01/01/2026"
        
    regime_input = input("Qual o regime da empresa? (SN = Simples Nacional, LR = Lucro Real, LP = Lucro Presumido): ").strip().upper()
    
    # Mapeia o regime para o código de Tipo de Contribuição esperado pelo Domínio
    if regime_input == 'LR':
        tipo_contrib_input = 'N'  # Não Cumulativo
    elif regime_input == 'LP':
        tipo_contrib_input = 'C'  # Cumulativo
    else:
        tipo_contrib_input = ''   # Simples Nacional ou outro → em branco
    
    # --- Etapa 2: Obtenção da tabela SPED ---
    print("\nIniciando a extração dos NCMs...")
    
    # Tenta baixar a versão mais recente direto do portal do SPED
    txt_path = 'tabela_temp.txt'
    baixado = download_latest_sped_table()
    
    if baixado:
        # Sucesso no download → converte o .doc baixado para .txt
        print(f"Convertendo tabela recém-baixada ({baixado}) para TXT...")
        txt_path = extract_txt_from_doc(baixado)
    else:
        # Fallback: procura um arquivo .doc ou .txt local que contenha "Tabela" no nome
        if not os.path.exists(txt_path):
            doc_files = [f for f in os.listdir('.') if f.lower().endswith('.doc') and 'tabela' in f.lower()]
            if doc_files:
                print(f"Encontrado arquivo DOC local: {doc_files[0]}")
                print("Convertendo DOC para TXT... (isso pode levar alguns segundos)")
                txt_path = extract_txt_from_doc(doc_files[0])
            else:
                txt_files = [f for f in os.listdir('.') if f.lower().endswith('.txt') and 'tabela' in f.lower()]
                if txt_files:
                    txt_path = txt_files[0]
                else:
                    print("ERRO: Nenhuma tabela SPED (.doc ou .txt) encontrada na pasta.")
                    return

    if not txt_path or not os.path.exists(txt_path):
        print("ERRO: Falha ao obter o arquivo de texto da tabela.")
        return

    # --- Etapa 3: Leitura e extração dos códigos e NCMs ---
    print("Lendo o conteúdo da tabela...")
    with open(txt_path, 'r', encoding='latin1', errors='replace') as f:
        text = f.read()

    # Divide o texto usando os códigos SPED (3 dígitos isolados em uma linha) como separadores
    blocks = re.split(r'(?m)^([1-9]\d{2})$', text)
    data = {}

    print("Extraindo os Códigos e NCMs...")
    for i in range(1, len(blocks), 2):
        code = blocks[i]
        block_text = blocks[i+1].strip()
        
        lines = [l.strip() for l in block_text.split('\n') if l.strip()]
        if not lines:
            continue
        desc = lines[0]  # Primeira linha após o código é a descrição do produto
        
        ncms = parse_ncms(block_text)
        if ncms:
            if code not in data:
                data[code] = {'descricao': desc, 'ncms': set()}
            data[code]['ncms'].update(ncms)

    # Ordena os NCMs dentro de cada código para manter consistência
    for code in data:
        data[code]['ncms'] = sorted(list(data[code]['ncms']))

    # Salva uma cópia em JSON para referência e debug
    json_path = 'tabela_sped_ncms.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"JSON estruturado gerado em: {json_path}")

    # --- Etapa 4: Regras de validação da Domínio Sistemas ---
    # Regra 1: Não importar NCM analítico (8 dígitos) se o sintético (4 dígitos) já existe
    # Regra 2: Não duplicar o mesmo NCM em códigos de configuração diferentes
    synthetic_ncms = set()
    for code, info in data.items():
        for ncm in info['ncms']:
            if len(ncm) == 4:
                synthetic_ncms.add(ncm)

    seen_ncms = set()
    for code in sorted(data.keys(), key=lambda x: int(x)):
        valid_ncms = []
        for ncm in data[code]['ncms']:
            if len(ncm) == 8 and ncm[:4] in synthetic_ncms:
                continue  # Regra 1: Sintético já cobre este NCM
            if ncm in seen_ncms:
                continue  # Regra 2: NCM já foi usado em outro código
            seen_ncms.add(ncm)
            valid_ncms.append(ncm)
        data[code]['ncms'] = valid_ncms

    # --- Etapa 5: Geração do arquivo TXT para importação ---
    print("Gerando o arquivo de texto para a Domínio...")
    csv_lines = [f'Vigencia;{vigencia_input};Vigencia_Inicial']

    for code, info in sorted(data.items(), key=lambda x: int(x[0])):
        # Código 400 foi desativado pela Receita Federal (válido até 30.04.2015)
        # Os demais da série 4xx (401, 402, 411, 415, etc.) continuam ativos
        if int(code) == 400:
            continue
            
        desc = info['descricao'].replace(';', ',').replace('\n', ' ')
        ncms = info['ncms']

        if ncms:
            # Registro "Item": Descrição com código SPED no início para facilitar a visualização
            desc_com_codigo = f"{code} - {desc}"
            csv_lines.append(f'Item;{desc_com_codigo};{code}')
            
            # Registro "Impostos": Linha com todos os campos tributários
            # Layout: Tipo; TipoContrib; CSTEntrada; VincCredito; BaseCredito; AliqPISEnt;
            #         AliqCOFINSEnt; CSTSaida; NatReceita; AliqPISSaida; AliqCOFINSSaida;
            #         CSTICMSEnt; CSTICMSSaida; AliqICMS; CSTIPIEnt; CSTIPISaida; AliqIPI;
            #         SujeitoPISCOFINS; TipoTribPISCOFINS
            impostos_row = [
                "Impostos",         #  1: Identificador do registro
                tipo_contrib_input, #  2: Tipo de Contribuição (N=Não Cumulativo, C=Cumulativo, vazio=SN)
                "70",               #  3: CST PIS/COFINS Entrada (70 = Sem Direito a Crédito)
                "01",               #  4: Vínculo de Crédito
                "1",                #  5: Base de Crédito
                "0",                #  6: Alíquota PIS Entrada
                "0",                #  7: Alíquota COFINS Entrada
                "04",               #  8: CST PIS/COFINS Saída (04 = Monofásica)
                str(code),          #  9: Natureza da Receita (= Código SPED do produto)
                "0",                # 10: Alíquota PIS Saída
                "0",                # 11: Alíquota COFINS Saída
                "",                 # 12: CST ICMS Entrada (em branco)
                "",                 # 13: CST ICMS Saída (em branco)
                "",                 # 14: Alíquota ICMS (em branco)
                "",                 # 15: CST IPI Entrada (em branco)
                "",                 # 16: CST IPI Saída (em branco)
                "",                 # 17: Alíquota IPI (em branco)
                "S",                # 18: Produto Sujeito a PIS/COFINS (S=Sim)
                ""                  # 19: Tipo de Tributação PIS/COFINS (em branco)
            ]
            csv_lines.append(";".join(impostos_row))
            
            # Registros "NCM": Um por linha, vinculados ao Item acima
            for ncm in ncms:
                csv_lines.append(f'NCM;{ncm};')

    # Salva o arquivo final em encoding latin1 (padrão da Domínio)
    output_path = 'gerardadosporncm_COMPLETO.txt'
    with open(output_path, 'w', encoding='latin1', errors='ignore') as f:
        f.write('\n'.join(csv_lines) + '\n')
    
    print(f"Arquivo gerado com sucesso em: {output_path}")


# ---------------------------------------------------------------------------
#  Ponto de entrada
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    main()
