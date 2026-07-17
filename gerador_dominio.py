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
import difflib
import urllib.request
import io
import pandas as pd
from datetime import datetime
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
    results = []
    seen = set()

    # --- SANITIZAÇÃO DE TEXTO ---
    # 1. Remove datas completas que podem ser confundidas com NCMs
    clean_text = re.sub(r'\b\d{2}[\./]\d{2}[\./]\d{4}\b', ' ', text_block)
    # 2. Remove números de Leis, Decretos, Artigos
    clean_text = re.sub(r'(?:Lei|Decreto|MP|LC|Art\.|Inciso)\s*n?o?\s*[\d\./-]+', ' ', clean_text, flags=re.IGNORECASE)
    # 3. Remove anos isolados (ex: 2013, 2015, 2026)
    clean_text = re.sub(r'\b20[1-3]\d\b', ' ', clean_text)
    # 4. Reconectar NCMs quebrados por quebra de linha na tabela do Word (ex: "22.\n01.10.00" -> "22.01.10.00")
    clean_text = re.sub(r'(\b\d{2,4}(?:\.\d{2})*\.)\s+(\d{2}(?:\.\d{2})*\b)', r'\1\2', clean_text)

    # --- Passo 1: Extrair NCMs com EX explícitos ---
    ex_pattern = re.compile(
        r'(\d{2,4}(?:\.\d{2}){1,3})\s*Ex\s*(\d{1,2})',
        re.IGNORECASE
    )
    for match in ex_pattern.finditer(clean_text):
        ncm_raw = match.group(1).replace('.', '')
        ex_num = match.group(2).zfill(2)
        if len(ncm_raw) == 4 or len(ncm_raw) == 8:
            ncm = ncm_raw
        elif 4 < len(ncm_raw) < 8:
            ncm = ncm_raw.ljust(8, '0')
        else:
            continue
        key = (ncm, ex_num)
        if key not in seen:
            seen.add(key)
            results.append({'ncm': ncm, 'ex': ex_num})

    # --- Passo 2: Remover trechos de exceção ('exceto ...') ---
    # Utiliza delimitadores seguros (ponto e vírgula, hífen de separação, ou quebra de linha) em vez de parênteses.
    clean_block = re.sub(r'exceto.*?(?:;|\s-\s|\n|$)', ' ', clean_text, flags=re.IGNORECASE)
    clean_block = ex_pattern.sub('', clean_block)

    # --- Passo 3: Buscar intervalos de NCM "XX.XX a YY.YY" ---
    ranges = re.findall(r'(\d{2}\.\d{2})\s+a\s+(\d{2}\.\d{2})', clean_block)
    for start, end in ranges:
        try:
            expanded = expand_range(start, end)
            for ncm in expanded:
                key = (ncm, None)
                if key not in seen:
                    seen.add(key)
                    results.append({'ncm': ncm, 'ex': None})
            clean_block = clean_block.replace(f'{start} a {end}', '')
        except:
            pass

    # --- Passo 4: Buscar NCMs isolados (sem EX) ---
    singles = re.findall(r'\b(?:\d{2}|\d{4})(?:\.(?:\d{2}|\d{4}))+\b', clean_block)
    for s in singles:
        ncm_raw = s.replace('.', '')
        if len(ncm_raw) == 4 or len(ncm_raw) == 8:
            ncm = ncm_raw
        elif 4 < len(ncm_raw) < 8:
            ncm = ncm_raw.ljust(8, '0')
        else:
            continue
        key = (ncm, None)
        if key not in seen:
            seen.add(key)
            results.append({'ncm': ncm, 'ex': None})

    return results


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
#  Download e scraping da tabela CEST do CONFAZ (Convênio ICMS 52/17)
# ---------------------------------------------------------------------------

def download_confaz_cest():
    """
    Acessa o portal do CONFAZ e extrai a tabela CEST (Código Especificador
    da Substituição Tributária) do Convênio ICMS 52/17.

    Retorna um dicionário no formato:
        { 'ncm_limpo': [ {'cest': '0300100', 'descricao': '...'}, ... ] }

    Usa cache local (confaz_cest_cache.json) para evitar downloads repetidos.
    O cache é renovado a cada 30 dias.
    """
    cache_path = 'confaz_cest_cache.json'

    # Verifica se o cache existe e se tem menos de 30 dias
    if os.path.exists(cache_path):
        cache_age = datetime.now().timestamp() - os.path.getmtime(cache_path)
        if cache_age < 30 * 24 * 3600:  # 30 dias em segundos
            print('Usando cache local da tabela CEST do CONFAZ...')
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)

    print('Baixando tabela CEST do CONFAZ (Convênio ICMS 52/17)...')
    url = 'https://www.confaz.fazenda.gov.br/legislacao/convenios/2017/CV052_17'
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        html = urllib.request.urlopen(req, timeout=30).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')

        cest_db = {}  # ncm_limpo -> [{ cest, descricao }]
        tables = soup.find_all('table')

        for table in tables:
            rows = table.find_all('tr')
            if len(rows) < 2:
                continue

            # Verifica se a tabela tem as colunas CEST e NCM/SH
            header_cells = [c.get_text(strip=True).upper() for c in rows[0].find_all(['th', 'td'])]
            if 'CEST' not in header_cells or 'NCM/SH' not in header_cells:
                continue

            # Identifica os índices das colunas
            idx_cest = header_cells.index('CEST')
            idx_ncm = header_cells.index('NCM/SH')
            idx_desc = None
            for desc_label in ['DESCRIÇÃO', 'DESCRICAO', 'DESCRI\xc7\xc3O']:
                if desc_label in header_cells:
                    idx_desc = header_cells.index(desc_label)
                    break
            if idx_desc is None:
                # Tenta pegar a última coluna como descrição
                idx_desc = len(header_cells) - 1

            for row in rows[1:]:
                cells = row.find_all(['th', 'td'])
                if len(cells) <= max(idx_cest, idx_ncm, idx_desc):
                    continue

                cest_raw = cells[idx_cest].get_text(strip=True)
                ncm_raw = cells[idx_ncm].get_text(strip=True)
                desc = cells[idx_desc].get_text(strip=True) if idx_desc < len(cells) else ''

                # Validação: CEST deve ter formato XX.XXX.XX
                if not re.match(r'\d{2}\.\d{3}\.\d{2}', cest_raw):
                    continue

                cest_limpo = cest_raw.replace('.', '')

                # O campo NCM/SH pode conter múltiplos NCMs concatenados ou separados
                # Extrai todos os NCMs de 4 ou 8 dígitos
                ncm_candidates = re.findall(r'\d{4}(?:\.\d{2}(?:\.\d{2})?)?', ncm_raw)
                for ncm_str in ncm_candidates:
                    ncm_clean = ncm_str.replace('.', '')
                    if len(ncm_clean) == 4 or len(ncm_clean) == 8:
                        pass  # Já está no formato correto
                    elif 4 < len(ncm_clean) < 8:
                        ncm_clean = ncm_clean.ljust(8, '0')
                    else:
                        continue

                    if ncm_clean not in cest_db:
                        cest_db[ncm_clean] = []
                    # Evitar CESTs duplicados para o mesmo NCM
                    if not any(e['cest'] == cest_limpo for e in cest_db[ncm_clean]):
                        cest_db[ncm_clean].append({
                            'cest': cest_limpo,
                            'descricao': desc
                        })

        # Salva em cache local
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cest_db, f, indent=2, ensure_ascii=False)

        print(f'Tabela CEST extraída: {len(cest_db)} NCMs mapeados com CEST.\n')
        return cest_db

    except Exception as e:
        print(f'Erro ao baixar tabela CEST: {e}.')
        # Tenta usar cache antigo como fallback
        if os.path.exists(cache_path):
            print('Usando cache antigo da tabela CEST como fallback...\n')
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        print('Nenhum cache disponível. Prosseguindo sem CEST (desempate desabilitado).\n')
        return {}


# ---------------------------------------------------------------------------
#  Scraper/Downloader da TIPI (Tabela de Incidência do IPI)
# ---------------------------------------------------------------------------

def download_tipi_excel():
    """
    Baixa a Tabela TIPI em Excel do portal do governo (ou usa cache),
    lê com pandas e monta um dicionário estruturado:
    { '21069010': { '02': 'Descrição detalhada do EX...' } }
    """
    cache_path = 'tipi_cache.json'
    
    # Verifica cache (válido por 30 dias)
    if os.path.exists(cache_path):
        cache_age = datetime.now().timestamp() - os.path.getmtime(cache_path)
        if cache_age < 30 * 24 * 3600:
            print('Usando cache local da tabela TIPI...')
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)

    print('Baixando tabela TIPI (Excel) da Receita Federal...')
    # URL de exemplo (pode ser ajustada para o link direto mais recente do Gov.br)
    url = 'https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/legislacao/documentos-e-arquivos/tipi.xlsx'
    
    tipi_db = {}
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=30) as response:
            excel_data = response.read()
            
        print('Processando arquivo Excel com pandas...')
        # Dependendo da formatação da RFB, lemos a aba principal.
        # Geralmente as colunas de interesse caem no índice 0 (NCM), 1 (EX), 2 (Descrição)
        df = pd.read_excel(io.BytesIO(excel_data), dtype=str, engine='openpyxl')
        
        # Iterar sobre as linhas e extrair NCM, EX e Descrição
        for index, row in df.iterrows():
            ncm_raw = str(row.iloc[0]).strip()
            ex_raw = str(row.iloc[1]).strip()
            desc = str(row.iloc[2]).strip()
            
            # Limpeza do NCM
            ncm_clean = ncm_raw.replace('.', '')
            if not ncm_clean.isdigit() or len(ncm_clean) != 8:
                continue
                
            # Limpeza do EX (extrair apenas os dígitos)
            if pd.notna(ex_raw) and ex_raw.lower() != 'nan':
                ex_match = re.search(r'(\d{1,2})', ex_raw)
                if ex_match:
                    ex_clean = ex_match.group(1).zfill(2)
                    
                    if ncm_clean not in tipi_db:
                        tipi_db[ncm_clean] = {}
                    tipi_db[ncm_clean][ex_clean] = desc

        # Salva o cache
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(tipi_db, f, indent=2, ensure_ascii=False)
            
        print(f'Tabela TIPI processada: {len(tipi_db)} NCMs com Exceções mapeados.\n')
        return tipi_db

    except Exception as e:
        print(f'Erro ao baixar/processar TIPI: {e}')
        if os.path.exists(cache_path):
            print('Tentando usar cache antigo da TIPI como fallback...\n')
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        print('Nenhum cache disponível. O desempate por EX usará fallback manual.\n')
        return {}


# ---------------------------------------------------------------------------
#  Fuzzy matching de descrições para desempate por CEST
# ---------------------------------------------------------------------------

def find_best_cest(sped_desc, ncm, cest_db):
    """
    Compara a descrição de uma Natureza da Receita (SPED) com as descrições
    dos registros CEST do CONFAZ.
    
    Busca por NCM exato ou prefixo, e faz fallback para busca global.
    Usa combinação de difflib e intersecção de palavras-chave.

    Retorna uma tupla (cest_code, score) onde:
      - cest_code: o código CEST mais próximo (str) ou '' se score < 0.35
      - score: float de 0 a 1 indicando a similaridade
    """
    # Override manual para casos onde a semântica do SPED/TIPI diverge do CONFAZ
    # ou onde o SPED utiliza NCMs desatualizados.
    MANUAL_CEST_MAP = {
        '2203': '0302100',          # Cervejas de malte
        '22071000': '0600101',      # Álcool etílico hidratado combustível
        '22072010': '0600100',      # Álcool etílico anidro combustível
        '21069010': '0301000',      # Preparações compostas (Xaropes para refrigerantes)
        '22011000': '0300100',      # Águas minerais (garante o CEST padrão)
        '22029000': '0301100',      # Refrescos / Energéticos (NCM desatualizado no SPED)
        '22029900': '0301100',      # Refrescos / Energéticos (NCM atualizado)
    }
    
    if ncm in MANUAL_CEST_MAP:
        return MANUAL_CEST_MAP[ncm], 1.0

    stopwords = {'de', 'do', 'da', 'dos', 'das', 'e', 'ou', 'em', 'com', 'a', 'o', 'as', 'os',
                 'para', 'por', 'no', 'na', 'nos', 'nas', 'um', 'uma', 'que', 'se', 'ao', 'ate',
                 'nao', 'nem', 'mais', 'outros', 'outras', 'incluindo', 'inclusive', 'exceto'}
    
    def tokens(text):
        words = set(re.findall(r'[a-z\u00e1\u00e0\u00e2\u00e3\u00e9\u00e8\u00ea\u00ed\u00ef\u00f3\u00f4\u00f5\u00fa\u00fc\u00e7]+', text.lower()))
        return words - stopwords
    
    def extract_vol(text):
        vols = re.findall(r'(\d+)\s*(?:ml|litros|litro|l)\b', text.lower())
        return set(vols)
    
    def combined_score(desc1, desc2):
        fuzzy = difflib.SequenceMatcher(None, desc1.lower(), desc2.lower()).ratio()
        t1, t2 = tokens(desc1), tokens(desc2)
        kw = len(t1 & t2) / min(len(t1), len(t2)) if t1 and t2 else 0.0
        
        v1 = extract_vol(desc1)
        v2 = extract_vol(desc2)
        vol_score = 0.0
        if v1 and v2:
            if v1 & v2:
                vol_score = 1.0  # Bônus enorme se o volume bater
            else:
                vol_score = -0.5 # Penalidade se tiverem volumes diferentes
                
        return max(0.0, 0.4 * fuzzy + 0.6 * kw + vol_score)

    candidates = []
    
    # 1. Busca por NCM exato ou prefixo (ex: SPED 2201 -> CONFAZ 22011000)
    for confaz_ncm, entries in cest_db.items():
        if confaz_ncm == ncm or confaz_ncm.startswith(ncm) or ncm.startswith(confaz_ncm):
            for e in entries:
                score = combined_score(sped_desc, e['descricao'])
                candidates.append((score, e['cest'], e['descricao']))
                
    # 2. Fallback: Busca global se não achar nada pelo NCM
    if not candidates:
        for confaz_ncm, entries in cest_db.items():
            for e in entries:
                score = combined_score(sped_desc, e['descricao'])
                if score >= 0.4:
                    candidates.append((score, e['cest'], e['descricao']))

    if not candidates:
        return '', 0.0

    candidates.sort(reverse=True)
    best_score, best_cest, best_desc = candidates[0]
    
    return best_cest, best_score


# ---------------------------------------------------------------------------
#  Função principal
# ---------------------------------------------------------------------------

def main():
    """
    Fluxo principal do script:
      1. Coleta os parâmetros do usuário (vigência e regime tributário)
      2. Baixa ou localiza a tabela SPED 4.3.10
      3. Extrai os códigos, NCMs e Exceções (EX) da tabela
      4. Aplica desempate em 3 níveis (NCM único → EX → CEST/CONFAZ)
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

    # --- Etapa 3: Leitura e extração dos códigos, NCMs e Exceções (EX) ---
    print("Lendo o conteúdo da tabela...")
    with open(txt_path, 'r', encoding='latin1', errors='replace') as f:
        text = f.read()

    # Divide o texto usando os códigos SPED (3 dígitos isolados em uma linha) como separadores
    blocks = re.split(r'(?m)^([1-9]\d{2})$', text)
    data = {}

    print("Extraindo os Códigos, NCMs e Exceções (EX)...")
    for i in range(1, len(blocks), 2):
        code = blocks[i]
        block_text = blocks[i+1].strip()

        lines = [l.strip() for l in block_text.split('\n') if l.strip()]
        if not lines:
            continue
        desc = lines[0]  # Primeira linha após o código é a descrição do produto

        ncm_entries = parse_ncms(block_text)
        if ncm_entries:
            if code not in data:
                data[code] = {'descricao': desc, 'ncms': []}
            # Adiciona sem duplicar (mesmo ncm+ex) dentro do mesmo código
            existing_keys = {(e['ncm'], e['ex']) for e in data[code]['ncms']}
            for entry in ncm_entries:
                key = (entry['ncm'], entry['ex'])
                if key not in existing_keys:
                    existing_keys.add(key)
                    data[code]['ncms'].append(entry)

    # Ordena os NCMs dentro de cada código para manter consistência
    for code in data:
        data[code]['ncms'] = sorted(data[code]['ncms'], key=lambda x: (x['ncm'], x['ex'] or ''))

    # Salva uma cópia em JSON para referência e debug
    json_path = 'tabela_sped_ncms.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"JSON estruturado gerado em: {json_path}")

    # --- Etapa 4: Desempate em 3 níveis + Regras de validação ---
    print("Aplicando regras de validação e desempate de NCMs duplicados...")

    # A regra foi movida para dentro do loop individual de cada código para evitar
    # apagar NCMs analíticos de códigos que não possuem o sintético.

    # 4.2: Mapear NCM → lista de códigos SPED onde aparece (para detectar duplicidades)
    ncm_to_codes = {}
    for code, info in data.items():
        for entry in info['ncms']:
            ncm = entry['ncm']
            ncm_to_codes.setdefault(ncm, set()).add(code)

    # Identificar NCMs duplicados (presentes em 2+ códigos SPED)
    duplicated_ncms = {ncm for ncm, codes in ncm_to_codes.items() if len(codes) > 1}

    # 4.3: Baixar CEST do CONFAZ somente se houver NCMs duplicados para desempatar
    cest_db = {}
    tipi_db = {}
    if duplicated_ncms:
        print(f"Encontrados {len(duplicated_ncms)} NCMs duplicados. Baixando tabela CEST para desempate...")
        cest_db = download_confaz_cest()
        tipi_db = download_tipi_excel() # <--- NOVA CHAMADA AQUI
    else:
        print("Nenhum NCM duplicado encontrado. Desempate por CEST não é necessário.")

    # 4.4: Preparar arquivo de log para NCMs sem CEST
    log_path = 'ncm_sem_cest_revisao.log'
    log_entries = []

    # 4.5: Desempate em 3 Níveis e montagem da lista final
    # Chave composta: (ncm, cest) — GLOBALMENTE ÚNICA para não dar erro na Domínio
    seen_keys = set()
    for code in sorted(data.keys(), key=lambda x: int(x)):
        valid_entries = []
        for entry in data[code]['ncms']:
            ncm = entry['ncm']
            ex = entry['ex']

            # Regra base: Não importar NCM analítico se o sintético já existe na MESMA Natureza
            synthetic_in_code = {e['ncm'] for e in data[code]['ncms'] if len(e['ncm']) == 4}
            if len(ncm) == 8 and ncm[:4] in synthetic_in_code:
                continue

            # --- Nível 1: NCM Único (sem duplicidade) ---
            if ncm not in duplicated_ncms:
                key = (ncm, '')
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': ''})
                continue

            # --- Nível 2: NCM com Exceção (EX) ---
            if ex is not None:
                tipi_desc = tipi_db.get(ncm, {}).get(ex)
                
                if cest_db:
                    sped_desc = data[code]['descricao']
                    # Se não houver TIPI (NCM extinto), usa apenas a descrição SPED
                    desc_combinada = f"{sped_desc} | Exceção TIPI: {tipi_desc}" if tipi_desc else sped_desc
                    
                    best_cest, score = find_best_cest(desc_combinada, ncm, cest_db)
                    
                    # O override manual retorna score 1.0, garantindo a aprovação aqui
                    if score >= 0.40:
                        key = (ncm, best_cest)
                        if key not in seen_keys:
                            seen_keys.add(key)
                            valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': best_cest})
                        continue
                
                # Fallback: Se realmente falhar no find_best_cest, exporta vazio e loga
                key = (ncm, '')
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': ''})
                    log_entries.append(
                        f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] "
                        f"Nat.Receita={code} | NCM={ncm} EX {ex} | CEST não resolvido."
                    )
                continue

            # --- Nível 3: Desempate por CEST (fuzzy matching com CONFAZ) ---
            if cest_db:
                sped_desc = data[code]['descricao']
                best_cest, score = find_best_cest(sped_desc, ncm, cest_db)

                if score >= 0.35:
                    key = (ncm, best_cest)
                    if key not in seen_keys:
                        seen_keys.add(key)
                        valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': best_cest})
                else:
                    # Fallback: exporta com CEST vazio + registra no log
                    # Chave igual à exportada para manter a unicidade global
                    key = (ncm, '')
                    if key not in seen_keys:
                        seen_keys.add(key)
                        valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': ''})
                        log_entries.append(
                            f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] "
                            f"Nat.Receita={code} | NCM={ncm} | Score={score:.2f} | "
                            f"Desc SPED: {sped_desc[:80]}"
                        )
            else:
                # NCM duplicado sem dados no CONFAZ → exporta com CEST vazio + log
                key = (ncm, '')
                if key not in seen_keys:
                    seen_keys.add(key)
                    valid_entries.append({'ncm': ncm, 'ex': ex, 'cest': ''})
                    log_entries.append(
                        f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] "
                        f"Nat.Receita={code} | NCM={ncm} | CEST não encontrado no CONFAZ | "
                        f"Desc: {data[code]['descricao'][:80]}"
                    )

        data[code]['ncms_final'] = valid_entries

    # Grava o arquivo de log (se houver entradas)
    if log_entries:
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write("=" * 70 + "\n")
            f.write("  NCMs SEM CEST - REVISÃO MANUAL NECESSÁRIA\n")
            f.write("=" * 70 + "\n\n")
            f.write('\n'.join(log_entries) + '\n')
        print(f"[AVISO] {len(log_entries)} NCMs sem correspondencia CEST registrados em: {log_path}")

    # --- Etapa 5: Geração do arquivo TXT para importação ---
    print("Gerando o arquivo de texto para a Domínio...")
    csv_lines = [f'Vigencia;{vigencia_input};Vigencia_Inicial']

    for code, info in sorted(data.items(), key=lambda x: int(x[0])):
        # Código 400 foi desativado pela Receita Federal (válido até 30.04.2015)
        if int(code) == 400:
            continue

        ncm_entries = info.get('ncms_final', [])
        if not ncm_entries:
            continue

        desc = info['descricao'].replace(';', ',').replace('\n', ' ')

        # Registro "Item": Descrição com código SPED no início para facilitar a visualização
        desc_com_codigo = f"{code} - {desc}"
        csv_lines.append(f'Item;{desc_com_codigo};{code}')

        # Registro "Impostos": Linha com todos os campos tributários
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

        # Registros "NCM": Um por linha, com CEST quando disponível (chave composta)
        for ncm_dict in ncm_entries:
            cest = ncm_dict.get('cest', '')
            csv_lines.append(f'NCM;{ncm_dict["ncm"]};{cest};')

    # Salva o arquivo final em encoding latin1 (padrão da Domínio)
    output_path = 'gerardadosporncm_COMPLETO.txt'
    with open(output_path, 'w', encoding='latin1', errors='ignore') as f:
        f.write('\n'.join(csv_lines) + '\n')

    print(f"Arquivo gerado com sucesso em: {output_path}")

    # Estatísticas finais
    total_ncms = sum(len(info.get('ncms_final', [])) for info in data.values())
    total_items = sum(1 for info in data.values() if info.get('ncms_final'))
    print(f"\n--- Resumo ---")
    print(f"Itens (Naturezas de Receita): {total_items}")
    print(f"NCMs exportados: {total_ncms}")
    if log_entries:
        print(f"NCMs para revisão (sem CEST): {len(log_entries)}")


# ---------------------------------------------------------------------------
#  Ponto de entrada
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    main()
