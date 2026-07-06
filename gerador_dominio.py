import os
import re
import json
import urllib.request
from bs4 import BeautifulSoup

def expand_range(start_ncm, end_ncm):
    start = int(start_ncm.replace('.', ''))
    end = int(end_ncm.replace('.', ''))
    return [str(i).zfill(len(start_ncm.replace('.', ''))) for i in range(start, end + 1)]

def parse_ncms(text_block):
    # Remove textos de exceção que atrapalham a extração
    text_block = re.sub(r'exceto.*?(;|\.|$|\))', '', text_block, flags=re.IGNORECASE)
    ncms = set()
    
    # Busca intervalos de NCM "XX.XX a YY.YY"
    ranges = re.findall(r'(\d{2}\.\d{2})\s+a\s+(\d{2}\.\d{2})', text_block)
    for start, end in ranges:
        try:
            ncms.update(expand_range(start, end))
            text_block = text_block.replace(f'{start} a {end}', '')
        except:
            pass
        
    # Busca NCMs isolados (2, 4, 6 ou 8 digitos separados por ponto)
    singles = re.findall(r'\b(?:\d{1,4}\.){1,3}\d{1,4}\b', text_block)
    for s in singles:
        clean = s.replace('.', '')
        # A Dominio aceita 4 dígitos (sintético) ou 8 dígitos (analítico)
        if len(clean) == 4 or len(clean) == 8:
            ncms.add(clean)
        elif 4 < len(clean) < 8:
            # Preenche com zeros à direita até formar 8 dígitos
            ncms.add(clean.ljust(8, '0'))
            
    return list(ncms)

def extract_txt_from_doc(doc_path):
    import win32com.client as win32
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        txt_path = os.path.abspath('tabela_temp.txt')
        doc.SaveAs(txt_path, FileFormat=2) # 2 = wdFormatText
        doc.Close()
        return txt_path
    except Exception as e:
        print(f'Erro ao extrair DOC: {e}')
        return None
    finally:
        try:
            word.Quit()
        except:
            pass

def download_latest_sped_table():
    print('Procurando a versão mais recente da Tabela 4.3.10 no site do SPED...')
    url = 'http://sped.rfb.gov.br/pasta/show/1616'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
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
            print(f'Baixando automaticamente a tabela atualizada...')
            
            # Remove tabelas antigas baixadas pelo script para evitar lixo
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
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        out_path = os.path.join(os.getcwd(), 'tabela_temp.txt')
        doc.SaveAs2(out_path, FileFormat=2) # 2 = wdFormatText
        doc.Close(False)
        return out_path
    except Exception as e:
        print(f"Erro ao ler o arquivo Word: {e}")
        return None
    finally:
        try:
            word.Quit()
        except:
            pass

def main():
    print("=" * 60)
    print(" GERADOR DE ARQUIVO DE NCMs PARA A DOMÍNIO SISTEMAS ")
    print("=" * 60)
    
    # Pergunta os parâmetros iniciais ao usuário
    vigencia_input = input("Digite a data da vigência inicial (ex: 01/01/2026): ").strip()
    if not vigencia_input:
        vigencia_input = "01/01/2026" # Default se der enter direto
        
    regime_input = input("Qual o regime da empresa? (SN para Simples Nacional, LR para Lucro Real, LP para Lucro Presumido): ").strip().upper()
    
    # Mapeia o regime para o Tipo de Contribuição esperado pelo Domínio
    if regime_input == 'LR':
        tipo_contrib_input = 'N' # Não Cumulativo
    elif regime_input == 'LP':
        tipo_contrib_input = 'C' # Cumulativo
    else:
        # Para SN (Simples Nacional) ou qualquer outra resposta, deixa em branco
        tipo_contrib_input = ''
    
    print("\nIniciando a extração dos NCMs...")
    
    # Tenta baixar a versão mais recente do site primeiro
    txt_path = 'tabela_temp.txt'
    baixado = download_latest_sped_table()
    
    # Se baixou com sucesso, processa o DOC baixado
    if baixado:
        print(f"Convertendo tabela recém-baixada ({baixado}) para TXT...")
        txt_path = extract_txt_from_doc(baixado)
    else:
        # Se falhou, procura um arquivo .doc ou .txt na pasta que contenha Tabela no nome
        if not os.path.exists(txt_path):
            doc_files = [f for f in os.listdir('.') if f.lower().endswith('.doc') and 'tabela' in f.lower()]
            if doc_files:
                print(f"Encontrado arquivo DOC local: {doc_files[0]}")
                print("Convertendo DOC para TXT... (isso pode levar alguns segundos)")
                txt_path = extract_txt_from_doc(doc_files[0])
            else:
                txt_path = [f for f in os.listdir('.') if f.lower().endswith('.txt') and 'tabela' in f.lower()]
                if txt_path:
                    txt_path = txt_path[0]
                else:
                    print("ERRO: Nenhuma tabela SPED em .doc ou .txt encontrada na pasta.")
                    return

    if not txt_path or not os.path.exists(txt_path):
        return

    print("Lendo o conteúdo da tabela...")
    with open(txt_path, 'r', encoding='latin1', errors='replace') as f:
        text = f.read()

    # Divide o texto procurando os códigos SPED (ex: 101, 201, 302, etc - que ficam isolados em uma linha)
    blocks = re.split(r'(?m)^([1-9]\d{2})$', text)
    data = {}

    print("Extraindo os Códigos e NCMs...")
    for i in range(1, len(blocks), 2):
        code = blocks[i]
        block_text = blocks[i+1].strip()
        
        lines = [l.strip() for l in block_text.split('\n') if l.strip()]
        if not lines: continue
        desc = lines[0] # Primeira linha logo após o código é a descrição
        
        ncms = parse_ncms(block_text)
        if ncms:
            if code not in data:
                data[code] = {'descricao': desc, 'ncms': set()}
            data[code]['ncms'].update(ncms)

    for code in data:
        data[code]['ncms'] = sorted(list(data[code]['ncms']))

    # Gera o JSON estruturado para facilitar futuras leituras
    json_path = 'tabela_sped_ncms.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"JSON estruturado gerado em: {json_path}")

    # Aplica as regras de validação do Sistema Domínio:
    # 1. Não importar NCM analítico (8 dig) se o sintético (4 dig) já existe
    # 2. Não duplicar o mesmo NCM em códigos de configuração diferentes
    synthetic_ncms = set()
    for code, info in data.items():
        for ncm in info['ncms']:
            if len(ncm) == 4:
                synthetic_ncms.add(ncm)

    seen_ncms = set()
    for code in sorted(data.keys(), key=lambda x: int(x)):
        valid_ncms = []
        for ncm in data[code]['ncms']:
            # Regra 1: Evitar conflito sintético x analítico
            if len(ncm) == 8 and ncm[:4] in synthetic_ncms:
                continue
            # Regra 2: Evitar NCMs duplicados no mesmo arquivo
            if ncm in seen_ncms:
                continue
            seen_ncms.add(ncm)
            valid_ncms.append(ncm)
        data[code]['ncms'] = valid_ncms

    # Gera o arquivo TXT no novo layout da Domínio (com Impostos)
    print("Gerando o arquivo de texto para a Domínio...")
    csv_lines = [f'Vigencia;{vigencia_input};Vigencia_Inicial']
    for code, info in sorted(data.items(), key=lambda x: int(x[0])):
        # Filtra apenas o código 400 que é obsoleto (até 30.04.2015). Os demais 40x (411, 415, etc) continuam ativos
        if int(code) == 400:
            continue
            
        desc = info['descricao'].replace(';', ',').replace('\n', ' ')
        ncms = info['ncms']
        if ncms:
            # Registro de Item
            # Adicionamos o código SPED diretamente no texto da descrição para que o usuário consiga visualizar na tela principal do Domínio
            desc_com_codigo = f"{code} - {desc}"
            csv_lines.append(f'Item;{desc_com_codigo};{code}')
            
            # Registro de Impostos (Mapeamento do Layout Personalizado)
            # Colunas: Impostos; TIPO_CONTRIBUICAO; CODIGO_CST_ENTRADA; VINCULO_CREDITO; BASE_CREDITO; ALIQ_PIS_ENTRADA; ALIQ_COFINS_ENTRADA; CODIGO_CST_SAIDA; NATUREZA_RECEITA; ALIQ_PIS_SAIDA; ALIQ_COFINS_SAIDA; CST_ICMS_ENTRADA; CST_ICMS_SAIDA; ALIQ_ICMS; CST_IPI_ENTRADA; CST_IPI_SAIDA; ALIQ_IPI; SUJEITO_PIS_COFINS; TIPO_TRIBUTACAO_PIS_COFINS
            impostos_row = [
                "Impostos", # 1: Tipo
                tipo_contrib_input, # 2: Tipo Contribuição (Variável definida pelo usuário: N, C ou vazio)
                "70",       # 3: CST Entrada
                "01",       # 4: Vinculo Credito
                "1",        # 5: Base de Credito
                "0",        # 6: Aliq PIS Entrada
                "0",        # 7: Aliq Cofins Entrada
                "04",       # 8: CST Saida
                str(code),  # 9: Natureza Receita (Código SPED)
                "0",        # 10: Aliq PIS Saida
                "0",        # 11: Aliq Cofins Saida
                "",         # 12: CST ICMS Entrada
                "",         # 13: CST ICMS Saida
                "",         # 14: Aliq ICMS
                "",         # 15: CST IPI Entrada
                "",         # 16: CST IPI Saida
                "",         # 17: Aliq IPI
                "S",        # 18: Produto Sujeito a PIS/COFINS
                ""          # 19: Tipo de Tributação PIS/COFINS
            ]
            csv_lines.append(";".join(impostos_row))
            
            # Registros de NCM
            for ncm in ncms:
                csv_lines.append(f'NCM;{ncm};')

    # Salva o resultado final no arquivo TXT
    csv_path = 'gerardadosporncm_COMPLETO.txt'
    with open(csv_path, 'w', encoding='latin1', errors='ignore') as f:
        f.write('\n'.join(csv_lines) + '\n')
    
    print(f"Arquivo gerado com sucesso em: {csv_path}")

if __name__ == '__main__':
    main()
