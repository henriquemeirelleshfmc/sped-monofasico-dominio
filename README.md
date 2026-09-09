# TRIBUTA NF-e — Gerador Automático para Domínio Sistemas

[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

Script de automação fiscal que extrai, classifica e exporta NCMs da Tabela SPED 4.3.10 (Incidência Monofásica — Bebidas Frias) no formato de importação do ERP Domínio Sistemas.

## O que este script faz

1. **Baixa automaticamente** a versão mais recente da Tabela 4.3.10 do portal oficial do SPED
2. **Extrai NCMs e Exceções (EX)** do documento Word com sanitização avançada (remove datas, leis e fragmentações de coluna)
3. **Cruza com 3 bases governamentais** para resolver duplicidades:
   - **SPED** (Receita Federal) — Naturezas da Receita
   - **CONFAZ** (Convênio ICMS 52/17) — Tabela CEST
   - **TIPI** (Tabela de Incidência do IPI) — Descrições de Exceções
4. **Aplica desempate inteligente em 3 níveis:**
   - Nível 1: NCM Único (sem conflito)
   - Nível 2: NCM com Exceção (EX) — enriquecido pela TIPI
   - Nível 3: Fuzzy Matching híbrido (difflib + keywords + detecção de volumes)
5. **Gera o arquivo TXT** pronto para importação na Domínio, com chave composta `(NCM, CEST)` globalmente única
6. **Registra pendências** em log de auditoria (`ncm_sem_cest_revisao.log`)

## Instalação

```bash
# Clone o repositório
git clone https://github.com/henriquemeirelleshfmc/sped-monofasico-dominio.git
cd sped-monofasico-dominio

# Instale as dependências
pip install -r requirements.txt
```

### Requisitos

- Python 3.8+
- Microsoft Word (para conversão DOC -> TXT no Windows)
- Conexão à internet (para download das tabelas; usa cache local de 30 dias)

## Como usar

```bash
python gerador_dominio.py
```

O script fará 2 perguntas no terminal:
- **Data de Vigência Inicial** (ex: `01/01/2026`)
- **Regime da Empresa** (`SN` = Simples Nacional, `LR` = Lucro Real, `LP` = Lucro Presumido)

### Saída

| Arquivo | Descrição |
|:---|:---|
| `gerardadosporncm_COMPLETO.txt` | Arquivo final para importação na Domínio (encoding `latin1`) |
| `ncm_sem_cest_revisao.log` | NCMs que exigem revisão manual do analista fiscal |
| `tabela_sped_ncms.json` | Dump estruturado dos NCMs extraídos (debug) |

## Overrides Manuais

Quando a descrição do SPED/TIPI é semanticamente incompatível com o CONFAZ (ex: "Preparações compostas" vs "Xarope para refrigerantes"), o fuzzy matching falha. Para estes casos, existe um ficheiro de mapeamento manual:

```
data/manual_overrides.json
```

O analista fiscal pode editar este JSON para adicionar ou corrigir mapeamentos NCM -> CEST sem precisar alterar o código Python:

```json
{
  "21069010": {"cest": "0301000", "motivo": "Xaropes para refrigerantes"},
  "22029000": {"cest": "0301100", "motivo": "NCM desatualizado no SPED"}
}
```

## Cache Inteligente

O script cria caches locais para evitar sobrecarregar os portais do governo:

| Cache | Fonte | Validade |
|:---|:---|:---|
| `confaz_cest_cache.json` | CONFAZ (Conv. ICMS 52/17) | 30 dias |
| `tipi_cache.json` | Receita Federal (TIPI Excel) | 30 dias |

Para forçar a renovação, basta apagar o ficheiro de cache correspondente.

## Estrutura do Projeto

```
monofasico/
├── gerador_dominio.py              # Script principal
├── requirements.txt                # Dependências Python
├── data/
│   └── manual_overrides.json       # Mapeamentos NCM->CEST manuais
├── .gitignore
└── README.md
```

## Dependências

| Pacote | Uso |
|:---|:---|
| `beautifulsoup4` | Scraping do CONFAZ e SPED |
| `pandas` + `openpyxl` | Leitura da TIPI (Excel) |
| `pywin32` | Conversão DOC -> TXT (Windows) |
| `rich` | Interface visual no terminal |

## Licença

MIT License — Henrique Meirelles