# SPED Monofásico - Gerador Domínio Sistemas

Este script foi desenvolvido para automatizar e otimizar a importação de alíquotas diferenciadas (Incidência Monofásica e por Pauta) da Tabela 4.3.10 do SPED para o ERP da Domínio Sistemas.

O script lê dinamicamente as Naturezas da Receita da Receita Federal e as cruza com a tabela de CEST do CONFAZ, garantindo que o cadastro de impostos por NCM dentro da Domínio não sofra erros de duplicidade e fique 100% aderente às normas fiscais vigentes.

## Funcionalidades Principais

* **Extração Direta do SPED**: O script baixa a versão mais recente da Tabela 4.3.10 diretamente do site oficial da EFD Contribuições, sem necessidade de arquivos manuais.
* **Scraper Automático do CONFAZ**: Conecta-se ao site da Fazenda (Convênio ICMS 52/17) e varre 31 tabelas extraindo CESTs e NCMs em tempo real para o desempate tributário (com sistema de cache inteligente de 30 dias).
* **Tratamento de Exceções (EX)**: Lê nativamente marcações como "Ex 01" da tabela da Receita Federal, criando as derivações corretas para não perder a especificidade do código tributário.
* **Motor Híbrido de Similaridade**: Conta com um algoritmo que cruza as descrições dos NCMs entre o SPED e o CONFAZ. O algoritmo faz expansão de NCMs (ex: SPED '2201' -> CONFAZ '22011000') e calcula um *Score de Similaridade* combinando Lógica Fuzzy (difflib) com Intersecção de Palavras-Chave, garantindo que as águas, cervejas especiais e refrigerantes sejam designados ao CEST exato.
* **Log de Auditoria**: Em vez de parar ou falhar, quando as descrições entre a Receita e o CONFAZ são tão distintas que impossibilitam a conexão segura (Score < 35%), o script preserva a exportação da Natureza da Receita e cria o arquivo `ncm_sem_cest_revisao.log` avisando o analista fiscal sobre a pendência.

## Como usar

1.  Tenha o Python 3.8+ instalado.
2.  Instale as dependências:
    ```bash
    pip install requests beautifulsoup4
    ```
3.  Execute o script:
    ```bash
    python gerador_dominio.py
    ```
4.  O script fará 2 perguntas simples no console:
    *   **Data de Vigência Inicial**: (ex: 01/01/2026)
    *   **Regime da Empresa**: (SN = Simples Nacional, LR = Lucro Real, LP = Lucro Presumido)
5.  O script criará o arquivo `gerardadosporncm_COMPLETO.txt` pronto para ser importado no Domínio Sistemas (padrão de codificação `latin1`).

## Fluxo de Exportação para o Domínio Sistemas

A integração com o Domínio depende do formato estrito do sistema. O arquivo final carrega o layout sequencial contendo `NCM;NCM;CEST;`.
O uso de chaves compostas pelo script garante que um mesmo NCM (ex: 2203 para Cervejas) seja importado em todas as Naturezas de Receita devidas, sem que a Domínio rejeite a importação por violação de unicidade.

## Manutenção

Caso as regras do layout de importação mudem, adapte a etapa 5 `Etapa 5: Geração do arquivo TXT para importação` diretamente no código, onde as strings de *csv_lines* são formatadas.
