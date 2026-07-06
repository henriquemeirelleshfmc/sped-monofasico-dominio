# Gerador de NCMs SPED (Monofásico) para Domínio Sistemas

Este projeto automatiza a extração e a formatação da Tabela Oficial do SPED (Tabela 4.3.10 - Produtos Sujeitos à Incidência Monofásica) para facilitar a importação em lote no sistema ERP da **Domínio Sistemas**.

## 🚀 Como funciona?

O script é **100% autônomo e dinâmico**:
1. **Download Automático**: Ele acessa o portal do SPED da Receita Federal e faz o download da versão mais recente da Tabela 4.3.10 (em formato `.doc`).
2. **Processamento**: O script converte o `.doc` para texto, varre o conteúdo com Inteligência de Padrões (Regex) e localiza todos os Códigos de Receita (ex: 101, 102, 103) e seus respectivos NCMs vinculados.
3. **Malha Fina**: Ele aplica regras de validação exclusivas da Domínio Sistemas, como remover NCMs analíticos caso o sintético já exista, remover duplicidades e excluir códigos que foram desativados (como o 400).
4. **Layout Final**: Monta um arquivo de texto estruturado (`gerardadosporncm_COMPLETO.txt`) contendo a Vigência, o Item, as Alíquotas, Códigos de CST e NCMs vinculados, no formato pronto para a leitura do importador da Domínio.

## ⚙️ Pré-requisitos

Para rodar este script, você precisará ter instalado no computador:

- **Python 3.x**
- **Microsoft Word** (Necessário pois o script utiliza automação COM do Windows para ler a formatação original das tabelas).
- As bibliotecas Python `beautifulsoup4` e `pywin32`.

Você pode instalar as bibliotecas usando:
```bash
pip install beautifulsoup4 pywin32
```

## 🛠️ Como usar

1. Dê um duplo clique no script `gerador_dominio.py` ou rode ele pelo terminal:
   ```bash
   python gerador_dominio.py
   ```
2. O script vai te perguntar a **data da vigência inicial** (ex: `01/01/2026`).
3. O script vai perguntar o **regime tributário da empresa** (`SN` para Simples Nacional, `LR` para Lucro Real, `LP` para Lucro Presumido) para ajustar a letra da coluna "Tipo de Contribuição" (`N`, `C` ou vazio).
4. Aguarde! O script vai acessar a internet, baixar a tabela da Receita Federal e fazer todo o processamento.
5. No final, um arquivo chamado **`gerardadosporncm_COMPLETO.txt`** será gerado na mesma pasta.

## 📤 Como importar na Domínio Sistemas

1. **(Passo Único)** Acesse a tela de configurações de importação da Domínio e importe o layout customizado que vem junto com este projeto: **`Dados de Impostos - Personalizado.xml`**. Esse layout foi modificado para exibir o código SPED diretamente como o ID Sequencial.
2. Com o layout devidamente atualizado, feche e abra a tela de importação.
3. Importe o arquivo **`gerardadosporncm_COMPLETO.txt`** recém-gerado.

> **💡 Dica de Visualização:** O script injeta o código da Natureza da Receita (ex: 101) no início da descrição do produto (ex: `101 - Gasolinas, Exceto...`). Assim, ao abrir a grade da Domínio, você bate o olho e já enxerga o código SPED de cada linha sem precisar entrar nas configurações do item!

## 📦 Arquivos do Projeto

- `gerador_dominio.py`: O coração do projeto. O script que baixa e processa tudo.
- `Dados de Impostos - Personalizado.xml`: O Layout XML preparado exclusivamente para a Domínio ler o TXT sem dar erros de números sequenciais automáticos.
- `tabela_sped_ncms.json`: Arquivo gerado temporariamente durante a execução como uma cópia de segurança dos dados extraídos do SPED.
