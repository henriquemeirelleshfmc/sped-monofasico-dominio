"""
Configuracoes globais, constantes e instancia do console Rich.
"""
import sys

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich import box

# Garante encoding UTF-8 no stdout (resolve UnicodeEncodeError no Windows CP1252)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Instancia global do console Rich
console = Console(force_terminal=True)

# URLs das fontes de dados governamentais
SPED_URL = "http://sped.rfb.gov.br/pasta/show/1616"
CONFAZ_URL = "https://www.confaz.fazenda.gov.br/legislacao/convenios/2017/CV052_17"
TIPI_URL = "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/legislacao/documentos-e-arquivos/tipi.xlsx"

# Validade do cache em segundos (30 dias)
CACHE_MAX_AGE = 30 * 24 * 3600

# Mapeamento de regimes tributarios para codigos da Dominio
REGIME_MAP = {
    "LR": ("N", "Lucro Real"),       # Nao Cumulativo
    "LP": ("C", "Lucro Presumido"),   # Cumulativo
    "SN": ("", "Simples Nacional"),   # Em branco
}

# Stopwords para o motor de fuzzy matching fiscal
STOPWORDS = {
    "de", "do", "da", "dos", "das", "e", "ou", "em", "com", "a", "o", "as", "os",
    "para", "por", "no", "na", "nos", "nas", "um", "uma", "que", "se", "ao", "ate",
    "nao", "nem", "mais", "outros", "outras", "incluindo", "inclusive", "exceto",
}