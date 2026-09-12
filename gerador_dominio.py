#!/usr/bin/env python3
# ============================================================================
# TRIBUTA NF-e - Gerador Automatico de Arquivo de Importacao
# SPED Monofasico -> Dominio Sistemas
#
# Autor : Henrique Meirelles (github.com/henriquemeirelleshfmc)
# Repo  : https://github.com/henriquemeirelleshfmc/sped-monofasico-dominio
# Tabela: 4.3.10 - Produtos Sujeitos a Aliquotas Diferenciadas (Monofasica)
#
# Este ficheiro e o ponto de entrada para execucao direta:
#   python gerador_dominio.py
#
# A logica de negocios esta modularizada em src/tributa/
# ============================================================================

from src.tributa.main import main

if __name__ == "__main__":
    main()