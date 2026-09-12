"""
Testes unitarios para o Tributa NF-e.
Cobre o parser de NCMs, sanitizacao de texto, e o motor de fuzzy matching.

Execucao: python -m pytest tests/ -v
"""
import sys
import os

# Adiciona o diretorio raiz ao path para importar os modulos
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.tributa.parsers.ncm import parse_ncms, expand_range
from src.tributa.matching.fuzzy import find_best_cest, load_manual_overrides


# =========================================================================
#  Testes do Parser de NCMs
# =========================================================================

class TestParseNcmSimples:
    def test_ncm_8_digitos(self):
        result = parse_ncms("22.01.10.00")
        assert any(r["ncm"] == "22011000" and r["ex"] is None for r in result)

    def test_ncm_4_digitos(self):
        result = parse_ncms("22.03")
        assert any(r["ncm"] == "2203" and r["ex"] is None for r in result)

    def test_ncm_incompleto_preenchido_com_zeros(self):
        result = parse_ncms("2710.11.59")
        assert any(r["ncm"] == "27101159" for r in result)

    def test_multiplos_ncms(self):
        text = "22.01.10.00 e 22.02.90.00"
        result = parse_ncms(text)
        ncms = {r["ncm"] for r in result}
        assert "22011000" in ncms
        assert "22029000" in ncms


class TestParseNcmComEx:
    def test_ex_simples(self):
        result = parse_ncms("21.06.90.10 Ex 02")
        assert any(r["ncm"] == "21069010" and r["ex"] == "02" for r in result)

    def test_ex_sem_espaco(self):
        result = parse_ncms("3826.00.00Ex 01")
        assert any(r["ncm"] == "38260000" and r["ex"] == "01" for r in result)

    def test_ex_maiusculo(self):
        result = parse_ncms("2208.90.00 EX 03")
        assert any(r["ncm"] == "22089000" and r["ex"] == "03" for r in result)

    def test_ex_digito_unico(self):
        result = parse_ncms("22.02.10.00 Ex 1")
        assert any(r["ex"] == "01" for r in result)


class TestSanitizacao:
    def test_remove_datas(self):
        result = parse_ncms("Vigente desde 01.01.2016 ate 31.12.2017")
        ncms = {r["ncm"] for r in result}
        assert "01012016" not in ncms
        assert "31122017" not in ncms

    def test_remove_datas_barra(self):
        result = parse_ncms("A partir de 08/03/2016")
        ncms = {r["ncm"] for r in result}
        assert "08032016" not in ncms

    def test_remove_leis(self):
        result = parse_ncms("Conforme Lei 10.147/00 e Decreto 8.950/16, NCM 22.02")
        ncms = {r["ncm"] for r in result}
        assert "2202" in ncms
        assert "10147000" not in ncms

    def test_remove_anos_isolados(self):
        result = parse_ncms("Valido em 2015 e 2026")
        assert len(result) == 0

    def test_reconecta_ncm_quebrado(self):
        result = parse_ncms("22.\n01.10.00")
        assert any(r["ncm"] == "22011000" for r in result)

    def test_exceto_com_parenteses(self):
        text = "Sucos de frutas exceto sucos (sumos) de frutas da posicao 20.09; NCM 22.02"
        result = parse_ncms(text)
        ncms = {r["ncm"] for r in result}
        assert "2009" not in ncms
        assert "2202" in ncms


class TestIntervalos:
    def test_expand_range(self):
        result = expand_range("22.01", "22.03")
        assert result == ["2201", "2202", "2203"]

    def test_intervalo_no_texto(self):
        result = parse_ncms("NCMs 22.01 a 22.03")
        ncms = {r["ncm"] for r in result}
        assert "2201" in ncms
        assert "2202" in ncms
        assert "2203" in ncms


class TestNaoDuplica:
    def test_sem_duplicatas(self):
        text = "22.01.10.00 e mais uma vez 22.01.10.00"
        result = parse_ncms(text)
        ncms = [r["ncm"] for r in result if r["ncm"] == "22011000"]
        assert len(ncms) == 1


class TestFindBestCest:
    def test_override_manual(self):
        manual = {"2203": "0302100"}
        cest, score = find_best_cest("Cerveja de malte", "2203", {}, manual)
        assert cest == "0302100"
        assert score == 1.0

    def test_sem_candidatos(self):
        cest, score = find_best_cest("Descricao qualquer", "99999999", {})
        assert cest == ""
        assert score == 0.0

    def test_match_por_ncm(self):
        cest_db = {"22011000": [{"cest": "0300100", "descricao": "Agua mineral natural"}]}
        cest, score = find_best_cest("Aguas minerais naturais", "22011000", cest_db)
        assert cest == "0300100"
        assert score > 0.30


class TestLoadOverrides:
    def test_fallback_inline(self):
        result = load_manual_overrides()
        assert len(result) >= 7
        assert "2203" in result
        assert "22072000" in result


if __name__ == "__main__":
    import pytest
    pytest.main([__file__, "-v"])