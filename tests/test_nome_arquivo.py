import os
import sys
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import identificacao_por_cnpj_6_0 as app
from cadastro_teste import CADASTRO_TESTE


def codigo_de(cnpj):
    return CADASTRO_TESTE[cnpj]["codigo"] if cnpj else None


class TestNomeArquivo(unittest.TestCase):
    def test_codigo_exato_no_nome(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 Klosters.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10004")

    def test_codigo_nao_confunde_com_data(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 Klosters 06.2026.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10004")

    def test_fuzzy_pelo_nome(self):
        cnpj, _ = app.buscar_por_nome_arquivo("ARAUJO LIMA QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10695")

    def test_palavra_tipo_documento_nao_atrapalha(self):
        # "QUITADO" derrubava o score abaixo de 0.72 antes do fix
        cnpj, _ = app.buscar_por_nome_arquivo("ARGENTINA QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10625")

    def test_ambiguidade_conde_de_bonfim_nao_escolhe(self):
        # TRAVA CRÍTICA: dois condomínios quase idênticos -> não pode escolher
        cnpj, _ = app.buscar_por_nome_arquivo("CONDE DE BONFIM.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_ambiguidade_conde_de_bonfim_com_ruido_nao_escolhe(self):
        cnpj, _ = app.buscar_por_nome_arquivo(
            "CENTRO COM CONDE DE BONFIM QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_dois_codigos_no_nome_e_ambiguo(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 10002 Klosters.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_sem_correspondencia(self):
        cnpj, _ = app.buscar_por_nome_arquivo("DOCUMENTO QUALQUER.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)


if __name__ == "__main__":
    unittest.main()
