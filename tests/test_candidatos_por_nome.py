import os
import sys
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app
from cadastro_teste import CADASTRO_TESTE


class TestCandidatosPorNome(unittest.TestCase):
    def test_nome_unico_devolve_um_candidato(self):
        candidatos = app.candidatos_por_nome("ARAUJO LIMA QUITADO 05.26.pdf", "", CADASTRO_TESTE)
        self.assertEqual(candidatos, ["40338774000141"])

    def test_ambiguidade_conde_de_bonfim_devolve_os_dois(self):
        # Aqui é o oposto de buscar_por_nome_arquivo: em vez de bloquear,
        # devolve os dois candidatos pro usuário escolher manualmente.
        candidatos = app.candidatos_por_nome("CONDE DE BONFIM.pdf", "", CADASTRO_TESTE)
        self.assertEqual(
            set(candidatos), {"05695194000100", "29361458000158"})

    def test_sem_correspondencia_devolve_lista_vazia(self):
        candidatos = app.candidatos_por_nome("DOCUMENTO QUALQUER.pdf", "", CADASTRO_TESTE)
        self.assertEqual(candidatos, [])

    def test_usa_texto_ocr_quando_nome_do_arquivo_e_generico(self):
        texto = "CO-ESTIPULANTE: KLOSTERS CNPJ: 01.195.716/0001-54"
        candidatos = app.candidatos_por_nome("DIGITALIZACAO_001.pdf", texto, CADASTRO_TESTE)
        self.assertEqual(candidatos, ["01195716000154"])

    def test_texto_vazio_nao_quebra(self):
        candidatos = app.candidatos_por_nome("SAN REMO.pdf", "", CADASTRO_TESTE)
        self.assertEqual(candidatos, ["08578541000103"])

    def test_respeita_limite(self):
        candidatos = app.candidatos_por_nome("CONDE DE BONFIM.pdf", "", CADASTRO_TESTE, limite=1)
        self.assertEqual(len(candidatos), 1)


if __name__ == "__main__":
    unittest.main()
