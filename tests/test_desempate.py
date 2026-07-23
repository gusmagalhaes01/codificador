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

IMODATA_EMITENTE = "31850191000104"   # não cadastrado
MANHATTAN = "07448975000126"          # cadastrado (11194)


class TestDesempate(unittest.TestCase):
    def test_dois_candidatos_so_um_cadastrado(self):
        resultado = app.desempatar_por_cadastro(
            [IMODATA_EMITENTE, MANHATTAN], CADASTRO_TESTE)
        self.assertEqual(resultado, [MANHATTAN])

    def test_dois_cadastrados_continua_ambiguo(self):
        dois = ["07448975000126", "10864886000175"]  # MANHATTAN e ARGENTINA
        resultado = app.desempatar_por_cadastro(dois, CADASTRO_TESTE)
        self.assertEqual(resultado, dois)

    def test_nenhum_cadastrado_inalterado(self):
        nenhum = ["31850191000104", "12184361000114"]
        resultado = app.desempatar_por_cadastro(nenhum, CADASTRO_TESTE)
        self.assertEqual(resultado, nenhum)

    def test_um_candidato_so_inalterado(self):
        resultado = app.desempatar_por_cadastro([MANHATTAN], CADASTRO_TESTE)
        self.assertEqual(resultado, [MANHATTAN])


if __name__ == "__main__":
    unittest.main()
