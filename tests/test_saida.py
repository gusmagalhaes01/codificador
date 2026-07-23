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


class TestNomeSaida(unittest.TestCase):
    def test_prefixa_codigo(self):
        self.assertEqual(
            app.nome_saida_com_codigo("PGR 10004 Klosters.pdf", "10004"),
            "10004 - PGR 10004 Klosters.pdf",
        )

    def test_sanitiza_caractere_invalido_no_codigo(self):
        # barras e dois-pontos não podem aparecer em nome de arquivo no Windows
        resultado = app.nome_saida_com_codigo("nota.pdf", "10/04")
        for proibido in '\\/:*?"<>|':
            self.assertNotIn(proibido, resultado.split(" - ")[0])

    def test_codigo_vazio_devolve_nome_original(self):
        self.assertEqual(app.nome_saida_com_codigo("nota.pdf", ""), "nota.pdf")


if __name__ == "__main__":
    unittest.main()
