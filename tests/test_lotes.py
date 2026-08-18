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

SAIDA = os.path.join("C:", os.sep, "saida")


class TestCaminhoDoLote(unittest.TestCase):
    def test_primeiro_lote(self):
        # 1º ao 20º arquivo (índices 0-19) vão para "Lote 01"
        self.assertEqual(app.caminho_do_lote(SAIDA, 0, 20), os.path.join(SAIDA, "Lote 01"))
        self.assertEqual(app.caminho_do_lote(SAIDA, 19, 20), os.path.join(SAIDA, "Lote 01"))

    def test_segundo_lote_comeca_no_21o(self):
        self.assertEqual(app.caminho_do_lote(SAIDA, 20, 20), os.path.join(SAIDA, "Lote 02"))
        self.assertEqual(app.caminho_do_lote(SAIDA, 39, 20), os.path.join(SAIDA, "Lote 02"))

    def test_numeracao_com_dois_digitos(self):
        # "Lote 10", não "Lote 010" — e ordena certo no explorador de arquivos
        self.assertEqual(app.caminho_do_lote(SAIDA, 180, 20), os.path.join(SAIDA, "Lote 10"))

    def test_lote_de_tamanho_diferente(self):
        self.assertEqual(app.caminho_do_lote(SAIDA, 5, 5), os.path.join(SAIDA, "Lote 02"))

    def test_tamanho_invalido_nao_separa(self):
        # 0 ou negativo = sem separação: devolve a pasta de saída original,
        # em vez de estourar divisão por zero no meio de um processamento.
        self.assertEqual(app.caminho_do_lote(SAIDA, 7, 0), SAIDA)
        self.assertEqual(app.caminho_do_lote(SAIDA, 7, -3), SAIDA)


if __name__ == "__main__":
    unittest.main()
