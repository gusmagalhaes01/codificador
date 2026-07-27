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


class TestCnpjValido(unittest.TestCase):
    def test_cnpjs_reais_do_cadastro_sao_validos(self):
        for cnpj in CADASTRO_TESTE:
            self.assertTrue(app.cnpj_valido(cnpj), f"deveria ser válido: {cnpj}")

    def test_checksum_invalido_do_nf927(self):
        # CNPJ que a FedCorp gerou errado (00.001.208-2497-38) — dígito não bate
        self.assertFalse(app.cnpj_valido("00001208249738"))

    def test_todos_digitos_iguais_e_invalido(self):
        self.assertFalse(app.cnpj_valido("11111111111111"))

    def test_menos_de_14_digitos_e_invalido(self):
        self.assertFalse(app.cnpj_valido("123"))
        self.assertFalse(app.cnpj_valido(""))


if __name__ == "__main__":
    unittest.main()
