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


class TestCpfValido(unittest.TestCase):
    def test_cpf_valido(self):
        self.assertTrue(app.cpf_valido("52998224725"))

    def test_digito_verificador_errado(self):
        self.assertFalse(app.cpf_valido("52998224724"))

    def test_todos_digitos_iguais_e_invalido(self):
        # Passa no cálculo dos DVs, mas não é CPF de ninguém — mesma
        # armadilha que cnpj_valido já cobre para os 14 dígitos iguais.
        self.assertFalse(app.cpf_valido("11111111111"))

    def test_comprimento_errado_e_invalido(self):
        self.assertFalse(app.cpf_valido("529982247"))
        self.assertFalse(app.cpf_valido(""))

    def test_cnpj_nao_passa_por_cpf(self):
        self.assertFalse(app.cpf_valido("01195716000154"))


if __name__ == "__main__":
    unittest.main()
