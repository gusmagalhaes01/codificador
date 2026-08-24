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


class TestFormatarDocumento(unittest.TestCase):
    def test_catorze_digitos_saem_como_cnpj(self):
        self.assertEqual(app.formatar_documento("01195716000154"),
                         "01.195.716/0001-54")

    def test_onze_digitos_saem_como_cpf(self):
        self.assertEqual(app.formatar_documento("52998224725"),
                         "529.982.247-25")

    def test_comprimento_desconhecido_sai_cru(self):
        # Mesmo contrato que formatar_cnpj tinha: não inventa formatação
        # para o que não reconhece.
        self.assertEqual(app.formatar_documento("123"), "123")
        self.assertEqual(app.formatar_documento(""), "")

    def test_aceita_entrada_ja_formatada(self):
        self.assertEqual(app.formatar_documento("529.982.247-25"),
                         "529.982.247-25")

    def test_formatar_cnpj_nao_existe_mais(self):
        self.assertFalse(hasattr(app, "formatar_cnpj"))


if __name__ == "__main__":
    unittest.main()
