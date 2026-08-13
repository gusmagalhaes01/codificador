import io
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
from pypdf import PdfReader


class TestMontarTextoProtocoloCorreio(unittest.TestCase):
    def test_formata_codigo_nome_cnpj(self):
        texto = app.montar_texto_protocolo_correio("10005", "VILLARS", "07945453000130")
        self.assertEqual(texto, "10005 VILLARS - 07.945.453/0001-30")


class TestCriarOverlayRotacionado(unittest.TestCase):
    def _texto_do_buffer(self, buffer):
        return PdfReader(buffer).pages[0].extract_text() or ""

    def test_angulo_90_texto_continua_extraivel(self):
        buffer = app.criar_overlay(
            largura=595, altura=842, texto="10005 VILLARS - 07.945.453/0001-30",
            fonte="Helvetica-Bold", tamanho=12, cor="#000000",
            x=0, y=0, centralizado=False, angulo=90,
        )
        texto_extraido = self._texto_do_buffer(buffer)
        self.assertIn("10005", texto_extraido)
        self.assertIn("VILLARS", texto_extraido)

    def test_angulo_zero_e_o_padrao_sem_quebrar_chamada_antiga(self):
        # Chamada exatamente como o resto do app já usa hoje, sem `angulo` —
        # precisa continuar funcionando sem alterar assinatura obrigatória.
        buffer = app.criar_overlay(
            largura=595, altura=842, texto="10004 KLOSTERS - CIPAA",
            fonte="Helvetica-Bold", tamanho=14, cor="#000000",
            x=120, y=815, centralizado=False,
        )
        texto_extraido = self._texto_do_buffer(buffer)
        self.assertIn("KLOSTERS", texto_extraido)


if __name__ == "__main__":
    unittest.main()
