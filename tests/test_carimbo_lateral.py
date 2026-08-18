import io
import os
import re
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
from reportlab.pdfbase.pdfmetrics import stringWidth

TEXTO_VALOR = "15 un × R$ 3,85 = R$ 57,75"


def posicao_do_texto(buffer):
    """Lê a coordenada em que o texto foi realmente desenhado, direto do
    content stream do PDF (reportlab grava "1 0 0 1 <x> <y> Tm" sem
    compressão). Comparar bytes do PDF não serviria: o cabeçalho é igual
    qualquer que seja o alinhamento."""
    dados = PdfReader(buffer).pages[0].get_contents().get_data()
    m = re.search(rb"1 0 0 1 ([-\d.]+) ([-\d.]+) Tm", dados)
    assert m, f"nenhum operador Tm encontrado em {dados!r}"
    return float(m.group(1)), float(m.group(2))


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


class TestAlinhamentoDireita(unittest.TestCase):
    def test_alinhado_a_esquerda_comeca_em_x(self):
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 567, 814,
            False, 0, "esquerda")
        x, y = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, 567, places=1)
        self.assertAlmostEqual(y, 814, places=1)

    def test_alinhado_a_direita_termina_em_x(self):
        """Alinhado à direita o texto TERMINA em x, então começa em
        x - largura. É isso que encosta o carimbo na margem direita."""
        largura_texto = stringWidth(TEXTO_VALOR, "Helvetica", 10)
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 567, 814,
            False, 0, "direita")
        x, y = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, 567 - largura_texto, places=1)
        self.assertAlmostEqual(y, 814, places=1)
        self.assertLess(x, 567)

    def test_padrao_continua_alinhando_a_esquerda(self):
        """Sem o parâmetro novo, o texto sai na mesma posição de sempre."""
        sem_parametro = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, False)
        com_parametro = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, False,
            0, "esquerda")
        self.assertEqual(posicao_do_texto(sem_parametro), (40.0, 40.0))
        self.assertEqual(posicao_do_texto(sem_parametro),
                         posicao_do_texto(com_parametro))

    def test_centralizado_continua_centralizando(self):
        """centralizado=True precisa seguir equivalendo a "centro"."""
        largura_texto = stringWidth(TEXTO_VALOR, "Helvetica", 10)
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, True)
        x, _ = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, (595.0 - largura_texto) / 2, places=1)


class TestCarimbosExtras(unittest.TestCase):
    def setUp(self):
        import tempfile
        from pypdf import PdfWriter
        self.pasta = tempfile.mkdtemp()
        self.entrada = os.path.join(self.pasta, "entrada.pdf")
        self.saida = os.path.join(self.pasta, "saida.pdf")
        escritor = PdfWriter()
        escritor.add_blank_page(width=595, height=842)
        escritor.add_blank_page(width=595, height=842)
        with open(self.entrada, "wb") as f:
            escritor.write(f)
        self.config = {
            "fonte": "Helvetica", "tamanho": 10, "cor": "#000000",
            "x": 40, "y": 40, "centralizado": False, "angulo": 90,
        }

    def tearDown(self):
        import shutil
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_sem_extras_carimba_so_o_principal(self):
        app.processar_pdf(self.entrada, self.saida, "10004 KLOSTERS", self.config)
        texto = app.extrair_texto_pdf(self.saida)
        self.assertIn("KLOSTERS", texto)
        self.assertNotIn("57,75", texto)

    def test_os_dois_carimbos_saem_em_todas_as_paginas(self):
        extras = [{
            "texto": "15 un × R$ 3,85 = R$ 57,75", "fonte": "Helvetica",
            "tamanho": 10, "cor": "#000000", "x": 567, "y": 814,
            "centralizado": False, "angulo": 0, "alinhamento": "direita",
        }]
        app.processar_pdf(self.entrada, self.saida, "10004 KLOSTERS",
                          self.config, extras)
        from pypdf import PdfReader
        paginas = PdfReader(self.saida).pages
        self.assertEqual(len(paginas), 2)
        for pagina in paginas:
            texto = pagina.extract_text()
            self.assertIn("KLOSTERS", texto)
            self.assertIn("57,75", texto)


if __name__ == "__main__":
    unittest.main()
