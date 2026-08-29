"""
Prévia do documento no painel de resultado (aba 3).

A prévia é conveniência: nunca pode derrubar o painel. Documento ilegível
devolve motivo em vez de imagem — e é justamente esse o caso que mais precisa
aparecer descrito, porque um arquivo de 0 byte não se resolve nem informando
o valor à mão (não há PDF para carimbar).
"""
import os
import shutil
import sys
import tempfile
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import logica as app


def pdf_de_teste(caminho, paginas=1, largura=595, altura=842):
    """PDF mínimo de verdade, gerado com reportlab (já é dependência)."""
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(caminho, pagesize=(largura, altura))
    for i in range(paginas):
        c.drawString(60, altura - 80, f"W700A VILLARS (10005) pagina {i + 1}")
        c.showPage()
    c.save()


@unittest.skipUnless(app.FITZ_DISPONIVEL, "PyMuPDF não instalado")
class TestRenderizarPrevia(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp(prefix="previa_")
        self.pdf = os.path.join(self.pasta, "protocolo.pdf")
        pdf_de_teste(self.pdf, paginas=2)

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_renderiza_a_primeira_pagina(self):
        img, motivo = app.renderizar_previa_pdf(self.pdf)
        self.assertIsNone(motivo)
        self.assertIsNotNone(img)

    def test_largura_fixa_para_a_coluna_nao_pular_de_tamanho(self):
        img, _ = app.renderizar_previa_pdf(self.pdf)
        self.assertEqual(img.width, app.LARGURA_PREVIA)

    def test_altura_respeita_a_proporcao_da_pagina(self):
        # A4 em pé: mais alta que larga. Se a proporção se perdesse, o
        # documento sairia esticado e o cabeçalho ficaria ilegível.
        img, _ = app.renderizar_previa_pdf(self.pdf)
        self.assertGreater(img.height, img.width)
        self.assertAlmostEqual(img.height / img.width, 842 / 595, places=1)

    def test_largura_personalizada(self):
        img, _ = app.renderizar_previa_pdf(self.pdf, largura=260)
        self.assertEqual(img.width, 260)

    def test_segunda_pagina(self):
        img, motivo = app.renderizar_previa_pdf(self.pdf, pagina=1)
        self.assertIsNone(motivo)
        self.assertIsNotNone(img)

    def test_pagina_fora_da_faixa_cai_na_ultima_em_vez_de_estourar(self):
        img, motivo = app.renderizar_previa_pdf(self.pdf, pagina=99)
        self.assertIsNone(motivo)
        self.assertIsNotNone(img)


@unittest.skipUnless(app.FITZ_DISPONIVEL, "PyMuPDF não instalado")
class TestDocumentoIlegivel(unittest.TestCase):
    """Nenhum destes pode levantar exceção — todos devolvem (None, motivo)."""

    def setUp(self):
        self.pasta = tempfile.mkdtemp(prefix="previa_ruim_")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_arquivo_vazio_explica_o_que_fazer(self):
        # Caso real do lote: doc16004020260821153930.pdf tinha 0 byte.
        vazio = os.path.join(self.pasta, "vazio.pdf")
        open(vazio, "wb").close()
        img, motivo = app.renderizar_previa_pdf(vazio)
        self.assertIsNone(img)
        self.assertIn("vazio", motivo.lower())

    def test_arquivo_inexistente(self):
        img, motivo = app.renderizar_previa_pdf(
            os.path.join(self.pasta, "nao_existe.pdf"))
        self.assertIsNone(img)
        self.assertIn("não encontrado", motivo.lower())

    def test_arquivo_que_nao_e_pdf(self):
        falso = os.path.join(self.pasta, "falso.pdf")
        with open(falso, "wb") as f:
            f.write(b"isto nao e um PDF" * 20)
        img, motivo = app.renderizar_previa_pdf(falso)
        self.assertIsNone(img)
        self.assertTrue(motivo)


@unittest.skipUnless(app.FITZ_DISPONIVEL, "PyMuPDF não instalado")
class TestPaginasDoPdf(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp(prefix="paginas_")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_conta_as_paginas(self):
        p = os.path.join(self.pasta, "tres.pdf")
        pdf_de_teste(p, paginas=3)
        self.assertEqual(app.paginas_do_pdf(p), 3)

    def test_arquivo_ruim_devolve_zero_em_vez_de_estourar(self):
        vazio = os.path.join(self.pasta, "vazio.pdf")
        open(vazio, "wb").close()
        self.assertEqual(app.paginas_do_pdf(vazio), 0)
        self.assertEqual(app.paginas_do_pdf(os.path.join(self.pasta, "sumiu.pdf")), 0)


if __name__ == "__main__":
    unittest.main()
