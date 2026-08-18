import os
import sys
import tempfile
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app


def sessao(indice, tamanho=1000):
    """Uma sessão sintética de log, do tamanho pedido, marcada com um índice
    pra dar pra saber qual sobreviveu à rotação."""
    return f"SESSAO {indice} " + ("x" * tamanho)


class TestRotacionarLog(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.caminho = os.path.join(self.pasta, "processamento.log")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_arquivo_abaixo_do_limite_fica_intacto(self):
        conteudo = "\n\n".join(sessao(i) for i in range(3))
        with open(self.caminho, "w", encoding="utf-8") as f:
            f.write(conteudo)
        app.rotacionar_log(self.caminho, limite_bytes=1_000_000)
        with open(self.caminho, encoding="utf-8") as f:
            self.assertEqual(f.read(), conteudo)

    def test_arquivo_sem_existir_nao_quebra(self):
        app.rotacionar_log(self.caminho, limite_bytes=1000)  # não lança exceção

    def test_descarta_sessoes_mais_antigas_mantem_as_recentes(self):
        sessoes = [sessao(i) for i in range(10)]
        with open(self.caminho, "w", encoding="utf-8") as f:
            f.write("\n\n".join(sessoes))

        app.rotacionar_log(self.caminho, limite_bytes=3000)

        with open(self.caminho, encoding="utf-8") as f:
            resultado = f.read()
        self.assertLessEqual(len(resultado.encode("utf-8")), 3000)
        self.assertIn("SESSAO 9", resultado)  # a mais recente sobrevive
        self.assertNotIn("SESSAO 0", resultado)  # a mais antiga foi descartada

    def test_uma_sessao_so_maior_que_o_limite_nao_trava(self):
        conteudo = sessao(0, tamanho=5000)
        with open(self.caminho, "w", encoding="utf-8") as f:
            f.write(conteudo)
        app.rotacionar_log(self.caminho, limite_bytes=1000)
        with open(self.caminho, encoding="utf-8") as f:
            resultado = f.read()
        self.assertIn("SESSAO 0", resultado)


if __name__ == "__main__":
    unittest.main()
