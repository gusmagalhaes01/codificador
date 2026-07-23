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


class TestInfra(unittest.TestCase):
    def test_app_importa(self):
        self.assertTrue(hasattr(app, "buscar_por_nome_arquivo"))

    def test_cadastro_tem_9_entradas(self):
        self.assertEqual(len(CADASTRO_TESTE), 9)
        # todos os CNPJs têm 14 dígitos
        for cnpj in CADASTRO_TESTE:
            self.assertEqual(len(cnpj), 14)

    def test_fixtures_existem_e_tem_texto(self):
        dados = os.path.join(_AQUI, "dados")
        for nome in ("fedcorp_recibo_hifen.txt", "nfse_ff.txt", "nfse_imodata.txt",
                     "nfse_avulsa.txt", "boleto_avulso.txt", "cnpj_checksum_invalido.txt"):
            caminho = os.path.join(dados, nome)
            self.assertTrue(os.path.isfile(caminho), f"faltando fixture: {nome}")
            with open(caminho, encoding="utf-8") as f:
                self.assertGreater(len(f.read().strip()), 100, f"fixture vazio: {nome}")


if __name__ == "__main__":
    unittest.main()
