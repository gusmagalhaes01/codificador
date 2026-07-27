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

FED = "35315360000167"   # FedCorp (emitente/intermediário)
FF = "13736666000154"    # F&F (emitente)
SAN_REMO = "08578541000103"
KLOSTERS = "01195716000154"
MANHATTAN = "07448975000126"
LAGO_MAGGIORE = "07945453000130"
VILLE = "29273778000156"
IMODATA_INTERM = "12184361000114"


def ler(nome):
    caminho = os.path.join(_AQUI, "dados", nome)
    with open(caminho, encoding="utf-8") as f:
        return f.read()


class TestExtracaoCnpj(unittest.TestCase):
    def test_co_estipulante_com_hifen(self):
        # recibo FedCorp com "08.578.541-0001-03" (hífen no lugar da barra)
        texto = ler("fedcorp_recibo_hifen.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertEqual(cands, [SAN_REMO])

    def test_intermediarios_nunca_aparecem(self):
        texto = ler("fedcorp_recibo_hifen.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertNotIn(FED, cands)
        self.assertNotIn(IMODATA_INTERM, cands)

    def test_nfse_ff_isola_klosters(self):
        # com o emitente F&F excluído, sobra só o tomador Klosters
        texto = ler("nfse_ff.txt")
        cands = app.extrair_cnpj_tomador(texto, FF)
        self.assertEqual(cands, [KLOSTERS])

    def test_nfse_imodata_devolve_os_dois(self):
        texto = ler("nfse_imodata.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertEqual(len(cands), 2)
        self.assertIn(MANHATTAN, cands)

    def test_nfse_avulsa_inclui_tomador(self):
        texto = ler("nfse_avulsa.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertIn(LAGO_MAGGIORE, cands)

    def test_boleto_inclui_pagador(self):
        texto = ler("boleto_avulso.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertIn(VILLE, cands)

    def test_checksum_invalido_rejeitado(self):
        # documento cujo CNPJ do tomador tem dígito verificador errado (NF-927)
        texto = ler("cnpj_checksum_invalido.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertEqual(cands, [])


if __name__ == "__main__":
    unittest.main()
