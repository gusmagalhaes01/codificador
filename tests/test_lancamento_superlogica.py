# -*- coding: utf-8 -*-
"""Lançamento no Superlógica direto da extração (aba 2)."""
import datetime
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

#  Texto nativo como o pypdf devolve um DARF do DCTFWeb (sintético, mesmo
#  layout das guias reais). O vencimento aparece três vezes: no cabeçalho,
#  na composição e no recibo do rodapé.
DARF_OK = (
    "Documento de Arrecadação\nde Receitas Federais\n"
    "Período de Apuração Data de Vencimento Número do Documento\n"
    "07.16.26261.0000000-0 Pagar este documento até\n"
    "18/09/2026Observações\n"
    "Valor Total do Documento\n58,51\n"
    "5952 RET DE CONTRIBUICOES PAGT PJ A PJ DE DIR PRIV 58,51 58,51\n"
    "PA:08/2026 Vencimento:18/09/2026\n"
    "SENDA (Versão:1.5.10) 18/09/2026 09:46:491 1Página: /\n"
    "Número: 07.16.26261.0000000-0\n"
    "Pagar até: 18/09/2026\n"
    "Valor: 58,51\n"
)


class TestVencimentoDoDarf(unittest.TestCase):

    def test_tres_ocorrencias_iguais_dao_a_data(self):
        self.assertEqual(app.vencimento_do_darf(DARF_OK), datetime.date(2026, 9, 18))

    def test_ocorrencias_divergentes_nao_escolhem_nenhuma(self):
        texto = DARF_OK.replace("Pagar até: 18/09/2026", "Pagar até: 19/09/2026")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_uma_ocorrencia_so_nao_basta(self):
        #  A conferência existe porque a data se repete; com uma só, não há
        #  com o que conferir.
        texto = ("Pagar até: 18/09/2026\n"
                 "5952 RET DE CONTRIBUICOES 58,51\n")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_nenhuma_ocorrencia(self):
        self.assertIsNone(app.vencimento_do_darf("Documento de Arrecadação sem data"))

    def test_data_que_nao_existe_e_recusada(self):
        texto = DARF_OK.replace("18/09/2026", "31/02/2026")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_carimbo_de_emissao_nao_conta(self):
        #  "SENDA ... 18/09/2026 09:46" é quando a guia foi emitida, sem
        #  rótulo de vencimento. Se contasse, trocar a data de emissão
        #  mudaria o resultado.
        texto = DARF_OK.replace("SENDA (Versão:1.5.10) 18/09/2026",
                                "SENDA (Versão:1.5.10) 01/01/2026")
        self.assertEqual(app.vencimento_do_darf(texto), datetime.date(2026, 9, 18))

    def test_multa_apos_vencimento_do_boleto_nao_casa(self):
        texto = "APOS VENCIMENTO MULTA DE 2% e JUROS DE 1%\n10/09/2026\n"
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_texto_vazio_ou_none(self):
        self.assertIsNone(app.vencimento_do_darf(""))
        self.assertIsNone(app.vencimento_do_darf(None))


if __name__ == "__main__":
    unittest.main()
