# -*- coding: utf-8 -*-
"""Formato "Meus Correios" (sistema Agile) e modo SEDEX."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

#  Texto como o winocr devolve: página numa linha só, colunas fora de ordem.
#  Arranjo em TABELA (o mais comum nos 6 arquivos de referência).
MEUS_CORREIOS_TABELA = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10852-VILLA BRANCA Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Histórico ENVIO CONTRACHEQUES FUNCIONÁRIOS "
    "Contar. 1 Total: 1"
)

#  Arranjo VERTICAL (mesmo sistema, largura diferente).
MEUS_CORREIOS_VERTICAL = (
    "Correios [1 2704] ID 12704 Data 18/08/2026 Status Gerado "
    "Cód. Condomínio 1 0193-MONACO Qtde. 01 Classificação "
    "590 - EBCT - CORREIO REG / AR Histórico ENVI RESCISAO DE CONTRATO"
)

TELNET = (
    "*** IMODATA * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL DATA: 11/08/2026 "
    "EDF: MENESCAL LJ-01 AP-204 TOTAL ENVIADO PELO CORREIO..: 27 COD.: 1.1122.8)"
)

PROTOCOLO_NOVO = (
    "W700A VILLARS (10005) Protocolo de Recebimento de Documento "
    "Listando 3 unidades 702 - Fulano de Tal Correio"
)


class TestDiscriminadorMeusCorreios(unittest.TestCase):

    def test_reconhece_o_arranjo_em_tabela(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(MEUS_CORREIOS_TABELA),
            "meus_correios")

    def test_reconhece_o_arranjo_vertical(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(MEUS_CORREIOS_VERTICAL),
            "meus_correios")

    def test_nao_confunde_com_telnet(self):
        self.assertEqual(logica.classificar_formato_protocolo(TELNET), "telnet")

    def test_nao_confunde_com_o_protocolo_novo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(PROTOCOLO_NOVO), "novo")

    def test_cod_do_telnet_nao_dispara_o_marcador(self):
        #  O telnet tem "COD.: 1.1122.8)". O marcador do Meus Correios exige
        #  "COND" logo depois de "Cód", então não pode casar aqui — se casasse,
        #  todo telnet viraria ambíguo e o lote inteiro cairia em pendente.
        self.assertNotEqual(
            logica.classificar_formato_protocolo(TELNET), "ambiguo")

    def test_misturado_com_telnet_vira_ambiguo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(TELNET + " " + MEUS_CORREIOS_TABELA),
            "ambiguo")

    def test_misturado_com_o_novo_vira_ambiguo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(
                PROTOCOLO_NOVO + " " + MEUS_CORREIOS_TABELA),
            "ambiguo")

    def test_texto_sem_marcador_nenhum(self):
        self.assertIsNone(logica.classificar_formato_protocolo("boleto qualquer"))

    def test_formato_do_protocolo_funde_meus_correios_em_none(self):
        #  O wrapper antigo tem contrato de dois formatos; quem precisa do
        #  terceiro usa classificar_formato_protocolo diretamente.
        self.assertIsNone(logica.formato_do_protocolo(MEUS_CORREIOS_TABELA))


if __name__ == "__main__":
    unittest.main()
