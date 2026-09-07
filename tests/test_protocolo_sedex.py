# -*- coding: utf-8 -*-
"""Formato "Meus Correios" (sistema Agile) e modo SEDEX."""
import os
import sys
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica
from cadastro_teste import CADASTRO_TESTE as CADASTRO

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

#  Um código que existe no cadastro de teste, com o nome que ele tem lá.
#  KLOSTERS é 10004 em tests/cadastro_teste.py.
MC_KLOSTERS = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10004-KLOSTERS Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Histórico ENVIO DE DOCUMENTOS Contar. 1 Total: 1"
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


class TestLeitorMeusCorreios(unittest.TestCase):

    def test_extrai_codigo_e_resolve_o_condominio(self):
        d = logica.extrair_dados_meus_correios(MC_KLOSTERS, CADASTRO)
        self.assertEqual(d["formato"], "meus_correios")
        self.assertEqual(d["codigo"], "10004")
        self.assertEqual(d["condominio"], "KLOSTERS")
        self.assertTrue(d["nome_confere"])

    def test_espaco_no_meio_dos_digitos(self):
        #  O OCR devolve "1 0004-KLOSTERS" com espaço no meio do número.
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "1 0004-KLOSTERS")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")

    def test_espaco_depois_do_traco(self):
        #  Medido: "10520- SENADOR LEITE OITICICA".
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "10004- KLOSTERS")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")
        self.assertTrue(d["nome_confere"])

    def test_codigo_fora_do_cadastro(self):
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "99999-INEXISTENTE")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertIsNone(d["codigo"])
        self.assertFalse(d["nome_confere"])

    def test_nome_divergente_nao_confere_mas_o_codigo_vale(self):
        #  Código bom, nome ilegível: o código manda (é o que o cadastro
        #  resolve), e nome_confere=False sinaliza a dúvida para a tela.
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "10004-XXXXXXX")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")
        self.assertFalse(d["nome_confere"])

    def test_servico_sedex(self):
        d = logica.extrair_dados_meus_correios(MC_KLOSTERS, CADASTRO)
        self.assertEqual(d["servico"], "824 - EBCT - SEDEX")

    def test_servico_registrado_corta_no_historico(self):
        texto = MC_KLOSTERS.replace(
            "824 - EBCT - SEDEX Histórico",
            "590 - EBCT- CORREIO REG / AR Histórico")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "590 - EBCT- CORREIO REG / AR")

    def test_servico_corta_em_ID_quando_o_ocr_embaralha(self):
        #  Medido no telnet-003: depois da Classificação vem "ID 12699
        #  Contar: 1 Data ...", não "Histórico".
        texto = MC_KLOSTERS.replace(
            "824 - EBCT - SEDEX Histórico ENVIO DE DOCUMENTOS",
            "590 - EBCT- CORREIO REG / AR ID 12699 Contar: 1 Data 13/08/2026")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "590 - EBCT- CORREIO REG / AR")

    def test_sem_classificacao_o_servico_fica_vazio(self):
        #  Best-effort: o serviço é informação de apoio, nunca bloqueia nada.
        texto = MC_KLOSTERS.replace(
            "Classificação 824 - EBCT - SEDEX ", "")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "")
        self.assertEqual(d["codigo"], "10004")

    def test_texto_vazio_nao_estoura(self):
        d = logica.extrair_dados_meus_correios("", CADASTRO)
        self.assertIsNone(d["codigo"])
        self.assertEqual(d["servico"], "")


if __name__ == "__main__":
    unittest.main()
