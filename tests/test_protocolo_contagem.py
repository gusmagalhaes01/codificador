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

# O winocr devolve a página inteira numa linha só e com as colunas fora de
# ordem — os "Correio" da coluna Entrega saem todos no fim, depois do rodapé.
# Nomes fictícios de propósito: fixture não guarda nome de morador real.
PROTOCOLO_OK = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Imodata - Condominios e Imoveis "
    "Rua Barata Ribeiro, 774 / 100 andar Copacabana } RJ -22.051-002 "
    "matriz@imodata.net - (21) 3816-7800 Assinatura Entrega "
    "Correio Correio Correio 1155 1 de 1"
)

# Caso real do protocolo de 15 unidades: o OCR duplicou o traço numa linha
# ("702 - - Enny"), então a regex de unidade conta 2 e o "Correio" conta 3.
PROTOCOLO_TRACO_DUPLO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega Correio Correio Correio 1 de 1"
)

PROTOCOLO_SEM_LISTANDO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva Assinatura Entrega "
    "Correio Correio 1 de 1"
)

PROTOCOLO_DIVERGENTE = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal Listando 3 unidades Assinatura Entrega Correio 1 de 1"
)

NAO_E_PROTOCOLO = "CO-ESTIPULANTE: KLOSTERS CNPJ: 01.195.716/0001-54"


class TestExtracaoProtocolo(unittest.TestCase):
    def test_documento_que_nao_e_protocolo(self):
        self.assertIsNone(app.extrair_dados_protocolo_correio(NAO_E_PROTOCOLO))

    def test_texto_vazio_nao_quebra(self):
        self.assertIsNone(app.extrair_dados_protocolo_correio(""))

    def test_le_codigo_e_condominio_do_cabecalho(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["codigo"], "10004")
        self.assertEqual(dados["condominio"], "W700A KLOSTERS")

    def test_conta_as_tres_fontes(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["linhas_contadas"], 3)
        self.assertEqual(dados["entregas_contadas"], 3)

    def test_rodape_nao_vira_unidade(self):
        """CEP, telefone e o manuscrito lido pelo OCR ficam depois do
        "Listando" e não podem entrar na contagem de linhas."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["linhas_contadas"], 3)

    def test_traco_duplicado_derruba_a_regex_mas_nao_a_contagem(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_TRACO_DUPLO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["entregas_contadas"], 3)

    def test_sem_listando_o_total_e_none(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_SEM_LISTANDO)
        self.assertIsNone(dados["total_impresso"])


class TestConferenciaProtocolo(unittest.TestCase):
    def test_aceita_quando_tudo_bate(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)
        self.assertEqual(motivo, "")

    def test_aceita_quando_so_um_conferidor_bate(self):
        """Traço duplicado: a regex de unidade erra, o "Correio" salva."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_TRACO_DUPLO)
        aceito, _ = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)

    def test_recusa_sem_total_impresso(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_SEM_LISTANDO)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        self.assertIn("total impresso", motivo)

    def test_recusa_quando_nenhum_conferidor_bate(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_DIVERGENTE)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        self.assertIn("3", motivo)


if __name__ == "__main__":
    unittest.main()
