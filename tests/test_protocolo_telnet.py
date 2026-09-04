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

#  Texto como o winocr devolve de verdade: página inteira numa linha só e
#  com as colunas fora de ordem. Os erros de leitura aqui são os REAIS
#  observados nos sete arquivos de referência — "EWIADO" no lugar de
#  "ENVIADO" e "D ÀS" no lugar de "COD.". Nomes de condomínio são os do
#  cadastro de teste.
TELNET_OK = (
    "IMODÀTÀ * PROTOCOLO CORRESPONDENCIA c CORREIO NORMAL DATA: 11/08/2026 "
    "*** EDF: REF ; LJ-01 LJ-25 AP-308 KLOSTERS 07/2026 R TESTE 100 "
    "PB608111.121 LJ-02 LJ-08 AP-204 AP-207 "
    "TOTAL EWIADO PELO CORREIO 3 o COD.. D ÀS 1.0004.8)"
)

#  Protocolo novo, no formato que a aba 3 já lê hoje.
PROTOCOLO_NOVO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega Correio Correio Correio 1 de 1"
)

#  "Meus Correios": terceiro formato que aparece na mesma pasta e que este
#  spec deliberadamente NÃO trata (SEDEX, cobrado por outro critério).
#  Não pode ser confundido com nenhum dos dois.
MEUS_CORREIOS = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10004-KLOSTERS Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Total: 1"
)


class TestFormatoDoProtocolo(unittest.TestCase):
    def test_reconhece_o_telnet(self):
        self.assertEqual(app.formato_do_protocolo(TELNET_OK), "telnet")

    def test_reconhece_o_protocolo_novo(self):
        self.assertEqual(app.formato_do_protocolo(PROTOCOLO_NOVO), "novo")

    def test_meus_correios_nao_e_nenhum_dos_dois(self):
        self.assertIsNone(app.formato_do_protocolo(MEUS_CORREIOS))

    def test_texto_vazio_e_none(self):
        self.assertIsNone(app.formato_do_protocolo(""))
        self.assertIsNone(app.formato_do_protocolo(None))

    def test_acento_e_caixa_nao_atrapalham(self):
        #  O OCR ora come o acento de "CORRESPONDÊNCIA", ora não.
        self.assertEqual(
            app.formato_do_protocolo("protocolo correspondência correio"),
            "telnet")

    def test_os_dois_marcadores_juntos_viram_ambiguo(self):
        #  Pior caso possível: aplicar a extração errada produz resultado
        #  plausível pelo motivo errado, e a conferência humana não tem o
        #  que estranhar na tela. Ambíguo vira None -> pendente.
        self.assertIsNone(app.formato_do_protocolo(TELNET_OK + " " + PROTOCOLO_NOVO))


class TestExtracaoDeUmaLeitura(unittest.TestCase):
    def test_codigo_sai_pelo_formato_mesmo_com_o_rotulo_corrompido(self):
        #  "D ÀS" é o que o OCR devolveu no lugar de "COD." num arquivo
        #  real. Ancorado no rótulo o código saía em 4 de 7; pelo formato,
        #  em 7 de 7.
        dados = app.extrair_dados_protocolo_telnet(TELNET_OK)
        self.assertIn("10004", dados["codigos"])

    def test_total_lido_com_ENVIADO_corrompido(self):
        #  O OCR devolve "EWIADO" e "EFvTIADO"; "PELO CORREIO" nunca falhou.
        dados = app.extrair_dados_protocolo_telnet(TELNET_OK)
        self.assertEqual(dados["total"], 3)

    def test_total_com_ENVIADO_escrito_de_outro_jeito(self):
        texto = "TOTAL EFvTIADO PELO CORREIO 12 o COD. . 1.0004.7)"
        self.assertEqual(app.extrair_dados_protocolo_telnet(texto)["total"], 12)

    def test_folga_curta_nao_captura_numero_distante(self):
        #  Caso real: com folga de 20 não-dígitos o regex pulava o número
        #  certo e capturava outro do canto da folha -- leu 38 no lugar de
        #  3, num condomínio de três unidades (R$ 146,30 em vez de R$ 11,55).
        texto = "TOTAL ENVIADO PELO CORREIO VALOR VAL x R EURICO 38"
        self.assertIsNone(app.extrair_dados_protocolo_telnet(texto)["total"])

    def test_o_digito_depois_do_codigo_e_ignorado(self):
        #  "1.1122.8)" -> o ".8)" não faz parte do código.
        dados = app.extrair_dados_protocolo_telnet("COD.: 1.1122.8)")
        self.assertEqual(dados["codigos"], ["11122"])

    def test_virgula_no_lugar_do_ponto(self):
        dados = app.extrair_dados_protocolo_telnet("COD.: 1,1122.8)")
        self.assertEqual(dados["codigos"], ["11122"])

    def test_sem_codigo_nem_total_devolve_vazio(self):
        dados = app.extrair_dados_protocolo_telnet("texto qualquer sem nada")
        self.assertEqual(dados["codigos"], [])
        self.assertIsNone(dados["total"])

    def test_texto_vazio_nao_estoura(self):
        dados = app.extrair_dados_protocolo_telnet(None)
        self.assertEqual(dados["codigos"], [])
        self.assertIsNone(dados["total"])


if __name__ == "__main__":
    unittest.main()
