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


class TestClassificarFormatoProtocolo(unittest.TestCase):
    """
    `classificar_formato_protocolo` é a versão que NÃO funde "nenhum
    marcador" com "os dois juntos" — a interface precisa dessa distinção
    para tratar o ambíguo como pendente com observação própria, em vez de
    deixá-lo cair no caminho do protocolo novo (achado I1 da revisão).
    """

    def test_telnet(self):
        self.assertEqual(app.classificar_formato_protocolo(TELNET_OK), "telnet")

    def test_novo(self):
        self.assertEqual(app.classificar_formato_protocolo(PROTOCOLO_NOVO), "novo")

    def test_nenhum_marcador_e_none_nao_ambiguo(self):
        self.assertIsNone(app.classificar_formato_protocolo("boleto qualquer"))
        self.assertIsNone(app.classificar_formato_protocolo(""))
        self.assertIsNone(app.classificar_formato_protocolo(None))

    def test_os_dois_marcadores_juntos_sao_ambiguo_nao_none(self):
        #  É exatamente esta distinção que `formato_do_protocolo` esconde.
        self.assertEqual(
            app.classificar_formato_protocolo(TELNET_OK + " " + PROTOCOLO_NOVO),
            "ambiguo")

    def test_consistente_com_formato_do_protocolo_nos_casos_nao_ambiguos(self):
        #  formato_do_protocolo funde "ambiguo" E "meus_correios" em None
        #  porque seu contrato é de apenas dois formatos (telnet e novo).
        for texto in (TELNET_OK, PROTOCOLO_NOVO, "", None):
            resultado_detalhado = app.classificar_formato_protocolo(texto)
            esperado = None if resultado_detalhado == "ambiguo" else resultado_detalhado
            self.assertEqual(app.formato_do_protocolo(texto), esperado)
        #  meus_correios é um caso especial: classificar retorna "meus_correios",
        #  mas formato_do_protocolo funde em None.
        self.assertEqual(app.classificar_formato_protocolo(MEUS_CORREIOS), "meus_correios")
        self.assertIsNone(app.formato_do_protocolo(MEUS_CORREIOS))


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


from cadastro_teste import CADASTRO_TESTE as CADASTRO


def _leitura(codigo_pontuado, total_texto, nome=""):
    """Monta o texto de uma leitura de OCR do telnet."""
    return (f"IMODÀTÀ * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL EDF: {nome} "
            f"AP-101 AP-102 TOTAL EWIADO PELO CORREIO {total_texto} o COD.. "
            f"D ÀS {codigo_pontuado})")


class TestVotacao(unittest.TestCase):
    def test_maioria_corrige_a_leitura_ruim(self):
        #  Caso real do arquivo 011: as três leituras deram 38, 3 e 3. Sem
        #  votar, "R$ 146,30" chegaria à tela num condomínio de três
        #  unidades. Este é o teste que justifica a votação existir.
        leituras = [_leitura("1.0004.8", "38", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["codigo"], "10004")

    def test_total_com_uma_leitura_so_nao_basta(self):
        #  O total vira dinheiro; uma leitura solitária não sustenta isso.
        leituras = [_leitura("1.0004.8", "7", "KLOSTERS"),
                    _leitura("1.0004.8", "", "KLOSTERS"),
                    _leitura("1.0004.8", "", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["total_impresso"])

    def test_tres_totais_diferentes_nao_elegem_nenhum(self):
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "5", "KLOSTERS"),
                    _leitura("1.0004.8", "9", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["total_impresso"])

    def test_codigo_aceita_um_voto_so(self):
        #  Ao contrário do total, o código não vira valor sozinho — ele
        #  ainda passa pela conferência do nome, que marca a linha.
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("", "3", "KLOSTERS"),
                    _leitura("", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertEqual(dados["codigo"], "10004")

    def test_empate_entre_codigos_nao_escolhe_nenhum(self):
        #  Escolher por ordem de chegada seria arbitrário, e um dígito
        #  errado tende a dar OUTRO condomínio real.
        leituras = [_leitura("1.0004.8", "3", ""),
                    _leitura("1.0002.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["codigo"])

    def test_codigo_fora_do_cadastro_e_descartado(self):
        leituras = [_leitura("1.9999.8", "3", ""),
                    _leitura("1.9999.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["codigo"])

    def test_sem_leitura_nenhuma_devolve_none(self):
        self.assertIsNone(app.apurar_protocolo_telnet([], CADASTRO))
        self.assertIsNone(app.apurar_protocolo_telnet(["", None], CADASTRO))


class TestConferenciaDoNome(unittest.TestCase):
    def test_nome_no_texto_confirma_o_codigo(self):
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertTrue(dados["nome_confere"])
        self.assertEqual(dados["condominio"], "KLOSTERS")

    def test_nome_ausente_nao_confirma(self):
        leituras = [_leitura("1.0004.8", "3", ""),
                    _leitura("1.0004.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertFalse(dados["nome_confere"])

    def test_nome_longe_do_rotulo_ainda_confirma(self):
        #  O winocr devolve a página numa linha só com as colunas fora de
        #  ordem: o "EDF:" e o valor não ficam adjacentes. Ancorado no
        #  rótulo o nome saía em 1 de 7; varrendo o texto, em 6 de 7.
        texto = ("IMODÀTÀ * PROTOCOLO CORRESPONDENCIA EDF : REF; AP-104 "
                 "TOTAL KLOSTERS 07/2026 PELO CORREIO 3 o COD.. 1.0004.8)")
        dados = app.apurar_protocolo_telnet([texto, texto], CADASTRO)
        self.assertTrue(dados["nome_confere"])


class TestAceiteTelnet(unittest.TestCase):
    def test_tudo_concordando_aceita_sem_observacao(self):
        dados = {"codigo": "10004", "total_impresso": 3, "nome_confere": True}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertTrue(aceito)
        self.assertEqual(observacao, "")

    def test_nome_nao_confirmado_PREENCHE_e_marca(self):
        #  Deliberadamente mais frouxo que o protocolo novo: o usuário
        #  confere imagem por imagem. Uma versão anterior barrava aqui e
        #  mandava para pendente um arquivo cujo código e total estavam
        #  ambos corretos.
        dados = {"codigo": "10004", "total_impresso": 3, "nome_confere": False}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertTrue(aceito)
        self.assertEqual(observacao, "Nome não confirmado")

    def test_sem_total_nao_aceita(self):
        dados = {"codigo": "10004", "total_impresso": None, "nome_confere": True}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertFalse(aceito)
        self.assertIn("TOTAL ENVIADO PELO CORREIO", observacao)

    def test_sem_codigo_nao_aceita(self):
        dados = {"codigo": None, "total_impresso": 3, "nome_confere": False}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertFalse(aceito)
        self.assertIn("Código", observacao)

    def test_mensagem_de_codigo_ausente_e_neutra_e_nao_contradiz_a_planilha(self):
        #  M1 da revisão: a frase antiga ("não encontrado no cadastro")
        #  afirmava uma causa específica (código lido, mas fora do
        #  cadastro) que não é a única que cai neste `if` — sem código
        #  algum lido e empate entre dois códigos caem aqui também, e
        #  `linha_planilha_protocolo` acrescenta a PRÓPRIA "Código não
        #  identificado no documento" logo depois. As duas juntas não podem
        #  se contradizer na mesma célula.
        dados = {"codigo": None, "total_impresso": 3, "nome_confere": False}
        _, observacao = app.conferir_contagem_telnet(dados)
        self.assertNotIn("cadastro", observacao.lower())
        #  `dados` com `codigo=None` (não `None` puro) é o que
        #  `apurar_protocolo_telnet` de fato devolve nos três casos que
        #  caem neste `if` — é esse dict que segue até
        #  `linha_planilha_protocolo`, que então acrescenta a própria
        #  observação de código ausente.
        linha = app.linha_planilha_protocolo(
            "arquivo.pdf", {"codigo": None, "condominio": ""}, CADASTRO,
            tarifa=None, unidades=None, observacao=observacao)
        observacao_final = linha[-1]
        self.assertNotIn("encontrado no cadastro", observacao_final)
        self.assertIn("Código não identificado no documento", observacao_final)

    def test_dados_none_nao_estoura(self):
        aceito, observacao = app.conferir_contagem_telnet(None)
        self.assertFalse(aceito)
        self.assertTrue(observacao)


if __name__ == "__main__":
    unittest.main()
