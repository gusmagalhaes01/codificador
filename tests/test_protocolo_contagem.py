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
# O endereço no rodapé ("774 - andar") simula a barra de "774 / 10° andar"
# lida como traço pelo OCR — dígito-traço-letra, no mesmo formato de uma
# linha de unidade (ver test_rodape_nao_vira_unidade: sem o corte por
# "Listando", essa linha entraria na contagem por engano).
# Nomes fictícios de propósito: fixture não guarda nome de morador real.
PROTOCOLO_OK = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Imodata - Condominios e Imoveis "
    "Rua Barata Ribeiro, 774 - andar Copacabana } RJ -22.051-002 "
    "matriz@imodata.net - (21) 3816-7800 Assinatura Entrega "
    "Correio Correio Correio 1155 1 de 1"
)

# Caso real do protocolo de 15 unidades: o OCR duplicou o traço numa linha
# ("702 - - Enny"). RE_UNIDADE_PROTOCOLO usa (?:\s*[-–—])+, que absorve
# quantos traços vierem — a linha ainda casa como uma única unidade, então
# os três números concordam (3/3/3). Não é um caso de divergência: é o caso
# que prova que esse detalhe da regex não derruba a contagem.
PROTOCOLO_TRACO_DUPLO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega Correio Correio Correio 1 de 1"
)

# Falha real de OCR: um "1" reconhecido como "l" minúsculo (401 -> 40l).
# RE_UNIDADE_PROTOCOLO exige \d{1,4} puro, então essa linha não casa — a
# regex de unidade erra sozinha, mas a contagem de "Correio" não depende do
# número da unidade e continua correta.
PROTOCOLO_UNIDADE_COM_LETRA = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 40l - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega "
    "Correio Correio Correio 1 de 1"
)

# Outra falha real de OCR, na coluna Entrega: um "Correio" sai como
# "Correlo" (i minúsculo lido como l). RE_ENTREGA_PROTOCOLO exige a palavra
# exata, então essa ocorrência não conta — a contagem de "Correio" erra
# sozinha, mas a regex de unidade não depende dessa coluna e continua
# correta.
PROTOCOLO_CORREIO_MAL_LIDO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega "
    "Correio Correio Correlo 1 de 1"
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

    def test_texto_none_nao_quebra(self):
        """Justifica o guard `texto = texto or ""`: sem ele, `None.lower()`
        estouraria AttributeError em vez de devolver None como os demais
        casos de "não é protocolo"."""
        self.assertIsNone(app.extrair_dados_protocolo_correio(None))

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
        """O corte `corpo = texto[:listando.start()]` existe porque o
        rodapé de PROTOCOLO_OK tem algo que casaria com RE_UNIDADE_PROTOCOLO
        se entrasse na conta: o endereço "774 - andar" é dígito-traço-letra,
        o mesmo formato de uma linha de unidade. Sem o corte, linhas_contadas
        seria 4 (as 3 unidades reais + essa linha do rodapé), não 3."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["linhas_contadas"], 3)

    def test_traco_duplicado_e_tolerado_pela_regex_de_unidade(self):
        """O traço duplicado que o OCR às vezes produz ("402 - - Beltrano")
        não derruba RE_UNIDADE_PROTOCOLO — o `(?:\\s*[-–—])+` absorve os dois
        traços e a linha ainda casa como uma unidade só. Os três números
        concordam entre si (3/3/3); não há divergência para o "Correio"
        salvar."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_TRACO_DUPLO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["linhas_contadas"], 3)
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

    def test_aceita_quando_regex_de_unidade_falha_mas_entregas_confirma(self):
        """Divergência real: "40l" (letra no lugar do dígito) derruba a
        regex de unidade (2, não 3), mas a contagem de "Correio" não olha
        para o número da unidade e continua batendo com o total (3) —
        é ela quem sustenta a aceitação."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_UNIDADE_COM_LETRA)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["linhas_contadas"], 2)
        self.assertEqual(dados["entregas_contadas"], 3)
        self.assertNotEqual(dados["linhas_contadas"], dados["entregas_contadas"])

        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)
        self.assertEqual(motivo, "")

    def test_aceita_quando_entregas_falha_mas_regex_de_unidade_confirma(self):
        """Divergência real, no outro sentido: "Correlo" derruba a contagem
        de "Correio" (2, não 3), mas a regex de unidade não depende da
        coluna Entrega e continua batendo com o total (3) — é ela quem
        sustenta a aceitação."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_CORREIO_MAL_LIDO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["linhas_contadas"], 3)
        self.assertEqual(dados["entregas_contadas"], 2)
        self.assertNotEqual(dados["linhas_contadas"], dados["entregas_contadas"])

        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)
        self.assertEqual(motivo, "")

    def test_recusa_sem_total_impresso(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_SEM_LISTANDO)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        self.assertIn("total impresso", motivo)

    def test_recusa_quando_nenhum_conferidor_bate(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_DIVERGENTE)
        self.assertEqual(dados["linhas_contadas"], 1)
        self.assertEqual(dados["entregas_contadas"], 1)

        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        # A mensagem precisa informar as contagens divergentes (1 e 1), não
        # só repetir o "3" do total impresso — que apareceria de qualquer
        # forma por vir do "Listando 3".
        self.assertEqual(motivo, "Listando 3, mas foram contadas 1 e 1 unidades")


from decimal import Decimal


class TestValorProtocolo(unittest.TestCase):
    def test_multiplicacao_simples(self):
        self.assertEqual(app.valor_protocolo(15, Decimal("3.85")), Decimal("57.75"))

    def test_aceita_tarifa_float_sem_perder_centavo(self):
        """1.015 como float carrega erro de representação binária
        (1.0150000000000000088...): convertida direto para Decimal (sem
        passar por str primeiro) o excedente somado ao restante do float
        arredonda para BAIXO (1.01). Só Decimal(str(tarifa)) preserva os
        três dígitos exatos e arredonda para CIMA (1.02) — é essa diferença
        que o teste precisa detectar, não só "bate com o valor esperado".
        Com (12, 3.85) do brief original, as duas conversões dão 46.20 e o
        teste passaria mesmo com o bug (Decimal(tarifa) sem str()); troquei
        para (1, 1.015) para o teste realmente falhar se a implementação
        perder o str()."""
        self.assertEqual(app.valor_protocolo(1, 1.015), Decimal("1.02"))

    def test_lote_de_referencia_fecha(self):
        """Os quatro protocolos reais de 24/07/2026, conferidos contra os
        valores escritos à mão no papel."""
        total = sum(app.valor_protocolo(u, Decimal("3.85")) for u in (2, 15, 3, 12))
        self.assertEqual(total, Decimal("123.20"))

    def test_arredonda_meio_centavo_para_cima(self):
        self.assertEqual(app.valor_protocolo(1, Decimal("3.855")), Decimal("3.86"))

    def test_zero_unidades(self):
        self.assertEqual(app.valor_protocolo(0, Decimal("3.85")), Decimal("0.00"))

    def test_formata_em_reais(self):
        self.assertEqual(app.formatar_reais(Decimal("57.75")), "R$ 57,75")
        self.assertEqual(app.formatar_reais(Decimal("1234.50")), "R$ 1.234,50")

    def test_texto_do_carimbo_mostra_a_conta(self):
        texto = app.montar_texto_valor_protocolo(15, Decimal("3.85"), Decimal("57.75"))
        self.assertEqual(texto, "15 un × R$ 3,85 = R$ 57,75")


if __name__ == "__main__":
    unittest.main()
