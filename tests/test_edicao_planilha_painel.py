"""
Edição da planilha embutida no painel dos protocolos (aba 3).

A regra que sustenta tudo: o valor e o código não vivem só na planilha — eles
estão CARIMBADOS no papel (bloco do Paybox e carimbo lateral), e o Superlógica
lê o papel por OCR. Editar sem recarimbar criaria divergência silenciosa. Por
isso a validação recusa cedo (código fora do cadastro, unidade zero) e
`aplicar_edicao_na_linha` devolve uma linha NOVA, para o chamador só trocar
depois que o recarimbo deu certo.
"""
import os
import sys
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app
from cadastro_teste import CADASTRO_TESTE

TARIFA = Decimal("3.85")
#  Arquivo | Condomínio | Código | Unidades | Tarifa | Valor | Observação
LINHA = ["doc1600.pdf", "KLOSTERS", "10004", 12, 3.85, 46.20, ""]


class TestColunasEditaveis(unittest.TestCase):
    def test_colunas_somente_leitura_sao_recusadas(self):
        for coluna, nome in ((0, "Arquivo"), (1, "Condomínio"), (4, "Tarifa")):
            with self.subTest(coluna=nome):
                _v, erro = app.validar_edicao_protocolo(coluna, "x", CADASTRO_TESTE)
                self.assertTrue(erro)

    def test_quais_colunas_obrigam_recarimbo(self):
        # Observação é a única editável que não muda o papel.
        self.assertIn(app.COL_CODIGO, app.COLUNAS_QUE_RECARIMBAM)
        self.assertIn(app.COL_UNIDADES, app.COLUNAS_QUE_RECARIMBAM)
        self.assertIn(app.COL_VALOR, app.COLUNAS_QUE_RECARIMBAM)
        self.assertNotIn(app.COL_OBSERVACAO, app.COLUNAS_QUE_RECARIMBAM)


class TestValidarCodigo(unittest.TestCase):
    def test_codigo_do_cadastro_e_aceito(self):
        valor, erro = app.validar_edicao_protocolo(
            app.COL_CODIGO, "10004", CADASTRO_TESTE)
        self.assertEqual((valor, erro), ("10004", ""))

    def test_codigo_fora_do_cadastro_e_recusado_na_hora(self):
        # Aceitar aqui carimbaria o papel com identificação inexistente, e o
        # erro só apareceria na geração das despesas.
        _v, erro = app.validar_edicao_protocolo(
            app.COL_CODIGO, "99999", CADASTRO_TESTE)
        self.assertIn("não está no cadastro", erro)

    def test_codigo_vazio_ou_com_letra(self):
        for texto in ("", "   ", "10004A", "abc"):
            with self.subTest(texto=texto):
                _v, erro = app.validar_edicao_protocolo(
                    app.COL_CODIGO, texto, CADASTRO_TESTE)
                self.assertTrue(erro)


class TestValidarUnidades(unittest.TestCase):
    def test_inteiro_positivo(self):
        valor, erro = app.validar_edicao_protocolo(
            app.COL_UNIDADES, "12", CADASTRO_TESTE, TARIFA)
        self.assertEqual((valor, erro), (12, ""))

    def test_zero_e_recusado(self):
        # Célula VAZIA representa "sem contagem"; 0 diria "entregou zero".
        _v, erro = app.validar_edicao_protocolo(
            app.COL_UNIDADES, "0", CADASTRO_TESTE, TARIFA)
        self.assertIn("maior que zero", erro)

    def test_nao_inteiro(self):
        for texto in ("12,5", "doze", "-3", ""):
            with self.subTest(texto=texto):
                _v, erro = app.validar_edicao_protocolo(
                    app.COL_UNIDADES, texto, CADASTRO_TESTE, TARIFA)
                self.assertTrue(erro)

    def test_sem_tarifa_no_lote_explica_o_caminho(self):
        _v, erro = app.validar_edicao_protocolo(
            app.COL_UNIDADES, "12", CADASTRO_TESTE, None)
        self.assertIn("Edite o valor", erro)


class TestValidarValor(unittest.TestCase):
    def test_formato_brasileiro(self):
        valor, erro = app.validar_edicao_protocolo(
            app.COL_VALOR, "46,20", CADASTRO_TESTE)
        self.assertEqual((valor, erro), (Decimal("46.20"), ""))

    def test_aceita_com_cifrao(self):
        valor, _e = app.validar_edicao_protocolo(
            app.COL_VALOR, "R$ 1.234,56", CADASTRO_TESTE)
        self.assertEqual(valor, Decimal("1234.56"))

    def test_zero_e_negativo_recusados(self):
        for texto in ("0", "0,00", "-5,00"):
            with self.subTest(texto=texto):
                _v, erro = app.validar_edicao_protocolo(
                    app.COL_VALOR, texto, CADASTRO_TESTE)
                self.assertTrue(erro)

    def test_espaco_no_meio_dos_digitos_continua_recusado(self):
        # Herdado de converter_valor_digitado: "3 85" não pode virar 385.
        _v, erro = app.validar_edicao_protocolo(
            app.COL_VALOR, "3 85", CADASTRO_TESTE)
        self.assertTrue(erro)


class TestObservacao(unittest.TestCase):
    def test_texto_livre_e_aceito_inclusive_vazio(self):
        for texto in ("conferido a mão", ""):
            valor, erro = app.validar_edicao_protocolo(
                app.COL_OBSERVACAO, texto, CADASTRO_TESTE)
            self.assertEqual(erro, "")
            self.assertEqual(valor, texto)


class TestAplicarEdicao(unittest.TestCase):
    def test_nao_altera_a_linha_original(self):
        # O chamador só troca depois que o recarimbo deu certo.
        original = list(LINHA)
        app.aplicar_edicao_na_linha(LINHA, app.COL_VALOR, Decimal("99.90"),
                                    CADASTRO_TESTE, TARIFA)
        self.assertEqual(LINHA, original)

    def test_editar_unidades_recalcula_o_valor_pela_tarifa(self):
        nova = app.aplicar_edicao_na_linha(LINHA, app.COL_UNIDADES, 20,
                                            CADASTRO_TESTE, TARIFA)
        self.assertEqual(nova[app.COL_UNIDADES], 20)
        self.assertEqual(nova[app.COL_VALOR], 77.00)   # 20 x 3,85
        self.assertEqual(nova[4], 3.85)

    def test_editar_valor_direto_esvazia_unidades_e_tarifa(self):
        # Sem contagem por trás, a conta deixou de valer — mesma representação
        # que linha_planilha_protocolo usa para valor informado à mão.
        nova = app.aplicar_edicao_na_linha(LINHA, app.COL_VALOR,
                                            Decimal("99.90"), CADASTRO_TESTE, TARIFA)
        self.assertEqual(nova[app.COL_VALOR], 99.90)
        self.assertIsNone(nova[app.COL_UNIDADES])
        self.assertIsNone(nova[4])

    def test_editar_codigo_atualiza_o_nome_do_condominio(self):
        nova = app.aplicar_edicao_na_linha(LINHA, app.COL_CODIGO, "10002",
                                            CADASTRO_TESTE, TARIFA)
        self.assertEqual(nova[app.COL_CODIGO], "10002")
        self.assertEqual(nova[1], "SAN REMO")

    def test_editar_observacao_nao_toca_nos_valores(self):
        nova = app.aplicar_edicao_na_linha(LINHA, app.COL_OBSERVACAO,
                                            "conferido", CADASTRO_TESTE, TARIFA)
        self.assertEqual(nova[app.COL_OBSERVACAO], "conferido")
        self.assertEqual(nova[app.COL_VALOR], LINHA[app.COL_VALOR])
        self.assertEqual(nova[app.COL_UNIDADES], LINHA[app.COL_UNIDADES])


if __name__ == "__main__":
    unittest.main()
