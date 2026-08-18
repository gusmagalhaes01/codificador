import os
import shutil
import sys
import tempfile
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

DADOS_CADASTRADO = {
    "codigo": "10004", "condominio": "W700A KLOSTERS",
    "total_impresso": 15, "linhas_contadas": 15, "entregas_contadas": 15,
}
DADOS_NAO_CADASTRADO = {
    "codigo": "99999", "condominio": "W700A DESCONHECIDO",
    "total_impresso": 3, "linhas_contadas": 3, "entregas_contadas": 3,
}


class TestLinhaPlanilhaProtocolo(unittest.TestCase):
    def test_codigo_cadastrado_usa_o_nome_do_cadastro(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=15)
        self.assertEqual(linha[1], "KLOSTERS")
        self.assertEqual(linha[2], "10004")
        self.assertEqual(linha[3], 15)
        self.assertEqual(linha[5], 57.75)
        self.assertEqual(linha[6], "")

    def test_codigo_nao_cadastrado_mantem_o_nome_do_documento(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=3)
        self.assertEqual(linha[1], "W700A DESCONHECIDO")
        self.assertEqual(linha[6], "Código não cadastrado")
        self.assertEqual(linha[5], 11.55)

    def test_pendente_deixa_valores_vazios_nunca_zero(self):
        """0 significaria "entregou zero unidades" — mesma regra das
        retenções federais na planilha das NFS-e."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            observacao="Listando 15, mas foram contadas 14 e 14 unidades")
        self.assertIsNone(linha[3])
        self.assertIsNone(linha[4])
        self.assertIsNone(linha[5])
        self.assertIn("Listando 15", linha[6])

    def test_nao_e_protocolo(self):
        linha = app.linha_planilha_protocolo(
            "outro.pdf", None, CADASTRO_TESTE,
            observacao="Não é um protocolo dos Correios")
        self.assertEqual(linha[0], "outro.pdf")
        self.assertEqual(len(linha), len(app.COLUNAS_PROTOCOLO))
        # Nenhum campo numérico pode restar com valor — nem código/nome, que
        # aqui não fazem sentido nenhum (documento não é sequer um protocolo).
        self.assertTrue(all(c is None for c in linha[3:6]))
        self.assertEqual(linha[6], "Não é um protocolo dos Correios")

    def test_codigo_nao_cadastrado_concatena_com_observacao_existente(self):
        """Achado da revisão: `if not registro and not observacao` fazia
        "Código não cadastrado" desaparecer quando já havia outra observação
        (ex.: motivo de contagem recusada) — um protocolo pode estar pendente
        E com código fora do cadastro ao mesmo tempo, e as duas informações
        precisam sobreviver na planilha."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
            observacao="Listando 3, mas foram contadas 2 e 2 unidades")
        self.assertIn("Listando 3", linha[6])
        self.assertIn("Código não cadastrado", linha[6])

    def test_codigo_nao_lido_distingue_de_codigo_nao_cadastrado(self):
        """Achado da revisão: quando o código não foi lido do documento
        (`codigo` vazio/None), a mensagem não pode dizer "não cadastrado" —
        isso manda a pessoa procurar no cadastro um código que na verdade
        nunca existiu, quando o problema é a leitura, não o cadastro."""
        dados_sem_codigo = {
            "codigo": None, "condominio": "W700A SEM CODIGO",
            "total_impresso": 3, "linhas_contadas": 3, "entregas_contadas": 3,
        }
        linha = app.linha_planilha_protocolo(
            "p.pdf", dados_sem_codigo, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=3)
        self.assertNotEqual(linha[6], "Código não cadastrado")
        self.assertIn("não identificado", linha[6].lower())

    def test_valor_e_decimal_convertido_pra_float_nunca_texto(self):
        """Reforço: dinheiro tem que sobreviver como número, não como str —
        uma implementação que devolvesse str(valor) passaria despercebida
        pelos asserts de igualdade acima (Python compara "57.75" == 57.75
        como False, mas um regressão sutil, ex. formatar_reais(valor), não
        seria pega pelos testes originais do brief)."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=15)
        self.assertIsInstance(linha[3], int)
        self.assertIsInstance(linha[4], float)
        self.assertIsInstance(linha[5], float)


class TestSalvarPlanilhaProtocolo(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.destino = os.path.join(self.pasta, "protocolos.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_grava_numeros_e_linha_de_total(self):
        from openpyxl import load_workbook
        linhas = [
            app.linha_planilha_protocolo("a.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                         Decimal("3.85"), 15),
            app.linha_planilha_protocolo("b.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
                                         Decimal("3.85"), 3),
            app.linha_planilha_protocolo("c.pdf", None, CADASTRO_TESTE,
                                         observacao="Não é um protocolo dos Correios"),
        ]
        app.salvar_planilha_protocolo(self.destino, linhas)

        sheet = load_workbook(self.destino).active
        self.assertEqual(sheet.cell(row=1, column=1).value, "Arquivo")
        self.assertEqual([c.value for c in sheet[1]], [c[0] for c in app.COLUNAS_PROTOCOLO])
        # Valor gravado como número, não texto (dá para somar no Excel)
        self.assertEqual(sheet.cell(row=2, column=6).value, 57.75)
        self.assertIsInstance(sheet.cell(row=2, column=6).value, float)
        # A linha "não é protocolo" tem que ficar mesmo vazia na planilha —
        # não pode ter virado 0 nem string vazia == 0 num somatório futuro.
        self.assertIsNone(sheet.cell(row=4, column=4).value)
        self.assertIsNone(sheet.cell(row=4, column=6).value)
        # Linha de total logo abaixo dos dados
        linha_total = sheet.max_row
        self.assertEqual(sheet.cell(row=linha_total, column=1).value, "TOTAL")
        self.assertEqual(sheet.cell(row=linha_total, column=4).value, 18)
        self.assertEqual(sheet.cell(row=linha_total, column=6).value, 69.30)
        # Reforço: total também tem que ser número de verdade (dá pra somar
        # de novo/usar em fórmula), não uma string "69.3" parecida.
        self.assertIsInstance(sheet.cell(row=linha_total, column=4).value, (int, float))
        self.assertIsInstance(sheet.cell(row=linha_total, column=6).value, (int, float))
        self.assertEqual(sheet.freeze_panes, "A2")
        self.assertIsNotNone(sheet.auto_filter.ref)

    def test_planilha_vazia_nao_quebra(self):
        from openpyxl import load_workbook
        app.salvar_planilha_protocolo(self.destino, [])
        sheet = load_workbook(self.destino).active
        self.assertEqual(sheet.cell(row=sheet.max_row, column=1).value, "TOTAL")
        self.assertEqual(sheet.cell(row=sheet.max_row, column=6).value, 0)


class TestValorManual(unittest.TestCase):
    def test_valor_manual_preenche_valor_e_deixa_unidades_vazias(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("300.30"))
        self.assertEqual(linha[1], "KLOSTERS")
        self.assertIsNone(linha[3])          # Unidades
        self.assertIsNone(linha[4])          # Tarifa
        self.assertEqual(linha[5], 300.30)   # Valor
        self.assertEqual(linha[6], "Valor informado manualmente")

    def test_valor_manual_nunca_grava_zero(self):
        """0 significaria "entregou zero unidades" — mesma regra das
        pendências e das retenções federais na planilha das NFS-e."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("10.00"))
        self.assertNotEqual(linha[3], 0)
        self.assertNotEqual(linha[4], 0)

    def test_valor_manual_grava_como_numero_nao_texto(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("300.30"))
        self.assertIsInstance(linha[5], float)

    def test_codigo_nao_cadastrado_mantem_os_dois_avisos(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("11.55"))
        self.assertIn("Valor informado manualmente", linha[6])
        self.assertIn("Código não cadastrado", linha[6])

    def test_total_soma_manuais_junto_com_calculados(self):
        from openpyxl import load_workbook
        pasta = tempfile.mkdtemp()
        try:
            destino = os.path.join(pasta, "p.xlsx")
            linhas = [
                app.linha_planilha_protocolo("a.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                             Decimal("3.85"), 15),
                app.linha_planilha_protocolo("b.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                             valor_manual=Decimal("300.30")),
            ]
            app.salvar_planilha_protocolo(destino, linhas)
            sheet = load_workbook(destino).active
            ultima = sheet.max_row
            self.assertEqual(sheet.cell(row=ultima, column=1).value, "TOTAL")
            #  Só as unidades efetivamente contadas entram no total de unidades
            self.assertEqual(sheet.cell(row=ultima, column=4).value, 15)
            #  57,75 calculado + 300,30 informado
            self.assertEqual(sheet.cell(row=ultima, column=6).value, 358.05)
        finally:
            shutil.rmtree(pasta, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
