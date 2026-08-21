import os
import shutil
import sys
import tempfile
import unittest
from decimal import Decimal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app
from cadastro_teste import CADASTRO_TESTE

KLOSTERS = {"arquivo": "a.pdf", "codigo": "10004", "condominio": "10004 KLOSTERS",
            "valor": Decimal("57.75")}
SAN_REMO = {"arquivo": "b.pdf", "codigo": "10002", "condominio": "10002 SAN REMO",
            "valor": Decimal("11.55")}
#  Cadastrado, mas sem ID no Superlógica — 10 dos 764 reais estão assim.
LAGO = {"arquivo": "c.pdf", "codigo": "10590", "condominio": "10590 LAGO MAGGIORE",
        "valor": Decimal("7.70")}
#  Resolvido à mão depois de "não foi possível ler": não tem código nenhum.
SEM_CODIGO = {"arquivo": "d.pdf", "codigo": None, "condominio": "",
              "valor": Decimal("300.30")}


def resultado_com(processados=(), pendentes=(), ignorados=()):
    return {"processados": list(processados), "pendentes": list(pendentes),
            "ignorados": list(ignorados)}

#  Modelo sintético no formato do arquivo real do Superlógica: cabeçalho na
#  linha 1 e uma linha de exemplo na 2, com os campos que se repetem
#  preenchidos e condomínio/valor em branco. Nomes de coluna com acento e
#  caixa mista de propósito — é assim que vêm do Superlógica.
CABECALHO_MODELO = [
    "condomínio", "vencimento", "fornecedor", "conta_categoria",
    "valor", "forma_de_pagamento", "chave",
]
MOLDE_MODELO = [
    None, None, "DINAMICA SERVICOS POSTAIS", "2.4.6 Correios - Postagem Simples",
    None, "Trans. bancária", 46,
]


def montar_modelo(caminho, cabecalho=None, molde=None, linhas_extras=0):
    """Escreve um modelo sintético no disco e devolve o caminho."""
    wb = Workbook()
    sheet = wb.active
    sheet.append(cabecalho if cabecalho is not None else CABECALHO_MODELO)
    for celula in sheet[1]:
        celula.font = Font(bold=True)
    sheet.append(molde if molde is not None else MOLDE_MODELO)
    for _ in range(linhas_extras):
        sheet.append(molde if molde is not None else MOLDE_MODELO)
    wb.save(caminho)
    return caminho


class TestGerarPlanilhaDespesas(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.modelo = montar_modelo(os.path.join(self.pasta, "modelo.xlsx"))
        self.saida = os.path.join(self.pasta, "despesas.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def gerar(self, lancamentos):
        app.gerar_planilha_despesas(self.modelo, self.saida, lancamentos)
        return load_workbook(self.saida).active

    def test_uma_linha_por_lancamento_na_ordem_dada(self):
        sheet = self.gerar([("45", Decimal("7.70")),
                            ("496", Decimal("138.60")),
                            ("496", Decimal("119.35"))])
        #  Cabeçalho + 3 lançamentos, sem sobra
        self.assertEqual(sheet.max_row, 4)
        self.assertEqual([sheet.cell(row=l, column=1).value for l in (2, 3, 4)],
                         ["45", "496", "496"])
        self.assertEqual([sheet.cell(row=l, column=5).value for l in (2, 3, 4)],
                         [7.70, 138.60, 119.35])

    def test_campos_do_molde_repetem_em_todas_as_linhas(self):
        sheet = self.gerar([("45", Decimal("7.70")), ("48", Decimal("57.75"))])
        for linha in (2, 3):
            self.assertEqual(sheet.cell(row=linha, column=3).value,
                             "DINAMICA SERVICOS POSTAIS")
            self.assertEqual(sheet.cell(row=linha, column=4).value,
                             "2.4.6 Correios - Postagem Simples")
            self.assertEqual(sheet.cell(row=linha, column=6).value, "Trans. bancária")
            self.assertEqual(sheet.cell(row=linha, column=7).value, 46)

    def test_linha_molde_nao_sobra_em_branco(self):
        """A linha 2 do modelo é o molde e vira a primeira linha real — o
        arquivo final não pode ter nenhuma linha com condomínio/valor vazios."""
        sheet = self.gerar([("45", Decimal("7.70"))])
        self.assertEqual(sheet.max_row, 2)
        self.assertEqual(sheet.cell(row=2, column=1).value, "45")
        self.assertEqual(sheet.cell(row=2, column=5).value, 7.70)

    def test_linhas_extras_do_modelo_nao_sobram(self):
        """Modelo salvo com várias linhas de exemplo não pode deixar resto."""
        modelo = montar_modelo(os.path.join(self.pasta, "modelo3.xlsx"),
                               linhas_extras=4)
        app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        sheet = load_workbook(self.saida).active
        self.assertEqual(sheet.max_row, 2)

    def test_valor_gravado_como_numero_nao_texto(self):
        sheet = self.gerar([("45", Decimal("7.70"))])
        self.assertIsInstance(sheet.cell(row=2, column=5).value, float)

    def test_acha_as_colunas_por_nome_nao_por_posicao(self):
        """O layout é do Superlógica e pode mudar: as colunas são localizadas
        pelo nome normalizado, com acento e caixa ignorados."""
        modelo = montar_modelo(
            os.path.join(self.pasta, "outra_ordem.xlsx"),
            cabecalho=["VALOR", "fornecedor", "  Condomínio  "],
            molde=[None, "DINAMICA SERVICOS POSTAIS", None])
        app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        sheet = load_workbook(self.saida).active
        self.assertEqual(sheet.cell(row=2, column=1).value, 7.70)
        self.assertEqual(sheet.cell(row=2, column=3).value, "45")

    def test_modelo_sem_coluna_condominio_avisa(self):
        modelo = montar_modelo(os.path.join(self.pasta, "sem_cond.xlsx"),
                               cabecalho=["valor", "fornecedor"],
                               molde=[None, "DINAMICA"])
        with self.assertRaises(ValueError) as erro:
            app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        self.assertIn("condominio", str(erro.exception))

    def test_modelo_sem_coluna_valor_avisa(self):
        modelo = montar_modelo(os.path.join(self.pasta, "sem_valor.xlsx"),
                               cabecalho=["condomínio", "fornecedor"],
                               molde=[None, "DINAMICA"])
        with self.assertRaises(ValueError) as erro:
            app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        self.assertIn("valor", str(erro.exception))

    def test_lista_vazia_nao_gera_arquivo_pela_metade(self):
        with self.assertRaises(ValueError):
            app.gerar_planilha_despesas(self.modelo, self.saida, [])
        self.assertFalse(os.path.isfile(self.saida))

    def test_estilo_do_molde_e_preservado_nas_linhas_novas(self):
        """Copiar em vez de reconstruir é o ponto do desenho: o importador do
        Superlógica pode depender de formato de célula."""
        wb = load_workbook(self.modelo)
        wb.active.cell(row=2, column=3).font = Font(bold=True, italic=True)
        wb.save(self.modelo)
        sheet = self.gerar([("45", Decimal("7.70")), ("48", Decimal("57.75"))])
        for linha in (2, 3):
            self.assertTrue(sheet.cell(row=linha, column=3).font.bold)
            self.assertTrue(sheet.cell(row=linha, column=3).font.italic)


class TestNormalizarCabecalho(unittest.TestCase):
    def test_tira_acento_caixa_e_espacos(self):
        self.assertEqual(app._normalizar_cabecalho("  Condomínio "), "condominio")
        self.assertEqual(app._normalizar_cabecalho("VALOR"), "valor")

    def test_celula_vazia_vira_string_vazia(self):
        self.assertEqual(app._normalizar_cabecalho(None), "")


class TestLancamentosDeDespesa(unittest.TestCase):
    def test_um_lancamento_por_protocolo_na_ordem(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, SAN_REMO]), CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75")),
                                       ("42", Decimal("11.55"))])
        self.assertEqual(travas, {"sem_valor": [], "sem_id_sl": []})

    def test_mesmo_condominio_duas_vezes_gera_dois_lancamentos(self):
        """Cada protocolo físico vira um lançamento — somar quebraria a
        correspondência com o papel que originou cada um."""
        outro = dict(KLOSTERS, arquivo="a2.pdf", valor=Decimal("138.60"))
        lancamentos, _ = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, outro]), CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75")),
                                       ("44", Decimal("138.60"))])

    def test_pendente_sem_valor_trava(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS], pendentes=[{"arquivo": "p.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(travas["sem_valor"], ["p.pdf"])
        self.assertEqual(travas["sem_id_sl"], [])

    def test_condominio_sem_id_sl_trava(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, LAGO]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["c.pdf"])
        self.assertEqual(travas["sem_valor"], [])

    def test_protocolo_sem_codigo_trava_por_falta_de_id(self):
        _, travas = app.lancamentos_de_despesa(
            resultado_com([SEM_CODIGO]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["d.pdf"])

    def test_codigo_fora_do_cadastro_trava_por_falta_de_id(self):
        fora = dict(KLOSTERS, arquivo="e.pdf", codigo="99999")
        _, travas = app.lancamentos_de_despesa(
            resultado_com([fora]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["e.pdf"])

    def test_arquivo_que_nao_e_protocolo_nao_trava_nem_vira_lancamento(self):
        """Ignorados nunca deveriam virar despesa — não podem travar o lote."""
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS], ignorados=[{"arquivo": "outro.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75"))])
        self.assertEqual(travas, {"sem_valor": [], "sem_id_sl": []})

    def test_lote_inteiro_travado_devolve_lista_vazia(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([LAGO], pendentes=[{"arquivo": "p.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(lancamentos, [])
        self.assertEqual(travas["sem_id_sl"], ["c.pdf"])
        self.assertEqual(travas["sem_valor"], ["p.pdf"])

    def test_valor_continua_decimal(self):
        """Dinheiro só vira float na escrita da planilha."""
        lancamentos, _ = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS]), CADASTRO_TESTE)
        self.assertIsInstance(lancamentos[0][1], Decimal)


if __name__ == "__main__":
    unittest.main()
