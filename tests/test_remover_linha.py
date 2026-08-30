"""
Remoção de uma linha da planilha do lote (aba 3).

Existe para o arquivo que não tem como ser resolvido — PDF corrompido ou de
0 byte — que de outro modo ficaria pendente para sempre travando a geração
das despesas.

O ponto delicado é o reindexar: `indice_linha` é POSIÇÃO em `ctx["linhas"]`,
então todo registro depois do removido anda uma casa para trás. Sem isso, o
registro seguinte apontaria para a linha errada e uma edição posterior
escreveria no vizinho.
"""
import os
import sys
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import logica as app


def cenario():
    """Lote de 4 linhas: 2 calculadas, 1 pendente, 1 ignorada."""
    ctx = {"linhas": [
        ["a.pdf", "A", "10001", 3, 3.85, 11.55, ""],
        ["b.pdf", "B", "10002", 5, 3.85, 19.25, ""],
        ["ruim.pdf", "", "", None, None, None, "Erro ao ler"],
        ["c.pdf", "C", "10003", 2, 3.85, 7.70, ""],
    ]}
    resultado = {
        "processados": [
            {"arquivo": "a.pdf", "indice_linha": 0, "valor": Decimal("11.55")},
            {"arquivo": "b.pdf", "indice_linha": 1, "valor": Decimal("19.25")},
            {"arquivo": "c.pdf", "indice_linha": 3, "valor": Decimal("7.70")},
        ],
        "pendentes": [{"arquivo": "ruim.pdf", "indice_linha": 2, "valor": None}],
        "ignorados": [],
        "total": 4,
        "total_valor": Decimal("38.50"),
    }
    return ctx, resultado


class TestRemoverLinha(unittest.TestCase):
    def test_tira_a_linha_da_planilha(self):
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 2)
        self.assertEqual([l[0] for l in ctx["linhas"]], ["a.pdf", "b.pdf", "c.pdf"])

    def test_tira_o_registro_do_painel(self):
        ctx, res = cenario()
        removido = app.remover_linha_do_lote(ctx, res, 2)
        self.assertEqual(removido["arquivo"], "ruim.pdf")
        self.assertEqual(res["pendentes"], [])

    def test_reindexa_o_que_vem_depois(self):
        # c.pdf estava na linha 3; com a 2 removida, passa para a 2.
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 2)
        posicoes = {r["arquivo"]: r["indice_linha"] for r in res["processados"]}
        self.assertEqual(posicoes, {"a.pdf": 0, "b.pdf": 1, "c.pdf": 2})

    def test_indice_continua_apontando_para_a_linha_certa(self):
        # A garantia que importa: depois de remover, o índice de cada
        # registro tem que casar com a linha do mesmo arquivo.
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 0)
        for registro in res["processados"] + res["pendentes"]:
            with self.subTest(arquivo=registro["arquivo"]):
                self.assertEqual(ctx["linhas"][registro["indice_linha"]][0],
                                 registro["arquivo"])

    def test_total_de_arquivos_cai(self):
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 2)
        self.assertEqual(res["total"], 3)

    def test_total_em_reais_perde_o_valor_removido(self):
        # Senão o cartão TOTAL passa a divergir da soma da planilha.
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 1)      # b.pdf, R$ 19,25
        self.assertEqual(res["total_valor"], Decimal("19.25"))

    def test_remover_pendente_sem_valor_nao_mexe_no_total(self):
        ctx, res = cenario()
        app.remover_linha_do_lote(ctx, res, 2)      # sem valor
        self.assertEqual(res["total_valor"], Decimal("38.50"))

    def test_indice_invalido_nao_faz_nada(self):
        ctx, res = cenario()
        for indice in (-1, 4, 99):
            with self.subTest(indice=indice):
                self.assertIsNone(app.remover_linha_do_lote(ctx, res, indice))
        self.assertEqual(len(ctx["linhas"]), 4)

    def test_remover_todas_esvazia_sem_estourar(self):
        ctx, res = cenario()
        while ctx["linhas"]:
            app.remover_linha_do_lote(ctx, res, 0)
        self.assertEqual(ctx["linhas"], [])
        self.assertEqual(res["total"], 0)


class TestCodigoSemIdSlNaEdicao(unittest.TestCase):
    """
    Editar o código na grade tem de exigir o ID SL, como o botão "Escolher
    condomínio" já exigia. Sem essa checagem a edição era aceita, a linha
    continuava pendente e nada na tela explicava por quê.
    """

    CADASTRO = {
        "11111111111111": {"codigo": "10001", "nome": "COM ID", "id_sl": "44"},
        "22222222222222": {"codigo": "10002", "nome": "SEM ID", "id_sl": ""},
    }

    def test_codigo_com_id_sl_e_aceito(self):
        valor, erro = app.validar_edicao_protocolo(
            app.COL_CODIGO, "10001", self.CADASTRO)
        self.assertEqual((valor, erro), ("10001", ""))

    def test_codigo_sem_id_sl_e_recusado_explicando(self):
        valor, erro = app.validar_edicao_protocolo(
            app.COL_CODIGO, "10002", self.CADASTRO)
        self.assertIsNone(valor)
        self.assertIn("ID SL", erro)
        self.assertIn("SEM ID", erro)

    def test_codigo_inexistente_continua_recusado(self):
        _v, erro = app.validar_edicao_protocolo(
            app.COL_CODIGO, "99999", self.CADASTRO)
        self.assertIn("não está no cadastro", erro)


if __name__ == "__main__":
    unittest.main()
