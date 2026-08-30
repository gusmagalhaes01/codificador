"""
Um registro pertence a `processados` quando tem valor e a `pendentes` quando
não tem — independentemente de POR ONDE o valor foi informado.

Regressão de um bug real: havia dois caminhos para dar valor a um protocolo,
o botão "Informar valor" (que movia a linha) e a edição da célula na grade
(que gravava o valor e deixava o registro em `pendentes`). No segundo caso a
linha continuava marcada como pendente e travava a geração da planilha de
despesas, sem nada na tela explicando o motivo.
"""
import os
import sys
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import logica as app


def resultado_com(pendentes=(), processados=(), ignorados=()):
    return {"pendentes": list(pendentes), "processados": list(processados),
            "ignorados": list(ignorados)}


class TestReclassificar(unittest.TestCase):
    def test_pendente_que_ganhou_valor_vai_para_processados(self):
        reg = {"arquivo": "a.pdf", "valor": None}
        res = resultado_com(pendentes=[reg])
        reg["valor"] = Decimal("50.00")          # como a edição da célula faz
        self.assertTrue(app.reclassificar_registro(res, reg))
        self.assertEqual(res["pendentes"], [])
        self.assertEqual(res["processados"], [reg])

    def test_ignorado_completado_a_mao_vira_cobravel(self):
        reg = {"arquivo": "x.pdf", "valor": Decimal("10.00")}
        res = resultado_com(ignorados=[reg])
        app.reclassificar_registro(res, reg)
        self.assertEqual(res["ignorados"], [])
        self.assertEqual(res["processados"], [reg])

    def test_processado_que_perdeu_o_valor_volta_a_pendente(self):
        reg = {"arquivo": "b.pdf", "valor": None}
        res = resultado_com(processados=[reg])
        app.reclassificar_registro(res, reg)
        self.assertEqual(res["processados"], [])
        self.assertEqual(res["pendentes"], [reg])

    def test_quem_ja_esta_no_lugar_certo_nao_muda(self):
        reg = {"arquivo": "c.pdf", "valor": Decimal("1.00")}
        res = resultado_com(processados=[reg])
        self.assertFalse(app.reclassificar_registro(res, reg))
        self.assertEqual(res["processados"], [reg])

    def test_nao_duplica_o_registro(self):
        reg = {"arquivo": "d.pdf", "valor": Decimal("1.00")}
        res = resultado_com(pendentes=[reg])
        app.reclassificar_registro(res, reg)
        app.reclassificar_registro(res, reg)
        self.assertEqual(len(res["processados"]), 1)

    def test_none_nao_estoura(self):
        self.assertFalse(app.reclassificar_registro(resultado_com(), None))

    def test_valor_zero_conta_como_valor(self):
        # 0 só chega aqui se alguém digitou; "sem valor" é None, não 0.
        reg = {"arquivo": "e.pdf", "valor": Decimal("0")}
        res = resultado_com(pendentes=[reg])
        app.reclassificar_registro(res, reg)
        self.assertEqual(res["processados"], [reg])


class TestDestravaAGeracao(unittest.TestCase):
    """O efeito que importa: a linha para de travar a planilha de despesas."""

    CADASTRO = {"11111111111111": {"codigo": "11225", "nome": "PRO RIOMAR",
                                    "id_sl": "213"}}

    def test_antes_trava_depois_nao(self):
        reg = {"arquivo": "doc.pdf", "codigo": "11225", "valor": None}
        res = {"pendentes": [reg], "processados": [], "ignorados": []}

        _l, travas = app.lancamentos_de_despesa(res, self.CADASTRO)
        self.assertIn("doc.pdf", travas["sem_valor"])

        reg["valor"] = Decimal("50.00")
        app.reclassificar_registro(res, reg)

        lancamentos, travas = app.lancamentos_de_despesa(res, self.CADASTRO)
        self.assertEqual(travas["sem_valor"], [])
        self.assertEqual(travas["sem_id_sl"], [])
        self.assertEqual(lancamentos, [("213", Decimal("50.00"))])


if __name__ == "__main__":
    unittest.main()
