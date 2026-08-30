"""
Qual valor vai no carimbo do protocolo.

Regressão de um erro real: escolher o condomínio de um pendente que ainda não
tinha valor caía em `valor_protocolo(None, tarifa)` e estourava
`int() argument must be ... not 'NoneType'`. O PDF então não era recarimbado,
e o papel ficava sem identificação nenhuma — justamente o oposto do que a
ação pretendia.
"""
import os
import sys
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import logica as app

TARIFA = Decimal("3.85")


class TestValorDoCarimbo(unittest.TestCase):
    def test_valor_informado_a_mao_vence_sempre(self):
        # Foi uma pessoa que decidiu, olhando o papel: prevalece sobre a conta.
        self.assertEqual(
            app.valor_do_carimbo(Decimal("78.00"), 12, TARIFA), Decimal("78.00"))

    def test_contagem_vezes_tarifa(self):
        self.assertEqual(app.valor_do_carimbo(None, 12, TARIFA), Decimal("46.20"))

    def test_sem_valor_e_sem_contagem_devolve_none(self):
        # O caso do bug: anotar o condomínio antes de informar o valor.
        self.assertIsNone(app.valor_do_carimbo(None, None, TARIFA))

    def test_sem_tarifa_no_lote_tambem_devolve_none(self):
        self.assertIsNone(app.valor_do_carimbo(None, 12, None))
        self.assertIsNone(app.valor_do_carimbo(None, None, None))

    def test_valor_zero_informado_a_mao_e_respeitado(self):
        # 0 só chega aqui se alguém digitou; a validação da grade recusa antes.
        # O que não pode é 0 virar "desconhecido" e sumir do carimbo.
        self.assertEqual(app.valor_do_carimbo(Decimal("0"), None, None), Decimal("0"))

    def test_devolve_decimal_nunca_float(self):
        # Dinheiro não passa por float em nenhum ponto do fluxo.
        self.assertIsInstance(app.valor_do_carimbo(None, 3, TARIFA), Decimal)


if __name__ == "__main__":
    unittest.main()
