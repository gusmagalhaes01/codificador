import os
import sys
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import logica as app


class TestConverterValorDigitado(unittest.TestCase):
    def test_virgula_e_ponto_dao_o_mesmo_valor(self):
        self.assertEqual(app.converter_valor_digitado("300,30")[0], Decimal("300.30"))
        self.assertEqual(app.converter_valor_digitado("300.30")[0], Decimal("300.30"))

    def test_aceita_separador_de_milhar(self):
        self.assertEqual(app.converter_valor_digitado("1.234,56")[0], Decimal("1234.56"))

    def test_aceita_cifrao_e_espacos(self):
        self.assertEqual(app.converter_valor_digitado(" R$ 57,75 ")[0], Decimal("57.75"))

    def test_valido_nao_traz_mensagem(self):
        self.assertEqual(app.converter_valor_digitado("3,85")[1], "")

    def test_recusa_tres_casas_decimais(self):
        valor, mensagem = app.converter_valor_digitado("300,305")
        self.assertIsNone(valor)
        self.assertIn("duas casas", mensagem)

    def test_recusa_zero(self):
        valor, mensagem = app.converter_valor_digitado("0")
        self.assertIsNone(valor)
        self.assertIn("maior que zero", mensagem)

    def test_recusa_negativo(self):
        valor, mensagem = app.converter_valor_digitado("-5")
        self.assertIsNone(valor)
        self.assertIn("maior que zero", mensagem)

    def test_recusa_texto(self):
        valor, mensagem = app.converter_valor_digitado("abc")
        self.assertIsNone(valor)
        self.assertIn("número", mensagem)

    def test_recusa_vazio(self):
        valor, mensagem = app.converter_valor_digitado("   ")
        self.assertIsNone(valor)
        self.assertIn("número", mensagem)

    def test_entrada_absurda_devolve_aviso_em_vez_de_estourar(self):
        """Entradas gigantes não podem levantar InvalidOperation crua — o
        usuário precisa ver o aviso normal, não um popup de erro técnico."""
        for entrada in ("1" * 40, "1e30"):
            valor, mensagem = app.converter_valor_digitado(entrada)
            self.assertIsNone(valor, entrada)
            self.assertTrue(mensagem, entrada)

    def test_limite_de_sanidade_tem_fronteira_no_lugar_certo(self):
        """Reforço além do brief: o teto (LIMITE_VALOR_DIGITADO) precisa
        aceitar exatamente o limite e recusar um centavo a mais — sem essa
        checagem de fronteira, um `>=` trocado por `>` (ou vice-versa) passa
        despercebido pelos outros testes, que só usam valores bem abaixo ou
        bem acima do teto."""
        limite = app.LIMITE_VALOR_DIGITADO
        no_limite, mensagem_no_limite = app.converter_valor_digitado(str(limite))
        self.assertEqual(no_limite, limite)
        self.assertEqual(mensagem_no_limite, "")

        acima, mensagem_acima = app.converter_valor_digitado(str(limite + Decimal("0.01")))
        self.assertIsNone(acima)
        self.assertTrue(mensagem_acima)

    def test_recusa_apenas_cifrao(self):
        """"R$" sozinho (sem dígito nenhum) precisa cair no mesmo aviso de
        "não é um número", não em outro erro — o texto sobra vazio depois de
        remover o prefixo, então tem que passar pelo mesmo caminho de
        `bruto` vazio que o teste de espaços em branco já cobre."""
        valor, mensagem = app.converter_valor_digitado("R$")
        self.assertIsNone(valor)
        self.assertIn("número", mensagem)


if __name__ == "__main__":
    unittest.main()
