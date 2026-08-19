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
        """Entradas gigantes não podem escapar como número gigante sem
        aviso — quem barra "1"*40 e "1e30" é o teto LIMITE_VALOR_DIGITADO
        (nenhuma das duas levanta InvalidOperation na construção do
        Decimal), então o usuário vê o aviso normal, não um popup de erro
        técnico."""
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

    def test_recusa_espaco_no_meio_dos_digitos(self):
        """"3 85" e "57 75" não podem virar 385/5775 por causa do
        `.replace(" ", "")` cru — isso multiplicaria o valor por cem sem
        nenhum aviso, e o resultado vira dinheiro cobrado de condomínio."""
        for entrada in ("3 85", "57 75"):
            valor, mensagem = app.converter_valor_digitado(entrada)
            self.assertIsNone(valor, entrada)
            self.assertTrue(mensagem, entrada)

    def test_tolera_espaco_nas_pontas_e_apos_cifrao(self):
        """O espaço existe pra tolerar o prefixo ("R$ 57,75") e sobras nas
        pontas, não pra juntar dígitos — esses dois formatos continuam
        aceitos depois da correção do espaço no meio."""
        self.assertEqual(app.converter_valor_digitado(" R$ 57,75 ")[0], Decimal("57.75"))
        self.assertEqual(app.converter_valor_digitado("R$57,75")[0], Decimal("57.75"))

    def test_aceita_zero_a_direita_sem_casa_fracionada(self):
        """"3,850" e "3,8500" são exatamente 3,85 — o zero à direita não
        acrescenta casa decimal nenhuma, então não podem ser recusados pela
        regra de duas casas (que é sobre o valor, não sobre a forma como foi
        digitado). A validação antiga da tarifa aceitava esse caso."""
        self.assertEqual(app.converter_valor_digitado("3,850")[0], Decimal("3.85"))
        self.assertEqual(app.converter_valor_digitado("3,8500")[0], Decimal("3.85"))

    def test_recusa_centavo_fracionado_de_verdade(self):
        """"3,855" e "300,305" têm centavo fracionado de verdade (não dá pra
        representar em duas casas sem perder valor) — continuam recusados
        mesmo com a regra reescrita em cima do valor."""
        for entrada in ("3,855", "300,305"):
            valor, mensagem = app.converter_valor_digitado(entrada)
            self.assertIsNone(valor, entrada)
            self.assertIn("duas casas", mensagem)

    def test_recusa_ponto_ambiguo_de_milhar(self):
        """Ponto sem vírgula seguido de três dígitos é separador de milhar
        mal digitado, não decimal — mesmo quando o valor "arredondaria
        certo". "1.200" é o caso mais perigoso: interpretado como decimal
        vira Decimal("1.200"), que quantiza pra 1,20 sem sobrar casa
        nenhuma (o zero à direita some), então passaria calado se a
        checagem dependesse só do quantize — quem digitou mil e duzentos
        reais seria cobrado um real e vinte, sem aviso nenhum."""
        for entrada in ("1.200", "1.500", "1.234"):
            valor, mensagem = app.converter_valor_digitado(entrada)
            self.assertIsNone(valor, entrada)
            self.assertIn("milhar", mensagem, entrada)

    def test_aceita_ponto_nao_ambiguo(self):
        """"300.30" (dois dígitos após o ponto) continua valendo como
        decimal, e "1.234,56" continua valendo porque a vírgula já deixa
        claro qual é o separador decimal — só o ponto sozinho seguido de
        três dígitos é recusado."""
        self.assertEqual(app.converter_valor_digitado("300.30")[0], Decimal("300.30"))
        self.assertEqual(app.converter_valor_digitado("1.234,56")[0], Decimal("1234.56"))
        self.assertEqual(app.converter_valor_digitado("3,85")[0], Decimal("3.85"))

    def test_mensagem_de_ambiguidade_usa_o_exemplo_recebido(self):
        """A função recebe um `exemplo` diferente por tela (tarifa: "3,85";
        valor manual do protocolo: "300,30" por padrão) justamente para não
        sugerir uma ordem de grandeza errada. A mensagem de ambiguidade
        precisa usar esse mesmo parâmetro, não um valor fixo na casa do
        milhar hardcoded no meio da função."""
        _, mensagem_tarifa = app.converter_valor_digitado("1.200", exemplo="3,85")
        self.assertIn("3,85", mensagem_tarifa)

        _, mensagem_valor_manual = app.converter_valor_digitado("1.200", exemplo="300,30")
        self.assertIn("300,30", mensagem_valor_manual)


if __name__ == "__main__":
    unittest.main()
