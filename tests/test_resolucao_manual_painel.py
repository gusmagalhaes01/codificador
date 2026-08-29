"""
Regressão: o registro criado ao resolver um pendente à mão (aba 3) precisa
carregar os campos que as ações do painel usam DEPOIS.

O bug real: `_listas_do_painel_protocolos` mostra junto dos pendentes todo
processado que ainda esteja sem ID SL. Um protocolo resolvido por "Informar
valor" mas cujo condomínio continua desconhecido volta, portanto, para a
tabela de pendentes — e "Escolher condomínio" estourava
`KeyError('caminho')` no recarimbo, seguido de `KeyError('indice_linha')`
fora do try/except (que caía no handler global).
"""
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

#  Campos que `_acao_escolher_condominio` e `_acao_abrir_pdf_protocolo` leem
#  do dict do pendente. Se o registro resolvido não os tiver, aquelas ações
#  quebram quando o item reaparece entre os pendentes.
CAMPOS_EXIGIDOS_PELAS_ACOES = ("arquivo", "caminho", "indice_linha")


def pendente(**extra):
    base = {
        "arquivo": "doc16004820260821155206.pdf",
        "caminho": os.path.join("C:", os.sep, "Correios II",
                                "doc16004820260821155206.pdf"),
        "indice_linha": 4,
        "condominio": "",
        "codigo": None,
        "motivo": "Não foi possível ler o documento",
        "motivo_original": "Não foi possível ler o documento",
    }
    base.update(extra)
    return base


class TestRegistroResolvidoManualmente(unittest.TestCase):
    def setUp(self):
        self.dados = pendente()
        self.registro = app.registro_painel_resolvido(
            self.dados, app.Decimal("46.20"), "Valor informado manualmente",
            os.path.join("C:", os.sep, "saida", "Lote 02"))

    def test_carrega_os_campos_que_as_acoes_do_painel_usam(self):
        for campo in CAMPOS_EXIGIDOS_PELAS_ACOES:
            with self.subTest(campo=campo):
                self.assertIn(campo, self.registro)
                self.assertIsNotNone(self.registro[campo])

    def test_caminho_e_indice_vem_do_pendente_original(self):
        self.assertEqual(self.registro["caminho"], self.dados["caminho"])
        self.assertEqual(self.registro["indice_linha"], self.dados["indice_linha"])

    def test_guarda_a_pasta_do_lote_para_o_recarimbo_cair_no_mesmo(self):
        self.assertEqual(self.registro["pasta_destino"],
                         os.path.join("C:", os.sep, "saida", "Lote 02"))

    def test_valor_fica_em_decimal_e_unidades_vazias(self):
        # Resolvido à mão não tem contagem de unidades: a planilha precisa
        # sair com Unidades e Tarifa em branco (ver linha_planilha_protocolo).
        self.assertEqual(self.registro["valor"], app.Decimal("46.20"))
        self.assertIsNone(self.registro["unidades"])

    def test_motivo_original_nao_vem_ja_mesclado(self):
        # Semente para uma 2ª resolução: se viesse a mesclada, um eventual
        # "Código não cadastrado" apareceria em dobro.
        self.assertEqual(self.registro["motivo_original"],
                         self.dados["motivo_original"])

    def test_sem_motivo_nao_inventa_a_chave(self):
        registro = app.registro_painel_resolvido(
            pendente(), app.Decimal("7.70"), "", None)
        self.assertNotIn("motivo", registro)


class TestFluxoInformarValorDepoisEscolherCondominio(unittest.TestCase):
    """
    Percorre a sequência que quebrava: resolve o valor pela função de
    produção e confere que o registro resultante ainda serve de entrada para
    o recarimbo — que é o que "Escolher condomínio" faz em seguida.
    """

    def test_registro_resolvido_serve_de_entrada_para_o_recarimbo(self):
        ctx = {
            "pasta_saida": os.path.join("C:", os.sep, "saida"),
            "separar_em_lotes": False,
            "carimbados": 3,
            "linhas": [None] * 6,
        }
        dados = pendente()
        _linha, motivo, pasta_destino = app.resolver_protocolo_manual(
            ctx, dados, app.Decimal("46.20"), {}, lambda _p, _r: None)

        registro = app.registro_painel_resolvido(
            dados, app.Decimal("46.20"), motivo, pasta_destino)

        # O recarimbo lê exatamente estes campos do dict.
        self.assertEqual(registro["caminho"], dados["caminho"])
        self.assertEqual(registro["arquivo"], dados["arquivo"])
        self.assertIsInstance(registro["indice_linha"], int)
        # E o índice tem que continuar apontando para a linha certa.
        self.assertLess(registro["indice_linha"], len(ctx["linhas"]))


if __name__ == "__main__":
    unittest.main()
