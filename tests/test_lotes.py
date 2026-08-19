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

SAIDA = os.path.join("C:", os.sep, "saida")


class TestCaminhoDoLote(unittest.TestCase):
    def test_primeiro_lote(self):
        # 1º ao 20º arquivo (índices 0-19) vão para "Lote 01"
        self.assertEqual(app.caminho_do_lote(SAIDA, 0, 20), os.path.join(SAIDA, "Lote 01"))
        self.assertEqual(app.caminho_do_lote(SAIDA, 19, 20), os.path.join(SAIDA, "Lote 01"))

    def test_segundo_lote_comeca_no_21o(self):
        self.assertEqual(app.caminho_do_lote(SAIDA, 20, 20), os.path.join(SAIDA, "Lote 02"))
        self.assertEqual(app.caminho_do_lote(SAIDA, 39, 20), os.path.join(SAIDA, "Lote 02"))

    def test_numeracao_com_dois_digitos(self):
        # "Lote 10", não "Lote 010" — e ordena certo no explorador de arquivos
        self.assertEqual(app.caminho_do_lote(SAIDA, 180, 20), os.path.join(SAIDA, "Lote 10"))

    def test_lote_de_tamanho_diferente(self):
        self.assertEqual(app.caminho_do_lote(SAIDA, 5, 5), os.path.join(SAIDA, "Lote 02"))

    def test_tamanho_invalido_nao_separa(self):
        # 0 ou negativo = sem separação: devolve a pasta de saída original,
        # em vez de estourar divisão por zero no meio de um processamento.
        self.assertEqual(app.caminho_do_lote(SAIDA, 7, 0), SAIDA)
        self.assertEqual(app.caminho_do_lote(SAIDA, 7, -3), SAIDA)

    def test_continuidade_do_lote_na_resolucao_manual(self):
        """A resolução manual de um pendente (aba 3, `_acao_informar_valor`)
        delega pra `resolver_protocolo_manual` (`logica.py`), que usa
        `ctx["carimbados"]` — o contador de arquivos efetivamente carimbados
        no laço automático — como índice do próximo arquivo, e só incrementa
        depois do carimbo dar certo (mesma convenção do laço automático:
        `carimbados` é sempre o índice 0-based do PRÓXIMO arquivo, não do
        último já gravado).

        Chama a função de produção de verdade (não só `caminho_do_lote`
        isolado) pra também cobrir o incremento de `ctx["carimbados"]` — sem
        ele, duas resoluções manuais seguidas carimbariam por cima do mesmo
        "Lote NN", sobrescrevendo o primeiro arquivo resolvido.

        Simula um lote de tamanho 5 em que o processamento automático já
        carimbou 7 arquivos (índices 0-6, preenchendo "Lote 01" inteiro e
        mais 2 do "Lote 02") e depois duas resoluções manuais seguidas — a
        numeração tem que continuar exatamente do 7º e 8º arquivo, sem pular
        (ex.: indo direto pro "Lote 03") nem repetir (ex.: caindo nos mesmos
        dois primeiros arquivos do "Lote 02").
        """
        tamanho_lote = 5
        ctx = {
            "pasta_saida": SAIDA,
            "separar_em_lotes": True,
            "tamanho_lote": tamanho_lote,
            "carimbados": 7,  # estado ao abrir o painel de resultado
            "linhas": [None],
        }
        cadastro = {}
        pastas_carimbadas = []

        def carimbar_fake(pasta_destino, registro_dados):
            # Só registra onde "gravaria" o PDF — sem tocar disco nem PDF de
            # verdade, é o próprio contrato de `carimbar` em
            # `resolver_protocolo_manual`.
            pastas_carimbadas.append(pasta_destino)

        dados = {"arquivo": "protocolo.pdf", "codigo": None,
                 "condominio": "", "indice_linha": 0,
                 "motivo_original": ""}

        # 1ª resolução manual: usa o índice 7 (valor atual de "carimbados"),
        # ainda no "Lote 02" (que vai até o índice 9).
        app.resolver_protocolo_manual(ctx, dados, app.Decimal("7.70"),
                                       cadastro, carimbar_fake)
        self.assertEqual(pastas_carimbadas[-1], os.path.join(SAIDA, "Lote 02"))
        self.assertEqual(ctx["carimbados"], 8)  # só avança após o carimbo

        # 2ª resolução manual: continua do 8, não repete o 7 nem pula pro 9.
        app.resolver_protocolo_manual(ctx, dados, app.Decimal("7.70"),
                                       cadastro, carimbar_fake)
        self.assertEqual(pastas_carimbadas[-1], os.path.join(SAIDA, "Lote 02"))
        self.assertEqual(ctx["carimbados"], 9)

        # Se a 3ª resolução manual viesse a seguir, já estouraria pro
        # "Lote 03" (índice 10) — confirma que a fronteira entre lotes
        # continua sendo respeitada pela numeração herdada do automático.
        pasta_3 = app.caminho_do_lote(SAIDA, ctx["carimbados"] + 1, tamanho_lote)
        self.assertEqual(pasta_3, os.path.join(SAIDA, "Lote 03"))


if __name__ == "__main__":
    unittest.main()
