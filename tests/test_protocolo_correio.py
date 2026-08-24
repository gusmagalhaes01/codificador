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
from cadastro_teste import CADASTRO_TESTE

# Trechos representativos do formato real ("Protocolo de Recebimento de
# Documento" da Imodata/Correios) — só o cabeçalho, sem nomes de moradores
# reais (dado sensível que não precisa estar no fixture pra testar a regex).
TEXTO_COM_MARCADOR = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento "
    "Unidade 401 - Fulano de Tal Listando 1 unidades Imodata - "
    "Condominios e Imoveis Rua Barata Ribeiro, 774"
)
TEXTO_SEM_MARCADOR = "CO-ESTIPULANTE: KLOSTERS CNPJ: 01.195.716/0001-54"


class TestProtocoloCorreio(unittest.TestCase):
    def test_extrai_codigo_quando_marcador_presente(self):
        self.assertEqual(app.extrair_codigo_protocolo_correio(TEXTO_COM_MARCADOR), "10004")

    def test_nenhum_codigo_quando_marcador_ausente(self):
        self.assertIsNone(app.extrair_codigo_protocolo_correio(TEXTO_SEM_MARCADOR))

    def test_texto_vazio_nao_quebra(self):
        self.assertIsNone(app.extrair_codigo_protocolo_correio(""))

    def test_marcador_sem_parenteses_nao_confunde(self):
        texto = "Protocolo de Recebimento de Documento sem codigo nenhum aqui"
        self.assertIsNone(app.extrair_codigo_protocolo_correio(texto))

    def test_codigo_extraido_resolve_no_cadastro(self):
        codigo = app.extrair_codigo_protocolo_correio(TEXTO_COM_MARCADOR)
        codigos_cadastro = app._codigos_do_cadastro(CADASTRO_TESTE)
        self.assertEqual(codigos_cadastro.get(codigo), "01195716000154")

    def test_codigo_nao_cadastrado_nao_resolve(self):
        texto = "W700A DESCONHECIDO (99999) Protocolo de Recebimento de Documento"
        codigo = app.extrair_codigo_protocolo_correio(texto)
        codigos_cadastro = app._codigos_do_cadastro(CADASTRO_TESTE)
        self.assertIsNone(codigos_cadastro.get(codigo))


class TestCodigoComRuidoAntesDoMarcador(unittest.TestCase):
    """O código vem entre parênteses logo antes do título do documento, mas
    nem sempre colado nele: gente escreve o valor à caneta bem ali, e o OCR
    lê esse manuscrito no meio. Caso real do 10536 DONATELLO, em que o
    "7,70" escrito à mão fez o programa não identificar o condomínio — o
    protocolo entrou na planilha com valor e sem código, e travou a geração
    da planilha de despesas sem nenhuma saída pela tela."""

    def test_manuscrito_entre_o_codigo_e_o_titulo(self):
        texto = ("W700A DONATELLO (10536) 7.70 Protocolo de Recebimento de "
                 "Documento Unidade 401 - Fulano de Tal Listando 1 unidade")
        self.assertEqual(app.extrair_codigo_protocolo_correio(texto), "10536")

    def test_codigo_colado_no_titulo_continua_funcionando(self):
        texto = ("W700A KLOSTERS (10004) Protocolo de Recebimento de Documento "
                 "Unidade 401 - Fulano de Tal")
        self.assertEqual(app.extrair_codigo_protocolo_correio(texto), "10004")

    def test_ruido_longo_demais_nao_vale(self):
        """A janela é curta de propósito: um número entre parênteses muito
        antes do título não é o código deste documento."""
        texto = ("(99999) " + "x" * 80 +
                 " Protocolo de Recebimento de Documento Unidade")
        self.assertIsNone(app.extrair_codigo_protocolo_correio(texto))

    def test_pega_o_codigo_mais_proximo_do_titulo(self):
        texto = ("(11111) documento anterior (10536) 7.70 "
                 "Protocolo de Recebimento de Documento")
        self.assertEqual(app.extrair_codigo_protocolo_correio(texto), "10536")


if __name__ == "__main__":
    unittest.main()
