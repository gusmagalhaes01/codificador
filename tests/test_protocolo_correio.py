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


if __name__ == "__main__":
    unittest.main()
