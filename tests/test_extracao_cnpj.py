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

FED = "35315360000167"   # FedCorp (emitente/intermediário)
FF = "13736666000154"    # F&F (emitente)
SAN_REMO = "08578541000103"
KLOSTERS = "01195716000154"
MANHATTAN = "07448975000126"
LAGO_MAGGIORE = "07945453000130"
VILLE = "29273778000156"
IMODATA_INTERM = "12184361000114"
CPF_VILA_MARINA = "52998224725"


def ler(nome):
    caminho = os.path.join(_AQUI, "dados", nome)
    with open(caminho, encoding="utf-8") as f:
        return f.read()


class TestExtracaoCnpj(unittest.TestCase):
    def test_co_estipulante_com_hifen(self):
        # recibo FedCorp com "08.578.541-0001-03" (hífen no lugar da barra)
        texto = ler("fedcorp_recibo_hifen.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertEqual(cands, [SAN_REMO])

    def test_intermediarios_nunca_aparecem(self):
        texto = ler("fedcorp_recibo_hifen.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertNotIn(FED, cands)
        self.assertNotIn(IMODATA_INTERM, cands)

    def test_nfse_ff_isola_klosters(self):
        # com o emitente F&F excluído, sobra só o tomador Klosters
        texto = ler("nfse_ff.txt")
        cands = app.extrair_cnpj_tomador(texto, FF)
        self.assertEqual(cands, [KLOSTERS])

    def test_nfse_imodata_devolve_os_dois(self):
        texto = ler("nfse_imodata.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertEqual(len(cands), 2)
        self.assertIn(MANHATTAN, cands)

    def test_nfse_avulsa_inclui_tomador(self):
        texto = ler("nfse_avulsa.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertIn(LAGO_MAGGIORE, cands)

    def test_boleto_inclui_pagador(self):
        texto = ler("boleto_avulso.txt")
        cands = app.extrair_cnpj_tomador(texto, "")
        self.assertIn(VILLE, cands)

    def test_checksum_invalido_rejeitado(self):
        # documento cujo CNPJ do tomador tem dígito verificador errado (NF-927)
        texto = ler("cnpj_checksum_invalido.txt")
        cands = app.extrair_cnpj_tomador(texto, FED)
        self.assertEqual(cands, [])


class TestExtracaoCpfNoCampoRotulado(unittest.TestCase):
    """CPF depois de um rótulo é sempre aceito: o rótulo é a garantia de que
    aquele documento é o do pagador, e não de uma pessoa qualquer da página."""

    def test_pagador_com_cpf(self):
        texto = "PAGADOR: VILA MARINA CNPJ/CPF: 529.982.247-25"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED), [CPF_VILA_MARINA])

    def test_tomador_com_cpf(self):
        texto = "TOMADOR: VILA MARINA CNPJ: 529.982.247-25"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED), [CPF_VILA_MARINA])

    def test_cpf_com_digito_verificador_errado_e_descartado(self):
        texto = "PAGADOR: VILA MARINA CNPJ/CPF: 529.982.247-24"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED), [])

    def test_cnpj_de_14_digitos_nao_vira_cpf_de_11(self):
        """Os 11 primeiros dígitos de um CNPJ não podem ser lidos como CPF —
        é o erro que os lookarounds de CPF_FLEX existem para impedir."""
        texto = "PAGADOR: SAN REMO CNPJ/CPF: 08.578.541/0001-03"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED), [SAN_REMO])


class TestCpfNoFallbackGenerico(unittest.TestCase):
    """Sem rótulo legível, um CPF solto só vira candidato se já for um
    condomínio conhecido. Num boleto, um CPF solto costuma ser de uma pessoa
    qualquer — síndico, avalista, quem assinou — e aceitá-lo às cegas
    carimbaria o código do condomínio errado."""

    def test_cpf_cadastrado_e_aceito(self):
        texto = "VILA MARINA 529.982.247-25 valor 300,00"
        cands = app.extrair_cnpj_tomador(texto, FED, CADASTRO_TESTE)
        self.assertEqual(cands, [CPF_VILA_MARINA])

    def test_cpf_nao_cadastrado_e_ignorado(self):
        texto = "Sacador avalista JOAO 111.444.777-35 valor 300,00"
        cands = app.extrair_cnpj_tomador(texto, FED, CADASTRO_TESTE)
        self.assertEqual(cands, [])

    def test_sem_cadastro_nenhum_cpf_entra(self):
        """Chamada antiga, de dois argumentos, não muda de comportamento."""
        texto = "VILA MARINA 529.982.247-25 valor 300,00"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED), [])

    def test_cnpj_cadastrado_no_fallback_nao_muda(self):
        texto = "SAN REMO 08.578.541/0001-03 valor 300,00"
        self.assertEqual(app.extrair_cnpj_tomador(texto, FED, CADASTRO_TESTE),
                         [SAN_REMO])

    def test_cnpj_nao_cadastrado_continua_entrando(self):
        """Só o CPF é restrito ao cadastro. O CNPJ desconhecido continua
        virando candidato, e é assim que ele vira o pendente 'não cadastrado'
        pelo qual um condomínio novo é descoberto."""
        texto = "ALGUM CONDOMINIO 04.252.011/0001-10 valor 300,00"
        cands = app.extrair_cnpj_tomador(texto, FED, CADASTRO_TESTE)
        self.assertEqual(cands, ["04252011000110"])


if __name__ == "__main__":
    unittest.main()
