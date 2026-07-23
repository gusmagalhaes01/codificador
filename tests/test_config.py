import json
import os
import shutil
import sys
import tempfile
import unittest

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import identificacao_por_cnpj_6_0 as app


class TestConfig(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self._pasta_base_original = app.pasta_base
        app.pasta_base = lambda: self.tmp

    def tearDown(self):
        app.pasta_base = self._pasta_base_original

    def _caminho_config(self):
        return os.path.join(self.tmp, "config.json")

    def test_sem_config_devolve_3_predefinicoes(self):
        c = app.carregar_config()
        self.assertEqual(set(c["predefinicoes"].keys()),
                         {"fedcorp", "ff", "notas_diversas"})
        self.assertEqual(c["predefinicao_ativa"], "fedcorp")

    def test_notas_diversas_sem_emitente(self):
        c = app.carregar_config()
        self.assertEqual(c["predefinicoes"]["notas_diversas"]["cnpj_emitente"], "")

    def test_config_corrompido_nao_crasha(self):
        with open(self._caminho_config(), "w", encoding="utf-8") as f:
            f.write("{ isso não é json válido")
        c = app.carregar_config()  # não deve lançar
        self.assertEqual(set(c["predefinicoes"].keys()),
                         {"fedcorp", "ff", "notas_diversas"})

    def test_migracao_de_config_antigo_plano(self):
        # formato antigo: campos de lote soltos na raiz (antes das predefinições)
        with open(self._caminho_config(), "w", encoding="utf-8") as f:
            json.dump({"cnpj_emitente": "11.111.111/0001-11",
                       "modo_texto": "rodape", "tema": "escuro"}, f)
        c = app.carregar_config()
        self.assertEqual(c["predefinicoes"]["fedcorp"]["cnpj_emitente"],
                         "11.111.111/0001-11")
        self.assertEqual(c["predefinicoes"]["fedcorp"]["modo_texto"], "rodape")
        self.assertEqual(c["tema"], "escuro")
        # outros perfis intactos
        self.assertEqual(c["predefinicoes"]["ff"]["cnpj_emitente"],
                         "13.736.666/0001-54")

    def test_salvar_e_recarregar_preserva(self):
        c = app.carregar_config()
        c["predefinicao_ativa"] = "ff"
        c["predefinicoes"]["ff"]["tipo_servico"] = "PCMSO"
        app.salvar_config(c)
        c2 = app.carregar_config()
        self.assertEqual(c2["predefinicao_ativa"], "ff")
        self.assertEqual(c2["predefinicoes"]["ff"]["tipo_servico"], "PCMSO")


if __name__ == "__main__":
    unittest.main()
