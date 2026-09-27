# Lançamento no Superlógica direto da extração (aba 2) — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fazer a aba 2 gerar, junto da planilha de extração, o arquivo de importação de despesas do Superlógica para boletos do sindicato, DARF do DCTFWeb e notas da F&F, com a opção de lançamento do lote escolhida num seletor.

**Architecture:** Cada opção de lançamento é um `.xlsx` em `modelos_superlogica/` (nome do arquivo = texto do seletor), cuja linha 2 é o molde — o mesmo mecanismo de `gerar_planilha_despesas` que a aba 3 já usa. Na thread de extração, cada documento reconhecido vira um "documento de lançamento" normalizado; os que não podem ser lançados ganham o motivo na coluna Observação da planilha de extração. Ao fim, na thread principal, o programa pergunta chave (e vencimento, se houver nota fiscal) e grava o arquivo do Superlógica com seis colunas por documento sobre o molde.

**Tech Stack:** Python 3.14, `openpyxl`, `re`, CustomTkinter (aba 2), `unittest`.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-09-27-lancamento-superlogica-extracao-design.md`.
- Comentários e nomes de variáveis em **português**.
- `logica.py` **não importa nada de interface**.
- Fixtures de teste são **strings inline** ou arquivos em `tests/dados/` — nunca PDF no repositório.
- As seis colunas por documento são **sempre escritas pelo programa, nunca herdadas do molde**: `condominio`, `valor`, `vencimento`, `linha_digitavel`, `numero_documento`, `competencia` (nomes já normalizados por `_normalizar_cabecalho`). Valor ausente = célula vazia, mesmo que o molde tenha algo ali.
- O programa **não escreve o `complemento`** — ele vem do molde como qualquer outra coluna.
- **`gerar_planilha_despesas` continua aceitando `(id_sl, valor)`**: os testes existentes da aba 3 têm de passar sem alteração.
- Cores só via `self.tema_atual[chave]`; todo widget CustomTkinter com `corner_radius=0`; cobalto (`acento`) só no botão primário e no que resolve pendência.
- Vocabulário da interface para leigos: nunca "OCR" nem "DPI" em texto visível.
- Suíte hoje: **417 testes**. Rodar com `python -m unittest discover -s tests -p "test_*.py"`.
- Validadores manuais levam prefixo `_` para o `unittest discover` não os coletar.

---

### Task 1: Vencimento do DARF lido do texto, com conferência

**Files:**
- Modify: `logica.py` (constante e função logo depois de `extrair_dados_boleto`, que começa na linha 2180)
- Test: `tests/test_lancamento_superlogica.py` (criar)

**Interfaces:**
- Consumes: nada novo.
- Produces: `vencimento_do_darf(texto) -> datetime.date | None`; constantes `RE_VENCIMENTO_DARF` e `MINIMO_OCORRENCIAS_VENCIMENTO_DARF = 2`.

- [ ] **Step 1: Write the failing tests**

Criar `tests/test_lancamento_superlogica.py`:

```python
# -*- coding: utf-8 -*-
"""Lançamento no Superlógica direto da extração (aba 2)."""
import datetime
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

#  Texto nativo como o pypdf devolve um DARF do DCTFWeb (sintético, mesmo
#  layout das guias reais). O vencimento aparece três vezes: no cabeçalho,
#  na composição e no recibo do rodapé.
DARF_OK = (
    "Documento de Arrecadação\nde Receitas Federais\n"
    "Período de Apuração Data de Vencimento Número do Documento\n"
    "07.16.26261.0000000-0 Pagar este documento até\n"
    "18/09/2026Observações\n"
    "Valor Total do Documento\n58,51\n"
    "5952 RET DE CONTRIBUICOES PAGT PJ A PJ DE DIR PRIV 58,51 58,51\n"
    "PA:08/2026 Vencimento:18/09/2026\n"
    "SENDA (Versão:1.5.10) 18/09/2026 09:46:491 1Página: /\n"
    "Número: 07.16.26261.0000000-0\n"
    "Pagar até: 18/09/2026\n"
    "Valor: 58,51\n"
)


class TestVencimentoDoDarf(unittest.TestCase):

    def test_tres_ocorrencias_iguais_dao_a_data(self):
        self.assertEqual(app.vencimento_do_darf(DARF_OK), datetime.date(2026, 9, 18))

    def test_ocorrencias_divergentes_nao_escolhem_nenhuma(self):
        texto = DARF_OK.replace("Pagar até: 18/09/2026", "Pagar até: 19/09/2026")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_uma_ocorrencia_so_nao_basta(self):
        #  A conferência existe porque a data se repete; com uma só, não há
        #  com o que conferir.
        texto = ("Pagar até: 18/09/2026\n"
                 "5952 RET DE CONTRIBUICOES 58,51\n")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_nenhuma_ocorrencia(self):
        self.assertIsNone(app.vencimento_do_darf("Documento de Arrecadação sem data"))

    def test_data_que_nao_existe_e_recusada(self):
        texto = DARF_OK.replace("18/09/2026", "31/02/2026")
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_carimbo_de_emissao_nao_conta(self):
        #  "SENDA ... 18/09/2026 09:46" é quando a guia foi emitida, sem
        #  rótulo de vencimento. Se contasse, trocar a data de emissão
        #  mudaria o resultado.
        texto = DARF_OK.replace("SENDA (Versão:1.5.10) 18/09/2026",
                                "SENDA (Versão:1.5.10) 01/01/2026")
        self.assertEqual(app.vencimento_do_darf(texto), datetime.date(2026, 9, 18))

    def test_multa_apos_vencimento_do_boleto_nao_casa(self):
        texto = "APOS VENCIMENTO MULTA DE 2% e JUROS DE 1%\n10/09/2026\n"
        self.assertIsNone(app.vencimento_do_darf(texto))

    def test_texto_vazio_ou_none(self):
        self.assertIsNone(app.vencimento_do_darf(""))
        self.assertIsNone(app.vencimento_do_darf(None))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: FAIL — `module 'logica' has no attribute 'vencimento_do_darf'`.

- [ ] **Step 3: Write the implementation**

Em `logica.py`, logo depois de `extrair_dados_boleto`:

```python
#  Vencimento de um DARF (guia do DCTFWeb). O código de arrecadação traz valor
#  e dígitos verificadores, mas NÃO traz data — ela só existe no texto da
#  guia. Três rótulos a carregam, em lugares diferentes da folha:
#
#      Pagar este documento até\n18/09/2026   (cabeçalho)
#      PA:08/2026 Vencimento:18/09/2026        (composição)
#      Pagar até: 18/09/2026                   (recibo do rodapé)
#
#  Só casa data COM rótulo: a guia também traz a data de emissão
#  ("SENDA ... 18/09/2026 09:46") sem rótulo de vencimento, e ela não pode
#  entrar na conta. "Vencimento" só conta com dois-pontos logo depois —
#  "Data de Vencimento Número do Documento" (cabeçalho da tabela) e "APOS
#  VENCIMENTO MULTA" (boleto) não têm.
RE_VENCIMENTO_DARF = re.compile(
    r"(?:Pagar\s+(?:este\s+documento\s+)?at[ée]:?|Vencimento:)\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE)

#  A conferência só existe com pelo menos duas ocorrências: com uma, não há
#  com o que comparar, e uma data solta de texto é exatamente o que a aba 2
#  evita. As guias reais trazem três.
MINIMO_OCORRENCIAS_VENCIMENTO_DARF = 2


def vencimento_do_darf(texto):
    """
    Data de vencimento de um DARF, lida do texto nativo.

    Devolve a data só quando ela aparece pelo menos
    MINIMO_OCORRENCIAS_VENCIMENTO_DARF vezes com rótulo e TODAS as
    ocorrências são a mesma data. Divergindo, faltando ou sendo uma data que
    não existe, devolve None — e o DARF fica fora do lançamento, para ser
    lançado à mão. Nunca escolhe uma das datas.
    """
    achadas = RE_VENCIMENTO_DARF.findall(texto or "")
    if len(achadas) < MINIMO_OCORRENCIAS_VENCIMENTO_DARF:
        return None
    if len(set(achadas)) != 1:
        return None
    try:
        return datetime.datetime.strptime(achadas[0], "%d/%m/%Y").date()
    except ValueError:
        return None
```

`logica.py` já importa `datetime` como módulo (linha 14) e `re` — não acrescente import.

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: PASS — 8 testes.

- [ ] **Step 5: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **425 testes, OK** (417 + 8).

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_lancamento_superlogica.py
git commit -m "Le o vencimento do DARF do texto, com conferencia entre as ocorrencias"
```

---

### Task 2: Documento de lançamento e o motivo de ficar de fora

**Files:**
- Modify: `logica.py` (depois de `vencimento_do_darf`)
- Test: `tests/test_lancamento_superlogica.py` (acrescentar)

**Interfaces:**
- Consumes: `vencimento_do_darf(texto)` (Task 1). Os dicts que a extração já produz: `extrair_dados_nfse(texto)` devolve, entre outras, as chaves `numero` (str), `competencia` (date), `cnpj_tomador` (str 14 dígitos), `valor_servico` (float), `valor_liquido` (float), `previdencia_retida` (float ou None), `contrib_sociais_retidas` (float ou None). `extrair_dados_boleto(texto, cadastro)` devolve `tipo` (`"Boleto bancário"` ou `"Arrecadação"`), `valor` (float ou None), `vencimento` (date ou None), `linha_digitavel` (str só dígitos), `documento_pagador` (str ou None — só preenchido quando está no cadastro).
- Produces:
  - constantes `TIPO_LANCAMENTO_NOTA = "nota fiscal"`, `TIPO_LANCAMENTO_BOLETO = "boleto"`, `TIPO_LANCAMENTO_ARRECADACAO = "arrecadacao"`
  - `documento_para_lancamento(nome_arquivo, dados_nfse=None, dados_boleto=None, texto="") -> dict | None` com as chaves `arquivo`, `tipo`, `documento`, `valor`, `vencimento`, `linha_digitavel`, `numero_documento`, `competencia`, `retencao` (bool)
  - `motivo_fora_do_lancamento(documento, cadastro) -> str` (`""` = pode ser lançado)

- [ ] **Step 1: Write the failing tests**

Acrescentar a `tests/test_lancamento_superlogica.py`, antes do `if __name__ == "__main__":`:

```python
from cadastro_teste import CADASTRO_TESTE as CADASTRO

KLOSTERS = "01195716000154"      # código 10004, ID SL 44
LAGO = "07945453000130"          # código 10590, SEM ID SL

NFSE_SEM_RETENCAO = {
    "numero": "12367", "competencia": datetime.date(2026, 8, 24),
    "cnpj_tomador": KLOSTERS, "valor_servico": 21.31, "valor_liquido": 21.31,
    "previdencia_retida": None, "contrib_sociais_retidas": None,
}
BOLETO = {
    "tipo": "Boleto bancário", "valor": 82.9,
    "vencimento": datetime.date(2026, 9, 10),
    "linha_digitavel": "23790472089000015979956006290706515650000008290",
    "documento_pagador": KLOSTERS,
}
ARRECADACAO = {
    "tipo": "Arrecadação", "valor": 58.51, "vencimento": None,
    "linha_digitavel": "858300000009585103852620610716262610830654448534",
    "documento_pagador": KLOSTERS,
}


class TestDocumentoParaLancamento(unittest.TestCase):

    def test_nota_fiscal(self):
        d = app.documento_para_lancamento("n.pdf", dados_nfse=NFSE_SEM_RETENCAO)
        self.assertEqual(d["tipo"], app.TIPO_LANCAMENTO_NOTA)
        self.assertEqual(d["documento"], KLOSTERS)
        self.assertEqual(d["valor"], 21.31)
        self.assertIsNone(d["vencimento"])          # vem do lote, não da nota
        self.assertIsNone(d["linha_digitavel"])
        self.assertEqual(d["numero_documento"], "12367")
        self.assertEqual(d["competencia"], datetime.date(2026, 8, 24))
        self.assertFalse(d["retencao"])

    def test_retencao_pela_diferenca_entre_servico_e_liquido(self):
        dados = dict(NFSE_SEM_RETENCAO, valor_servico=217.33, valor_liquido=207.22)
        self.assertTrue(app.documento_para_lancamento("n.pdf", dados_nfse=dados)["retencao"])

    def test_retencao_pelas_contribuicoes_sociais_retidas(self):
        #  O segundo sinal sozinho: serviço e líquido iguais, mas o campo de
        #  retenção preenchido. Basta um dos dois.
        dados = dict(NFSE_SEM_RETENCAO, contrib_sociais_retidas=10.11)
        self.assertTrue(app.documento_para_lancamento("n.pdf", dados_nfse=dados)["retencao"])

    def test_retencao_pela_previdencia_retida(self):
        dados = dict(NFSE_SEM_RETENCAO, previdencia_retida=2.34)
        self.assertTrue(app.documento_para_lancamento("n.pdf", dados_nfse=dados)["retencao"])

    def test_boleto(self):
        d = app.documento_para_lancamento("b.pdf", dados_boleto=BOLETO)
        self.assertEqual(d["tipo"], app.TIPO_LANCAMENTO_BOLETO)
        self.assertEqual(d["vencimento"], datetime.date(2026, 9, 10))
        self.assertEqual(d["linha_digitavel"], BOLETO["linha_digitavel"])
        self.assertIsNone(d["numero_documento"])
        self.assertIsNone(d["competencia"])

    def test_arrecadacao_pega_o_vencimento_do_texto(self):
        d = app.documento_para_lancamento("g.pdf", dados_boleto=ARRECADACAO, texto=DARF_OK)
        self.assertEqual(d["tipo"], app.TIPO_LANCAMENTO_ARRECADACAO)
        self.assertEqual(d["vencimento"], datetime.date(2026, 9, 18))
        self.assertIsNone(d["numero_documento"])
        self.assertIsNone(d["competencia"])

    def test_documento_nao_reconhecido(self):
        self.assertIsNone(app.documento_para_lancamento("x.pdf"))


class TestMotivoForaDoLancamento(unittest.TestCase):

    def doc(self, **mudancas):
        base = app.documento_para_lancamento("b.pdf", dados_boleto=BOLETO)
        base.update(mudancas)
        return base

    def test_documento_completo_pode_ser_lancado(self):
        self.assertEqual(app.motivo_fora_do_lancamento(self.doc(), CADASTRO), "")

    def test_nota_com_retencao(self):
        d = app.documento_para_lancamento(
            "n.pdf", dados_nfse=dict(NFSE_SEM_RETENCAO, contrib_sociais_retidas=10.11))
        self.assertIn("retenção", app.motivo_fora_do_lancamento(d, CADASTRO))

    def test_condominio_nao_identificado(self):
        self.assertIn("não identificado", app.motivo_fora_do_lancamento(
            self.doc(documento=None), CADASTRO))

    def test_documento_fora_do_cadastro(self):
        self.assertIn("não identificado", app.motivo_fora_do_lancamento(
            self.doc(documento="11222333000181"), CADASTRO))

    def test_condominio_sem_id_sl(self):
        self.assertIn("ID SL", app.motivo_fora_do_lancamento(
            self.doc(documento=LAGO), CADASTRO))

    def test_sem_valor(self):
        self.assertIn("valor", app.motivo_fora_do_lancamento(
            self.doc(valor=None), CADASTRO))

    def test_boleto_sem_vencimento(self):
        #  Fator 0000/9999: o boleto real do Itaú de referência é assim. Não
        #  se usa a data impressa na folha — ela seria uma data solta, sem a
        #  conferência que o DARF tem.
        self.assertIn("vencimento", app.motivo_fora_do_lancamento(
            self.doc(vencimento=None), CADASTRO))

    def test_darf_sem_vencimento_confirmado(self):
        d = app.documento_para_lancamento("g.pdf", dados_boleto=ARRECADACAO, texto="")
        self.assertIn("DARF", app.motivo_fora_do_lancamento(d, CADASTRO))

    def test_nota_sem_vencimento_proprio_pode_ser_lancada(self):
        #  A nota não tem vencimento próprio — ele vem do lote. Não é motivo
        #  de exclusão.
        d = app.documento_para_lancamento("n.pdf", dados_nfse=NFSE_SEM_RETENCAO)
        self.assertEqual(app.motivo_fora_do_lancamento(d, CADASTRO), "")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: FAIL — `module 'logica' has no attribute 'documento_para_lancamento'`.

- [ ] **Step 3: Write the implementation**

Em `logica.py`, depois de `vencimento_do_darf`:

```python
#  Tipos de documento que podem virar lançamento no Superlógica. É o que a
#  trava de pasta misturada compara: uma opção de lançamento vale para o lote
#  inteiro, então o lote tem de ser de um tipo só.
TIPO_LANCAMENTO_NOTA = "nota fiscal"
TIPO_LANCAMENTO_BOLETO = "boleto"
TIPO_LANCAMENTO_ARRECADACAO = "arrecadacao"

#  Diferença mínima entre valor do serviço e valor líquido para contar como
#  retenção. Os valores vêm como float da extração; um centavo de diferença
#  já é retenção de verdade.
TOLERANCIA_RETENCAO = 0.005


def documento_para_lancamento(nome_arquivo, dados_nfse=None, dados_boleto=None,
                              texto=""):
    """
    Normaliza um documento reconhecido pela extração num formato único,
    independente do tipo. Devolve None para documento não reconhecido.

    `texto` é o texto nativo do PDF — só é usado no DARF, para ler o
    vencimento (`vencimento_do_darf`), porque o código de arrecadação não
    traz data.

    `documento` é SEMPRE o CNPJ lido do documento: o do tomador na nota, o do
    pagador no boleto e no DARF (que `extrair_dados_boleto` só preenche
    quando está no cadastro). A identificação pelo nome do arquivo, que a
    planilha de extração usa como reserva nos boletos, NÃO vale aqui:
    lançamento é dinheiro, e essa origem é mais fraca que o CNPJ.
    """
    if dados_nfse:
        servico = dados_nfse.get("valor_servico")
        liquido = dados_nfse.get("valor_liquido")
        #  Dois sinais independentes, basta um: a diferença entre serviço e
        #  líquido, ou algum campo de retenção federal preenchido.
        diferenca = (servico is not None and liquido is not None
                     and abs(servico - liquido) > TOLERANCIA_RETENCAO)
        retencao = bool(diferenca
                        or dados_nfse.get("contrib_sociais_retidas")
                        or dados_nfse.get("previdencia_retida"))
        return {
            "arquivo": nome_arquivo,
            "tipo": TIPO_LANCAMENTO_NOTA,
            "documento": dados_nfse.get("cnpj_tomador"),
            "valor": servico,
            "vencimento": None,             # vem do lote
            "linha_digitavel": None,
            "numero_documento": dados_nfse.get("numero"),
            "competencia": dados_nfse.get("competencia"),
            "retencao": retencao,
        }

    if dados_boleto:
        arrecadacao = dados_boleto.get("tipo") == "Arrecadação"
        return {
            "arquivo": nome_arquivo,
            "tipo": TIPO_LANCAMENTO_ARRECADACAO if arrecadacao else TIPO_LANCAMENTO_BOLETO,
            "documento": dados_boleto.get("documento_pagador"),
            "valor": dados_boleto.get("valor"),
            "vencimento": (vencimento_do_darf(texto) if arrecadacao
                           else dados_boleto.get("vencimento")),
            "linha_digitavel": dados_boleto.get("linha_digitavel"),
            "numero_documento": None,
            "competencia": None,
            "retencao": False,
        }

    return None


def motivo_fora_do_lancamento(documento, cadastro):
    """
    Por que este documento NÃO pode entrar no arquivo do Superlógica, em
    frase para a coluna Observação. `""` significa que pode.

    A ordem importa: o primeiro motivo que se aplica é o que aparece. A nota
    com retenção vem primeiro porque é regra do usuário (lançar à mão), e não
    defeito do documento.
    """
    if documento.get("retencao"):
        return "Nota com retenção — lançar à mão no Superlógica"

    registro = (cadastro or {}).get(documento.get("documento")) if documento.get("documento") else None
    if registro is None:
        return "Condomínio não identificado pelo CNPJ — fora do lançamento"
    if not registro.get("id_sl"):
        return "Condomínio sem ID SL no cadastro — fora do lançamento"

    if documento.get("valor") is None:
        return "Sem valor no código — lançar à mão"

    if documento.get("vencimento") is None:
        if documento.get("tipo") == TIPO_LANCAMENTO_BOLETO:
            return "Boleto sem vencimento no código — lançar à mão"
        if documento.get("tipo") == TIPO_LANCAMENTO_ARRECADACAO:
            return "Vencimento do DARF não confirmado — lançar à mão"

    return ""
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: PASS — 24 testes (8 + 16).

- [ ] **Step 5: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **441 testes, OK** (425 + 16).

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_lancamento_superlogica.py
git commit -m "Normaliza o documento de lancamento e diz por que ele fica de fora"
```

---

### Task 3: Trava de pasta misturada e montagem dos lançamentos

**Files:**
- Modify: `logica.py` (depois de `motivo_fora_do_lancamento`)
- Test: `tests/test_lancamento_superlogica.py` (acrescentar)

**Interfaces:**
- Consumes: `documento_para_lancamento`, `motivo_fora_do_lancamento`, `TIPO_LANCAMENTO_*` (Task 2).
- Produces:
  - `COLUNAS_POR_DOCUMENTO = ("condominio", "valor", "vencimento", "linha_digitavel", "numero_documento", "competencia")`
  - `conferir_lote_homogeneo(documentos) -> str` (`""` = lote de um tipo só)
  - `montar_lancamentos(documentos, cadastro, vencimento_lote=None) -> list[dict]`, cada dict com **exatamente** as chaves de `COLUNAS_POR_DOCUMENTO`

- [ ] **Step 1: Write the failing tests**

Acrescentar a `tests/test_lancamento_superlogica.py`, antes do `if __name__`:

```python
class TestLoteHomogeneo(unittest.TestCase):

    def test_um_tipo_so_passa(self):
        docs = [app.documento_para_lancamento("a.pdf", dados_boleto=BOLETO),
                app.documento_para_lancamento("b.pdf", dados_boleto=BOLETO)]
        self.assertEqual(app.conferir_lote_homogeneo(docs), "")

    def test_lote_vazio_passa(self):
        self.assertEqual(app.conferir_lote_homogeneo([]), "")

    def test_nota_no_meio_de_boletos_recusa_e_diz_qual(self):
        docs = [app.documento_para_lancamento("boleto.pdf", dados_boleto=BOLETO),
                app.documento_para_lancamento("nota_perdida.pdf", dados_nfse=NFSE_SEM_RETENCAO)]
        mensagem = app.conferir_lote_homogeneo(docs)
        self.assertTrue(mensagem)
        self.assertIn("nota_perdida.pdf", mensagem)

    def test_boleto_e_darf_sao_tipos_diferentes(self):
        docs = [app.documento_para_lancamento("b.pdf", dados_boleto=BOLETO),
                app.documento_para_lancamento("g.pdf", dados_boleto=ARRECADACAO, texto=DARF_OK)]
        self.assertTrue(app.conferir_lote_homogeneo(docs))


class TestMontarLancamentos(unittest.TestCase):

    def test_boleto_vira_lancamento_com_as_seis_colunas(self):
        docs = [app.documento_para_lancamento("b.pdf", dados_boleto=BOLETO)]
        lancamentos = app.montar_lancamentos(docs, CADASTRO)
        self.assertEqual(len(lancamentos), 1)
        l = lancamentos[0]
        self.assertEqual(set(l), set(app.COLUNAS_POR_DOCUMENTO))
        self.assertEqual(l["condominio"], "44")
        self.assertEqual(l["valor"], 82.9)
        self.assertEqual(l["vencimento"], datetime.date(2026, 9, 10))
        self.assertEqual(l["linha_digitavel"], BOLETO["linha_digitavel"])
        self.assertIsNone(l["numero_documento"])
        self.assertIsNone(l["competencia"])

    def test_nota_usa_o_vencimento_do_lote(self):
        docs = [app.documento_para_lancamento("n.pdf", dados_nfse=NFSE_SEM_RETENCAO)]
        l = app.montar_lancamentos(docs, CADASTRO, datetime.date(2026, 9, 5))[0]
        self.assertEqual(l["vencimento"], datetime.date(2026, 9, 5))
        self.assertEqual(l["numero_documento"], "12367")
        self.assertEqual(l["competencia"], datetime.date(2026, 8, 24))
        self.assertIsNone(l["linha_digitavel"])

    def test_nota_sem_vencimento_do_lote_e_erro_de_quem_chama(self):
        docs = [app.documento_para_lancamento("n.pdf", dados_nfse=NFSE_SEM_RETENCAO)]
        with self.assertRaises(ValueError):
            app.montar_lancamentos(docs, CADASTRO)

    def test_documentos_com_motivo_ficam_de_fora(self):
        docs = [app.documento_para_lancamento("ok.pdf", dados_boleto=BOLETO),
                app.documento_para_lancamento("lago.pdf", dados_boleto=dict(BOLETO, documento_pagador=LAGO)),
                app.documento_para_lancamento("sem_venc.pdf", dados_boleto=dict(BOLETO, vencimento=None))]
        lancamentos = app.montar_lancamentos(docs, CADASTRO)
        self.assertEqual(len(lancamentos), 1)

    def test_ordem_dos_documentos_e_preservada(self):
        SAN_REMO = "08578541000103"      # ID SL 42
        docs = [app.documento_para_lancamento("1.pdf", dados_boleto=dict(BOLETO, documento_pagador=SAN_REMO)),
                app.documento_para_lancamento("2.pdf", dados_boleto=BOLETO)]
        self.assertEqual([l["condominio"] for l in app.montar_lancamentos(docs, CADASTRO)],
                         ["42", "44"])
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: FAIL — `module 'logica' has no attribute 'conferir_lote_homogeneo'`.

- [ ] **Step 3: Write the implementation**

Em `logica.py`, depois de `motivo_fora_do_lancamento`:

```python
#  As colunas do arquivo do Superlógica que variam por documento. São SEMPRE
#  escritas pelo programa (vazias quando o documento não tem o dado), nunca
#  herdadas do molde: foi deixar o vencimento no modelo que fez o Superlógica
#  gravar 01/01/1970 e recusar lançamentos (v6.14.0). Nomes já normalizados
#  por `_normalizar_cabecalho`. O `complemento` fica de fora de propósito —
#  o usuário o preenche à mão.
COLUNAS_POR_DOCUMENTO = ("condominio", "valor", "vencimento",
                         "linha_digitavel", "numero_documento", "competencia")

#  Como cada tipo aparece na mensagem da trava de pasta misturada.
ROTULOS_TIPO_LANCAMENTO = {
    TIPO_LANCAMENTO_NOTA: "nota fiscal",
    TIPO_LANCAMENTO_BOLETO: "boleto",
    TIPO_LANCAMENTO_ARRECADACAO: "guia de arrecadação (DARF)",
}


def conferir_lote_homogeneo(documentos):
    """
    Se o lote tem documentos de mais de um tipo, devolve a mensagem que diz
    quais; senão, `""`.

    Existe porque a opção de lançamento vale para o lote inteiro: uma nota
    perdida numa pasta de boletos do sindicato seria lançada com o
    fornecedor do sindicato. Só conta documento reconhecido — o
    "Detalhamento do Faturamento" que vem nas pastas da F&F não é nota nem
    boleto, e nem chega aqui (`documento_para_lancamento` devolve None).
    """
    por_tipo = {}
    for documento in documentos:
        por_tipo.setdefault(documento["tipo"], []).append(documento["arquivo"])
    if len(por_tipo) <= 1:
        return ""

    partes = []
    for tipo, arquivos in sorted(por_tipo.items(), key=lambda item: -len(item[1])):
        exemplos = ", ".join(arquivos[:3])
        resto = f" e mais {len(arquivos) - 3}" if len(arquivos) > 3 else ""
        partes.append(f"{len(arquivos)} {ROTULOS_TIPO_LANCAMENTO.get(tipo, tipo)}: {exemplos}{resto}")
    return ("A pasta mistura tipos de documento, e a opção de lançamento vale "
            "para o lote inteiro — o arquivo do Superlógica não foi gerado.\n\n"
            + "\n".join(partes))


def montar_lancamentos(documentos, cadastro, vencimento_lote=None):
    """
    Lançamentos para `gerar_planilha_despesas`: um dict por documento que
    PODE ser lançado, na ordem recebida, com exatamente as chaves de
    COLUNAS_POR_DOCUMENTO. Documento com motivo em `motivo_fora_do_lancamento`
    é pulado — o motivo já foi para a planilha de extração.

    `vencimento_lote` é o vencimento das notas fiscais (nota não tem
    vencimento próprio). Faltar com nota no lote é erro de quem chama.
    """
    lancamentos = []
    for documento in documentos:
        if motivo_fora_do_lancamento(documento, cadastro):
            continue
        if documento["tipo"] == TIPO_LANCAMENTO_NOTA:
            if vencimento_lote is None:
                raise ValueError("Nota fiscal no lote exige o vencimento do lote.")
            vencimento = vencimento_lote
        else:
            vencimento = documento["vencimento"]
        lancamentos.append({
            "condominio": cadastro[documento["documento"]]["id_sl"],
            "valor": documento["valor"],
            "vencimento": vencimento,
            "linha_digitavel": documento["linha_digitavel"],
            "numero_documento": documento["numero_documento"],
            "competencia": documento["competencia"],
        })
    return lancamentos
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: PASS — 33 testes (24 + 9).

- [ ] **Step 5: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **450 testes, OK** (441 + 9).

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_lancamento_superlogica.py
git commit -m "Trava de pasta misturada e montagem dos lancamentos do Superlogica"
```

---

### Task 4: `gerar_planilha_despesas` aceita colunas por documento; pasta de modelos

**Files:**
- Modify: `logica.py:3046-3144` (`gerar_planilha_despesas`)
- Modify: `logica.py` (constante e função de listagem, logo antes de `gerar_planilha_despesas`)
- Test: `tests/test_lancamento_superlogica.py` (acrescentar)

**Interfaces:**
- Consumes: `COLUNAS_POR_DOCUMENTO` (Task 3); existentes `_normalizar_cabecalho`, `COLUNA_DESPESA_CONDOMINIO` (`"condominio"`), `COLUNA_DESPESA_VALOR` (`"valor"`), `COLUNAS_DATA_DESPESAS`, `FORMATO_DATA_DESPESAS`, `LINHA_MOLDE_DESPESAS`.
- Produces:
  - `gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos, vencimento=None, chave=None)` — **mesma assinatura**; cada item de `lancamentos` pode ser `(id_sl, valor)` (como hoje) ou um `dict {coluna_normalizada: valor}`
  - `PASTA_MODELOS_LANCAMENTO = "modelos_superlogica"`
  - `listar_modelos_lancamento(pasta) -> list[tuple[str, str]]` — `(rótulo, caminho)` ordenado pelo rótulo

- [ ] **Step 1: Write the failing tests**

Acrescentar a `tests/test_lancamento_superlogica.py`, antes do `if __name__` (imports no topo do bloco):

```python
import shutil
import tempfile

from openpyxl import Workbook, load_workbook

#  Modelo sintético com as colunas que importam aqui. O molde traz lixo nas
#  colunas por documento DE PROPÓSITO: o programa tem de sobrescrevê-las,
#  nunca herdar.
CABECALHO = ["condomínio", "vencimento", "competencia", "fornecedor",
             "conta_categoria", "numero_documento", "complemento", "valor",
             "linha_digitavel", "chave"]
MOLDE = [None, None, None, "4521", "2.1.49 Envio Informações E-Social",
         "LIXO DO MOLDE", "Exames Médicos", None, "LIXO DO MOLDE", 46]


def montar_modelo(caminho):
    wb = Workbook()
    sheet = wb.active
    sheet.append(CABECALHO)
    sheet.append(MOLDE)
    wb.save(caminho)
    return caminho


class TestGerarComColunasPorDocumento(unittest.TestCase):

    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.modelo = montar_modelo(os.path.join(self.pasta, "modelo.xlsx"))
        self.saida = os.path.join(self.pasta, "saida.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def gerar(self, lancamentos, **kwargs):
        app.gerar_planilha_despesas(self.modelo, self.saida, lancamentos, **kwargs)
        sheet = load_workbook(self.saida).active
        colunas = {app._normalizar_cabecalho(c.value): c.column for c in sheet[1]}
        return sheet, colunas

    def lancamento(self, **mudancas):
        base = {"condominio": "44", "valor": 82.9,
                "vencimento": datetime.date(2026, 9, 10),
                "linha_digitavel": "23790472089000015979956006290706515650000008290",
                "numero_documento": None, "competencia": None}
        base.update(mudancas)
        return base

    def test_escreve_as_colunas_por_documento(self):
        sheet, col = self.gerar([self.lancamento()])
        self.assertEqual(sheet.cell(2, col["condominio"]).value, "44")
        self.assertEqual(sheet.cell(2, col["valor"]).value, 82.9)
        self.assertEqual(sheet.cell(2, col["linha_digitavel"]).value,
                         "23790472089000015979956006290706515650000008290")

    def test_data_sai_como_data_de_verdade(self):
        sheet, col = self.gerar([self.lancamento()])
        celula = sheet.cell(2, col["vencimento"])
        self.assertTrue(celula.is_date)
        self.assertEqual(celula.number_format, app.FORMATO_DATA_DESPESAS)

    def test_none_apaga_o_que_o_molde_tinha(self):
        #  "—" no spec é célula vazia, mesmo que o molde tenha algo ali.
        sheet, col = self.gerar([self.lancamento(linha_digitavel=None)])
        self.assertIsNone(sheet.cell(2, col["linha_digitavel"]).value)
        self.assertIsNone(sheet.cell(2, col["numero_documento"]).value)

    def test_colunas_do_molde_continuam_sendo_copiadas(self):
        sheet, col = self.gerar([self.lancamento(), self.lancamento()])
        for linha in (2, 3):
            self.assertEqual(sheet.cell(linha, col["fornecedor"]).value, "4521")
            self.assertEqual(sheet.cell(linha, col["complemento"]).value, "Exames Médicos")
            self.assertEqual(sheet.cell(linha, col["chave"]).value, 46)

    def test_chave_do_lote_continua_valendo(self):
        sheet, col = self.gerar([self.lancamento()], chave=99)
        self.assertEqual(sheet.cell(2, col["chave"]).value, 99)

    def test_modelo_sem_uma_coluna_por_documento_avisa(self):
        wb = Workbook()
        wb.active.append(["condomínio", "valor"])
        wb.active.append([None, None])
        wb.save(self.modelo)
        with self.assertRaises(ValueError) as erro:
            self.gerar([self.lancamento()])
        self.assertIn("linha_digitavel", str(erro.exception))

    def test_formato_antigo_de_tupla_continua_funcionando(self):
        #  A aba 3 continua chamando com (id_sl, valor).
        sheet, col = self.gerar([("44", 57.75)])
        self.assertEqual(sheet.cell(2, col["condominio"]).value, "44")
        self.assertEqual(sheet.cell(2, col["valor"]).value, 57.75)


class TestListarModelosLancamento(unittest.TestCase):

    def setUp(self):
        self.pasta = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def criar(self, nome):
        open(os.path.join(self.pasta, nome), "wb").close()

    def test_lista_os_xlsx_pelo_nome_sem_extensao_em_ordem(self):
        self.criar("Sindicato - Contr. Assistencial.xlsx")
        self.criar("FF - Exames Medicos.xlsx")
        rotulos = [r for r, _ in app.listar_modelos_lancamento(self.pasta)]
        self.assertEqual(rotulos, ["FF - Exames Medicos", "Sindicato - Contr. Assistencial"])

    def test_devolve_o_caminho_do_arquivo(self):
        self.criar("FF - PCMSO.xlsx")
        (_, caminho), = app.listar_modelos_lancamento(self.pasta)
        self.assertEqual(caminho, os.path.join(self.pasta, "FF - PCMSO.xlsx"))

    def test_ignora_o_arquivo_de_trava_do_excel(self):
        #  Com um modelo aberto no Excel aparece "~$Nome.xlsx" na pasta; não
        #  é modelo e não pode virar opção.
        self.criar("FF - PCMSO.xlsx")
        self.criar("~$FF - PCMSO.xlsx")
        self.assertEqual(len(app.listar_modelos_lancamento(self.pasta)), 1)

    def test_ignora_o_que_nao_e_xlsx(self):
        self.criar("leia-me.txt")
        self.criar("antigo.xls")
        self.assertEqual(app.listar_modelos_lancamento(self.pasta), [])

    def test_pasta_que_nao_existe(self):
        self.assertEqual(app.listar_modelos_lancamento(os.path.join(self.pasta, "nao")), [])
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_lancamento_superlogica -v`
Expected: FAIL — os testes de dict falham (o gerador hoje desempacota `(id_sl, valor)`) e `listar_modelos_lancamento` não existe.

- [ ] **Step 3: Add the models folder listing**

Em `logica.py`, logo **antes** de `def gerar_planilha_despesas`:

```python
#  Pasta, ao lado do programa, com um .xlsx por opção de lançamento da aba 2.
#  O nome do arquivo (sem extensão) é o texto do seletor; a linha 2 é o
#  molde, igual ao `modelo_despesas.xlsx` da aba 3. O zip de instalação NÃO
#  traz modelos: são do usuário, e um pacote novo sobrescreveria as edições.
PASTA_MODELOS_LANCAMENTO = "modelos_superlogica"


def listar_modelos_lancamento(pasta):
    """
    Opções de lançamento disponíveis: `[(rótulo, caminho), ...]`, em ordem
    alfabética do rótulo. Pasta inexistente devolve lista vazia.

    Ignora o `~$Nome.xlsx` que o Excel cria enquanto o modelo está aberto —
    não é modelo, e virar opção faria o gerador tentar abrir um arquivo de
    trava.
    """
    if not os.path.isdir(pasta):
        return []
    modelos = []
    for nome in os.listdir(pasta):
        if nome.startswith("~$") or not nome.lower().endswith(".xlsx"):
            continue
        modelos.append((os.path.splitext(nome)[0], os.path.join(pasta, nome)))
    return sorted(modelos, key=lambda modelo: modelo[0].lower())
```

- [ ] **Step 4: Teach the generator the per-document columns**

Em `gerar_planilha_despesas`:

(a) Logo depois do `if not lancamentos: raise ...`, acrescentar a normalização:

```python
    #  Cada lançamento vira um dict {coluna: valor}. A aba 3 continua mandando
    #  (id_sl, valor); a aba 2 manda as seis colunas de COLUNAS_POR_DOCUMENTO.
    campos_por_linha = []
    for lancamento in lancamentos:
        if isinstance(lancamento, dict):
            campos_por_linha.append(dict(lancamento))
        else:
            id_sl, valor = lancamento
            campos_por_linha.append({COLUNA_DESPESA_CONDOMINIO: id_sl,
                                     COLUNA_DESPESA_VALOR: valor})
```

(b) Na montagem de `obrigatorias`, logo depois de
`obrigatorias = [COLUNA_DESPESA_CONDOMINIO, COLUNA_DESPESA_VALOR]`, acrescentar:

```python
    for campos in campos_por_linha:
        for nome in campos:
            if nome not in obrigatorias:
                obrigatorias.append(nome)
```

(c) Remover as duas linhas que deixam de ser usadas:

```python
    coluna_condominio = colunas[COLUNA_DESPESA_CONDOMINIO]
    coluna_valor = colunas[COLUNA_DESPESA_VALOR]
```

(d) Trocar o laço de escrita. O trecho atual:

```python
    for indice, (id_sl, valor) in enumerate(lancamentos):
        numero_linha = LINHA_MOLDE_DESPESAS + indice
        for coluna, (valor_molde, estilo, formato) in enumerate(molde, start=1):
            celula = sheet.cell(row=numero_linha, column=coluna)
            celula.value = valor_molde
            celula._style = copy.copy(estilo)
            if formato:
                celula.number_format = formato
        sheet.cell(row=numero_linha, column=coluna_condominio).value = id_sl
        sheet.cell(row=numero_linha, column=coluna_valor).value = float(valor)
```

passa a ser:

```python
    for indice, campos in enumerate(campos_por_linha):
        numero_linha = LINHA_MOLDE_DESPESAS + indice
        for coluna, (valor_molde, estilo, formato) in enumerate(molde, start=1):
            celula = sheet.cell(row=numero_linha, column=coluna)
            celula.value = valor_molde
            celula._style = copy.copy(estilo)
            if formato:
                celula.number_format = formato
        #  Colunas do lançamento vencem o molde — inclusive com None, que
        #  apaga o que o molde tivesse ali.
        for nome, valor in campos.items():
            celula = sheet.cell(row=numero_linha, column=colunas[nome])
            if nome == COLUNA_DESPESA_VALOR and valor is not None:
                celula.value = float(valor)
            elif nome in COLUNAS_DATA_DESPESAS and isinstance(valor, datetime.date):
                celula.value = valor
                celula.number_format = FORMATO_DATA_DESPESAS
            else:
                celula.value = valor
```

O restante do laço (o `if vencimento is not None:` e o `if chave is not None:`) fica **como está**, logo depois.

(e) A remoção das sobras usa `len(lancamentos)` — continua correta, não mexa.

(f) Atualizar a primeira frase do docstring, que diz `` `lancamentos` é [(id_sl, valor), ...] ``, para:

```
    Gera a planilha de importação de despesas do Superlógica a partir do
    modelo do usuário. `lancamentos` é uma lista, na ordem de saída, de
    `(id_sl, valor)` (aba 3) ou de dicts `{coluna_normalizada: valor}` (aba
    2, ver COLUNAS_POR_DOCUMENTO) — as colunas do dict vencem o molde,
    inclusive com None, que deixa a célula vazia.
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m unittest tests.test_lancamento_superlogica tests.test_despesas_superlogica -v`
Expected: PASS — 45 testes do arquivo novo (33 + 12) e **todos** os de `test_despesas_superlogica.py` sem alteração.

- [ ] **Step 6: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **462 testes, OK** (450 + 12).

- [ ] **Step 7: Commit**

```bash
git add logica.py tests/test_lancamento_superlogica.py
git commit -m "Gerador do Superlogica aceita colunas por documento; lista a pasta de modelos"
```

---

### Task 5: Seletor na aba 2 e geração ao fim da extração

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` — import de `logica` (topo); `_montar_aba_extracao` (linhas ~2288-2443); `_atualizar_botao_extrair` (~2450); `_iniciar_extracao` (~2467); `_extrair_em_thread` (~2490); método novo depois de `_concluir_extracao` (~2580)

**Interfaces:**
- Consumes: `pasta_base()`, `PASTA_MODELOS_LANCAMENTO`, `listar_modelos_lancamento`, `documento_para_lancamento`, `motivo_fora_do_lancamento`, `conferir_lote_homogeneo`, `montar_lancamentos`, `TIPO_LANCAMENTO_NOTA`, `gerar_planilha_despesas` (Tasks 1-4). Existentes na interface: `self._pedir_chave_despesas()` → chave ou `None` (cancelou); `self._pedir_vencimento_despesas()` → `datetime.date` ou `None` (cancelou). Os dois já usam `_janela_para_dialogo`, que cai na janela principal quando o painel da aba 3 não está aberto — servem na aba 2 sem mudança.
- Produces: nada consumido por outras tarefas.

- [ ] **Step 1: Import the new functions**

Na lista `from logica import (...)` no topo do arquivo, acrescentar:

```python
    PASTA_MODELOS_LANCAMENTO, listar_modelos_lancamento,
    documento_para_lancamento, motivo_fora_do_lancamento,
    conferir_lote_homogeneo, montar_lancamentos, TIPO_LANCAMENTO_NOTA,
```

Confirme que `gerar_planilha_despesas` e `pasta_base` já estão importados (a aba 3 usa os dois); se algum faltar, acrescente.

Logo abaixo do bloco de imports, junto das outras constantes de módulo (perto de `QUALIDADE_LEITURA_PROTOCOLO`), acrescentar:

```python
#  Opção do seletor da aba 2 que desliga o lançamento no Superlógica: a aba
#  funciona exatamente como antes.
SEM_LANCAMENTO = "Nenhum"
```

- [ ] **Step 2: Build the selector in tab 2**

Em `_montar_aba_extracao`, o botão de extrair está em `row=4`, e depois vêm `explicacao` (`row=5`), `self.barra_extracao` (`row=6`) e `self.label_status_extracao` (`row=7`). Some 1 ao `row` desses três, **de baixo para cima** para os números não colidirem no meio da edição:

| Widget | De | Para |
|---|---|---|
| `self.label_status_extracao.grid(...)` | `row=7` | `row=8` |
| `self.barra_extracao.grid(...)` | `row=6` | `row=7` |
| `explicacao.grid(...)` | `row=5` | `row=6` |

Depois, logo **depois** de `self.botao_extrair.grid(row=4, ...)`, acrescentar:

```python
        # --- Lançamento no Superlógica (opção do lote) ---
        #  Cada opção é um .xlsx em modelos_superlogica/, ao lado do programa;
        #  o nome do arquivo é o texto do seletor. "Nenhum" deixa a aba como
        #  sempre foi.
        self.opcao_lancamento = tk.StringVar(value=SEM_LANCAMENTO)
        self._modelos_lancamento = {}

        linha_lancamento = registrar(
            ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        linha_lancamento.grid(row=5, column=0, sticky="w", pady=(0, 8))

        registrar(
            ctk.CTkLabel(linha_lancamento, text="Lançamento no Superlógica:",
                         font=(fonte, 13), text_color=tema["texto"]),
            {"text_color": "texto"},
        ).pack(side="left", padx=(0, 8))

        self.menu_lancamento = ctk.CTkOptionMenu(
            linha_lancamento, values=[SEM_LANCAMENTO], variable=self.opcao_lancamento,
            corner_radius=0, fg_color=tema["superficie"], button_color=tema["borda_forte"],
            button_hover_color=tema["borda_forte"], text_color=tema["texto"],
            dropdown_fg_color=tema["superficie"], dropdown_text_color=tema["texto"],
            font=(fonte, 13),
        )
        self.menu_lancamento.pack(side="left")

        self.label_aviso_modelos = registrar(
            ctk.CTkLabel(linha_lancamento, text="", font=(fonte, 12),
                         text_color=tema["texto_terciario"]),
            {"text_color": "texto_terciario"},
        )
        self.label_aviso_modelos.pack(side="left", padx=(12, 0))
```

`button_hover_color` usa `borda_forte`, **não** `acento`: o seletor não é botão primário nem resolve pendência.

- [ ] **Step 3: Reload the options whenever the tab refreshes**

Acrescentar, logo depois de `_atualizar_botao_extrair`:

```python
    def _recarregar_opcoes_lancamento(self):
        """
        Relê a pasta de modelos e atualiza o seletor. Chamado sempre que a
        aba se atualiza (ao escolher pasta ou destino), para um modelo recém-
        -colocado na pasta aparecer sem reabrir o programa.

        Mantém a opção escolhida se o arquivo dela ainda existir; senão volta
        para "Nenhum" — gerar com um modelo que sumiu daria erro no fim da
        extração, depois de o usuário já ter esperado o lote inteiro.
        """
        menu = getattr(self, "menu_lancamento", None)
        if menu is None:
            return
        pasta = os.path.join(pasta_base(), PASTA_MODELOS_LANCAMENTO)
        try:
            os.makedirs(pasta, exist_ok=True)
        except Exception:
            pass
        self._modelos_lancamento = dict(listar_modelos_lancamento(pasta))
        menu.configure(values=[SEM_LANCAMENTO] + list(self._modelos_lancamento))
        if self.opcao_lancamento.get() not in self._modelos_lancamento:
            self.opcao_lancamento.set(SEM_LANCAMENTO)
        self.label_aviso_modelos.configure(
            text="" if self._modelos_lancamento else
            f"Nenhum modelo em {pasta}")
```

E na **última linha** do corpo de `_atualizar_botao_extrair`, acrescentar a chamada:

```python
        self._recarregar_opcoes_lancamento()
```

- [ ] **Step 4: Pass the chosen model to the worker**

Em `_iniciar_extracao`, logo antes de criar a thread:

```python
        #  None quando "Nenhum": a aba funciona como antes.
        modelo = self._modelos_lancamento.get(self.opcao_lancamento.get())
```

e trocar `args=(pasta, destino)` por `args=(pasta, destino, modelo)`.

Na assinatura: `def _extrair_em_thread(self, pasta, destino, modelo=None):`.

- [ ] **Step 5: Build documents and annotate the reason in the worker**

Em `_extrair_em_thread`:

(a) Junto de `linhas = []`, acrescentar:

```python
        documentos = []       # só usado com uma opção de lançamento escolhida
        fora_do_lancamento = 0
```

(b) No início do corpo do laço, junto de `dados = None`, acrescentar `texto = ""` — ele passa a ser lido fora do `try`.

(c) Logo **antes** de `if dados_boleto is not None:` (a montagem das linhas), acrescentar:

```python
            #  Com uma opção de lançamento escolhida, cada documento
            #  reconhecido vira documento de lançamento, e o motivo de ficar
            #  de fora vai para a Observação da planilha de extração — é ali
            #  que o usuário descobre o que lançar à mão.
            motivo = ""
            if modelo is not None:
                documento = documento_para_lancamento(nome, dados, dados_boleto, texto)
                if documento is not None:
                    documentos.append(documento)
                    motivo = motivo_fora_do_lancamento(documento, self.cadastro)
                    if motivo:
                        fora_do_lancamento += 1
```

(d) Passar o motivo às duas montagens de linha:

```python
                linha = linha_planilha_boleto(nome, dados_boleto, self.cadastro, motivo)
```

e

```python
                linha = linha_planilha_nfse(nome, dados, self.cadastro, observacao or motivo)
```

(`observacao` só é preenchida quando o documento **não** foi lido; quando foi, é `""` e o motivo entra no lugar.)

(e) No fim, trocar a última linha

```python
        self.after(0, lambda: self._concluir_extracao(destino, resumo))
```

por:

```python
        if modelo is None:
            self.after(0, lambda: self._concluir_extracao(destino, resumo))
        else:
            self.after(0, lambda: self._gerar_lancamento_superlogica(
                modelo, documentos, fora_do_lancamento, destino, resumo))
```

- [ ] **Step 6: Generate the Superlógica file on the main thread**

Acrescentar, logo depois de `_concluir_extracao`:

```python
    def _gerar_lancamento_superlogica(self, modelo, documentos, fora, destino, resumo):
        """
        Gera o arquivo de importação do Superlógica ao fim da extração.

        As perguntas vêm AQUI, e não no início como na aba 3: lá o vencimento
        vai carimbado no PDF durante o processamento; aqui nada é carimbado, e
        esperar o fim permite perguntar o vencimento só quando o lote de fato
        tem nota fiscal. Qualquer cancelamento termina sem o arquivo do
        Superlógica — a planilha de extração já foi gravada e não se perde.
        """
        if not documentos:
            messagebox.showinfo(
                "Lançamento no Superlógica",
                "Nenhuma nota fiscal nem boleto foi reconhecido no lote — o "
                "arquivo do Superlógica não foi gerado.")
            self._concluir_extracao(destino, resumo)
            return

        mistura = conferir_lote_homogeneo(documentos)
        if mistura:
            messagebox.showwarning("Lançamento não gerado", mistura)
            self._concluir_extracao(destino, resumo)
            return

        chave = self._pedir_chave_despesas()
        if chave is None:
            self._concluir_extracao(destino, resumo)
            return

        vencimento = None
        if any(d["tipo"] == TIPO_LANCAMENTO_NOTA for d in documentos):
            vencimento = self._pedir_vencimento_despesas()
            if vencimento is None:
                self._concluir_extracao(destino, resumo)
                return

        lancamentos = montar_lancamentos(documentos, self.cadastro, vencimento)
        if not lancamentos:
            messagebox.showwarning(
                "Lançamento não gerado",
                "Nenhum documento pôde ser lançado — o motivo de cada um está "
                "na coluna Observação da planilha de extração.")
            self._concluir_extracao(destino, resumo)
            return

        destino_superlogica = os.path.splitext(destino)[0] + " - superlogica.xlsx"
        if os.path.isfile(destino_superlogica) and not messagebox.askyesno(
            "Substituir arquivo",
            f"Este arquivo já existe e será substituído:\n\n{destino_superlogica}\n\nContinuar?"
        ):
            self._concluir_extracao(destino, resumo)
            return

        try:
            gerar_planilha_despesas(modelo, destino_superlogica, lancamentos, chave=chave)
        except Exception as e:
            messagebox.showerror(
                "Erro ao gerar",
                f"Não foi possível gerar o arquivo do Superlógica:\n{e}\n\n"
                "Se ele estiver aberto no Excel, feche e tente de novo.")
            self._concluir_extracao(destino, resumo)
            return

        texto = f"{len(lancamentos)} lançamento(s) gerado(s)"
        if fora:
            texto += (f", {fora} documento(s) ficaram de fora — veja a coluna "
                      "Observação na planilha de extração")
        #  Não oferece abrir o arquivo do Superlógica: abrir e salvar no Excel
        #  é o gesto que já fez a coluna de vencimento virar número e o
        #  Superlógica gravar 01/01/1970 (v6.17.0).
        messagebox.showinfo(
            "Arquivo do Superlógica pronto",
            f"{texto}.\n\nSalvo em:\n{destino_superlogica}\n\n"
            "Importe esse arquivo direto no Superlógica, sem abrir no Excel.")
        self._concluir_extracao(destino, resumo)
```

- [ ] **Step 7: Verify it imports and the suite still passes**

Run: `python -c "import identificacao_por_cnpj_6_0; print('ok')"`
Expected: `ok`

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **462 testes, OK** (esta tarefa não acrescenta teste — é interface).

- [ ] **Step 8: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Seletor de lancamento na aba 2 e geracao do arquivo do Superlogica"
```

---

### Task 6: Validação contra as pastas reais, CLAUDE.md e versão

**Files:**
- Create: `tests/_validar_lancamento.py`
- Modify: `CLAUDE.md` (seção nova antes de `## Contagem dos Protocolos dos Correios (aba 3)`)
- Modify: `identificacao_por_cnpj_6_0.py:212` (`self.title`)

**Interfaces:**
- Consumes: `extrair_texto_pdf`, `extrair_dados_nfse`, `extrair_dados_boleto`, `documento_para_lancamento`, `motivo_fora_do_lancamento`, `conferir_lote_homogeneo`, `montar_lancamentos`, `carregar_cadastro`, `_codigos_do_cadastro`.

- [ ] **Step 1: Write the manual validator**

Criar `tests/_validar_lancamento.py`:

```python
# -*- coding: utf-8 -*-
"""
Valida o lançamento no Superlógica contra as pastas reais.

Fora da suíte (prefixo `_`): depende de PDFs fora do repositório.
Reexecutável. Não grava nada — só monta os lançamentos e confere.

Rode:  python tests/_validar_lancamento.py
"""
import collections
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
CADASTRO = logica.carregar_cadastro("cadastro_condominios.xlsx")
CODIGO_PARA_DOC = logica._codigos_do_cadastro(CADASTRO)

LOTES = [
    ("Sindicato", r"C:\Users\Desktop\Music\sindicato\boletos",
     lambda nome: CODIGO_PARA_DOC.get((re.match(r"(\d{5})_", nome) or [None, None])[1])),
    ("DARF", r"C:\Users\Desktop\Music\inss",
     lambda nome: (re.search(r"GuiaPagamento_(\d{14})_", nome) or [None, None])[1]),
    ("F&F fevereiro", r"C:\Users\Desktop\Downloads\02 Fevereiro 2026-20260401T035105Z-1-001",
     lambda nome: CODIGO_PARA_DOC.get((re.search(r"\b(1\d{4})\b", nome) or [None, None])[1])),
]


def documentos_do_lote(pasta):
    documentos = []
    for raiz, _, arquivos in os.walk(pasta):
        for nome in sorted(arquivos):
            if not nome.lower().endswith(".pdf"):
                continue
            caminho = os.path.join(raiz, nome)
            try:
                texto = logica.extrair_texto_pdf(caminho)
            except Exception:
                continue
            nfse = logica.extrair_dados_nfse(texto)
            boleto = None if nfse else logica.extrair_dados_boleto(texto, CADASTRO)
            doc = logica.documento_para_lancamento(nome, nfse, boleto, texto)
            if doc:
                documentos.append(doc)
    return documentos


for rotulo, pasta, esperado_do_nome in LOTES:
    if not os.path.isdir(pasta):
        print(f"\n== {rotulo}: pasta não encontrada ({pasta})")
        continue
    docs = documentos_do_lote(pasta)
    motivos = collections.Counter(logica.motivo_fora_do_lancamento(d, CADASTRO) or "(lançado)"
                                  for d in docs)
    mistura = logica.conferir_lote_homogeneo(docs)
    vencimento_lote = datetime.date(2026, 9, 5)
    lancamentos = logica.montar_lancamentos(docs, CADASTRO, vencimento_lote) if not mistura else []
    divergentes = [d["arquivo"] for d in docs
                   if not logica.motivo_fora_do_lancamento(d, CADASTRO)
                   and esperado_do_nome(d["arquivo"])
                   and esperado_do_nome(d["arquivo"]) != d["documento"]]
    print(f"\n== {rotulo}: {len(docs)} documentos reconhecidos")
    print(f"   pasta misturada ........ {'SIM: ' + mistura.splitlines()[0] if mistura else 'não'}")
    for motivo, n in motivos.most_common():
        print(f"   {motivo:55} {n}")
    print(f"   lançamentos montados ... {len(lancamentos)}")
    print(f"   CONDOMÍNIO DIVERGE DO NOME DO ARQUIVO: {len(divergentes)}   <- tem que ser 0")
    for nome in divergentes[:5]:
        print(f"      ! {nome}")
```

- [ ] **Step 2: Run the validator**

Run: `python tests/_validar_lancamento.py`

Expected:
- **Sindicato**: 284 documentos, pasta não misturada, 284 lançados, **0 divergências**.
- **DARF**: 12 documentos, pasta não misturada, 12 lançados (vencimento lido em todos), **0 divergências**.
- **F&F fevereiro**: pasta não misturada; as notas **com retenção** aparecem como "Nota com retenção — lançar à mão no Superlógica" (na medição de desenho foram 228 de 594); as demais lançadas; **0 divergências**.

Se qualquer número divergir disso, **pare e relate** — não ajuste o validador para casar com o esperado.

**Uma exceção conhecida, só na F&F:** os dados reais têm arquivos com o **nome errado** — o caso registrado no `CLAUDE.md` é `10871 Esperança`, que traz o CNPJ do `10872 PAIVA`. Ali o CNPJ está certo e o nome do arquivo não. Divergência na F&F, portanto, **não é automaticamente defeito**: relate cada uma com o nome do arquivo, o código do nome e o código do CNPJ, para o controlador conferir. No sindicato e no DARF a expectativa continua sendo 0.

- [ ] **Step 3: Document it in CLAUDE.md**

Inserir **antes** de `## Contagem dos Protocolos dos Correios (aba 3)`:

````markdown
## Lançamento no Superlógica direto da extração (aba 2, v6.21.0)

A aba 2 gera, junto da planilha de extração, o arquivo de importação de
despesas do Superlógica para **boletos do sindicato, DARF do DCTFWeb e notas
da F&F**. Um seletor escolhe a **opção de lançamento do lote** ("Nenhum" =
aba como antes). Uma opção por lote, por decisão do usuário.

**Cada opção é um `.xlsx` em `modelos_superlogica/`**, ao lado do programa:
nome do arquivo = texto do seletor, linha 2 = molde — o mesmo mecanismo do
`modelo_despesas.xlsx` da aba 3. Fornecedor e favorecido podem ir pelo **ID
do Superlógica** (o importador aceita); `conta_categoria` continua como
`código nome`. O zip **não traz** modelos — são do usuário. O seletor ignora
o `~$Nome.xlsx` que o Excel cria com o modelo aberto. A alternativa de
guardar as opções no `config.json` com editor em Configurações foi preferida
pelo usuário para o futuro, mas adiada.

**Seis colunas são escritas pelo programa, nunca herdadas do molde**
(`COLUNAS_POR_DOCUMENTO`): `condominio`, `valor`, `vencimento`,
`linha_digitavel`, `numero_documento`, `competencia` — vazias quando o
documento não tem o dado, mesmo que o molde tenha algo. O **`complemento`
não é escrito**: o usuário preenche à mão (no modelo ou no Superlógica —
nunca abrindo a planilha gerada no Excel, que é o gesto que já fez o
Superlógica gravar 01/01/1970).

| Coluna | Boleto do sindicato | DARF | Nota da F&F |
|---|---|---|---|
| `condominio` (ID SL) | CNPJ do pagador | CNPJ | CNPJ do tomador |
| `valor` | código de barras | código de barras | valor do serviço |
| `vencimento` | código de barras | texto (≥2 ocorrências iguais) | o do lote |
| `linha_digitavel` | 47 dígitos | 48 dígitos | — |
| `numero_documento` | — | — | nº da NFS-e |
| `competencia` | — | — | competência da nota |

**O condomínio só vale pelo CNPJ lido do documento.** A planilha de extração
usa o nome do arquivo como reserva nos boletos; o lançamento não — é
dinheiro, e essa origem é mais fraca.

**Vencimento do DARF**: o código de arrecadação não traz data. Ela está no
texto com rótulo em três lugares (`Pagar este documento até`,
`Vencimento:` e `Pagar até:`), e `vencimento_do_darf` exige **pelo menos
duas ocorrências e todas iguais** — divergindo ou faltando, o DARF vai para o
manual. A data de emissão ("SENDA ... 18/09/2026 09:46") não tem rótulo e
não entra na conta.

**O que fica de fora** vai para a Observação da planilha de extração
(`motivo_fora_do_lancamento`), e o arquivo do Superlógica sai com o resto —
diferente da aba 3, que trava tudo por uma pendência, porque aqui parte dos
documentos é manual **por regra do usuário**:

- **nota da F&F com retenção** (serviço ≠ líquido, **ou** contrib. sociais /
  previdência retidas) — regra do usuário: lançar à mão
- condomínio não identificado pelo CNPJ, ou sem ID SL
- sem valor no código
- boleto sem vencimento no código (fator `0000`/`9999`) — **não** se usa a
  data impressa na folha: sem a repetição do DARF, seria data solta
- DARF sem vencimento confirmado

**Trava de pasta misturada** (`conferir_lote_homogeneo`): documentos
reconhecidos de mais de um tipo (nota, boleto, arrecadação) → o arquivo do
Superlógica não é gerado. Documento não reconhecido não conta — o
"Detalhamento do Faturamento" das pastas da F&F travaria todo lote.

**Perguntas ao fim da extração**, não no início como na aba 3: chave sempre,
vencimento só se o lote tem nota fiscal. Lá o vencimento vai carimbado no
PDF durante o processamento; aqui nada é carimbado.

**`gerar_planilha_despesas` aceita dois formatos de lançamento**: a tupla
`(id_sl, valor)` da aba 3 e o dict por coluna da aba 2.

Validação (`tests/_validar_lancamento.py`): 284 boletos do sindicato e 12
DARF lançados, e o condomínio pelo CNPJ bateu com o nome do arquivo em todos.

**Pendente de teste real de importação:** se o importador aceita o código de
arrecadação de 48 dígitos em `linha_digitavel` (ele recusava o que não fosse
boleto), e em que formato espera a `competencia` (a da NFS-e é data cheia).
O Paybox anexar os boletos sozinho com a linha digitável é plausível e **não
testado**.

Spec: `docs/superpowers/specs/2026-09-27-lancamento-superlogica-extracao-design.md`.
````

- [ ] **Step 4: Bump the version**

Em `identificacao_por_cnpj_6_0.py:212`:

```python
        self.title("Codificador v6.21.0")
```

- [ ] **Step 5: Run the whole suite one last time**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **462 testes, OK**.

- [ ] **Step 6: Commit**

```bash
git add CLAUDE.md identificacao_por_cnpj_6_0.py tests/_validar_lancamento.py
git commit -m "Documenta o lancamento no Superlogica pela aba 2 e marca a versao 6.21.0"
```

---

## Notas para quem executa

- **Teste de importação real é do usuário, antes de liberar.** Nenhuma tarefa acima o substitui: gerar um arquivo com **um** DARF e **uma** nota, importar no Superlógica e conferir (1) se a coluna `linha_digitavel` aceita o código de arrecadação de 48 dígitos e (2) como a `competencia` da nota ficou gravada. Se o DARF for recusado, a correção é esvaziar `linha_digitavel` para `TIPO_LANCAMENTO_ARRECADACAO` em `montar_lancamentos`.
- **O modelo precisa ter as seis colunas** de `COLUNAS_POR_DOCUMENTO`. O modelo baixado do Superlógica tem as 32 — incluindo todas elas —, mas um modelo cortado à mão falha com mensagem dizendo qual coluna falta.
- O teste manual pela interface fica para o controlador: rodar a aba 2 com "Nenhum" (tem de sair idêntico a hoje) e com uma opção escolhida, numa pasta do sindicato.
