# Planilha de Despesas do Superlógica — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transformar o lote de protocolos dos Correios na planilha de importação de despesas do Superlógica — uma linha por protocolo cobrável, com `condomínio` preenchido pelo ID SL do cadastro e `valor` pelo valor apurado.

**Architecture:** Duas funções puras em `logica.py` (montar a lista de lançamentos a partir do resultado do painel; gerar o arquivo copiando o modelo do usuário) e um botão no painel de resultado da aba 3 que as liga. O modelo do Superlógica é aberto e preenchido, nunca reconstruído.

**Tech Stack:** Python 3, openpyxl 3.1.5, CustomTkinter, unittest.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-08-19-despesas-superlogica-design.md`.
- Branch de trabalho: `despesas-superlogica` (já criada, spec já commitado).
- Comentários, nomes de variáveis e mensagens em **português**.
- Nenhum texto visível ao usuário pode dizer "OCR" ou "DPI".
- Todo widget CustomTkinter usa `corner_radius=0`; cores só via `self.tema_atual[chave]`, nunca hex solto.
- Cobalto (`acento`) só no que **resolve** uma pendência e no cartão/rótulo "Pendentes". O botão novo é neutro: contornado, fundo transparente.
- Dinheiro em `Decimal`; `float` só na escrita de planilha.
- `logica.py` não pode ganhar nenhuma dependência de interface (nada de `tkinter`, `customtkinter`, `messagebox`).
- Fixtures de teste nunca usam `cadastro_condominios.xlsx` nem o `Despesas.xlsx` real — usar `tests/cadastro_teste.py` e modelos sintéticos.
- A suíte roda com `python -m unittest discover -s tests -p "test_*.py"` (hoje 164 testes).
- **Nunca executar `python identificacao_por_cnpj_6_0.py`** — abre a janela principal e trava o processo. Verificar com `python -m py_compile` e `python -c "import identificacao_por_cnpj_6_0"`; para exercitar a tela, instanciar `App()` e chamar `withdraw()`.

---

## File Structure

| Arquivo | Responsabilidade |
|---|---|
| `logica.py` (modificar) | `_normalizar_cabecalho`, `gerar_planilha_despesas`, `lancamentos_de_despesa`; docstring de `carregar_cadastro`. |
| `tests/cadastro_teste.py` (modificar) | Ganha o campo `id_sl`, com um condomínio sem ele de propósito. |
| `tests/test_despesas_superlogica.py` (criar) | Geração do arquivo e montagem dos lançamentos. |
| `identificacao_por_cnpj_6_0.py` (modificar) | Botão no painel, ação, e o `codigo` que falta no item resolvido à mão. |
| `CLAUDE.md` (modificar) | Seção da aba 3 + histórico. |

---

### Task 1: Gerar o arquivo a partir do modelo (`logica.py`)

**Files:**
- Modify: `logica.py` (funções novas no fim do arquivo, depois de `salvar_planilha_protocolo`)
- Test: `tests/test_despesas_superlogica.py`

**Interfaces:**
- Produces:
  - `_normalizar_cabecalho(texto) -> str`
  - `COLUNA_DESPESA_CONDOMINIO = "condominio"`, `COLUNA_DESPESA_VALOR = "valor"`, `LINHA_MOLDE_DESPESAS = 2`
  - `gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos) -> None` — `lancamentos` é uma lista de `(id_sl: str, valor: Decimal)` na ordem de saída

**Contexto que o implementador precisa:** o modelo real é o arquivo de importação
de despesas baixado do Superlógica. Linha 1 é o cabeçalho (32 colunas na versão
atual); linha 2 é uma linha de exemplo que o usuário preencheu com os campos que
se repetem (fornecedor, favorecido, categoria, tipo de documento, forma de
pagamento e uma chave numérica), deixando `condomínio` e `valor` em branco. Essa
linha 2 é o **molde** e é substituída pela primeira linha real — não pode sobrar
no arquivo final.

- [ ] **Step 1: Write the failing test**

Criar `tests/test_despesas_superlogica.py`:

```python
import os
import shutil
import sys
import tempfile
import unittest
from decimal import Decimal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app

#  Modelo sintético no formato do arquivo real do Superlógica: cabeçalho na
#  linha 1 e uma linha de exemplo na 2, com os campos que se repetem
#  preenchidos e condomínio/valor em branco. Nomes de coluna com acento e
#  caixa mista de propósito — é assim que vêm do Superlógica.
CABECALHO_MODELO = [
    "condomínio", "vencimento", "fornecedor", "conta_categoria",
    "valor", "forma_de_pagamento", "chave",
]
MOLDE_MODELO = [
    None, None, "DINAMICA SERVICOS POSTAIS", "2.4.6 Correios - Postagem Simples",
    None, "Trans. bancária", 46,
]


def montar_modelo(caminho, cabecalho=None, molde=None, linhas_extras=0):
    """Escreve um modelo sintético no disco e devolve o caminho."""
    wb = Workbook()
    sheet = wb.active
    sheet.append(cabecalho if cabecalho is not None else CABECALHO_MODELO)
    for celula in sheet[1]:
        celula.font = Font(bold=True)
    sheet.append(molde if molde is not None else MOLDE_MODELO)
    for _ in range(linhas_extras):
        sheet.append(molde if molde is not None else MOLDE_MODELO)
    wb.save(caminho)
    return caminho


class TestGerarPlanilhaDespesas(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.modelo = montar_modelo(os.path.join(self.pasta, "modelo.xlsx"))
        self.saida = os.path.join(self.pasta, "despesas.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def gerar(self, lancamentos):
        app.gerar_planilha_despesas(self.modelo, self.saida, lancamentos)
        return load_workbook(self.saida).active

    def test_uma_linha_por_lancamento_na_ordem_dada(self):
        sheet = self.gerar([("45", Decimal("7.70")),
                            ("496", Decimal("138.60")),
                            ("496", Decimal("119.35"))])
        #  Cabeçalho + 3 lançamentos, sem sobra
        self.assertEqual(sheet.max_row, 4)
        self.assertEqual([sheet.cell(row=l, column=1).value for l in (2, 3, 4)],
                         ["45", "496", "496"])
        self.assertEqual([sheet.cell(row=l, column=5).value for l in (2, 3, 4)],
                         [7.70, 138.60, 119.35])

    def test_campos_do_molde_repetem_em_todas_as_linhas(self):
        sheet = self.gerar([("45", Decimal("7.70")), ("48", Decimal("57.75"))])
        for linha in (2, 3):
            self.assertEqual(sheet.cell(row=linha, column=3).value,
                             "DINAMICA SERVICOS POSTAIS")
            self.assertEqual(sheet.cell(row=linha, column=4).value,
                             "2.4.6 Correios - Postagem Simples")
            self.assertEqual(sheet.cell(row=linha, column=6).value, "Trans. bancária")
            self.assertEqual(sheet.cell(row=linha, column=7).value, 46)

    def test_linha_molde_nao_sobra_em_branco(self):
        """A linha 2 do modelo é o molde e vira a primeira linha real — o
        arquivo final não pode ter nenhuma linha com condomínio/valor vazios."""
        sheet = self.gerar([("45", Decimal("7.70"))])
        self.assertEqual(sheet.max_row, 2)
        self.assertEqual(sheet.cell(row=2, column=1).value, "45")
        self.assertEqual(sheet.cell(row=2, column=5).value, 7.70)

    def test_linhas_extras_do_modelo_nao_sobram(self):
        """Modelo salvo com várias linhas de exemplo não pode deixar resto."""
        modelo = montar_modelo(os.path.join(self.pasta, "modelo3.xlsx"),
                               linhas_extras=4)
        app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        sheet = load_workbook(self.saida).active
        self.assertEqual(sheet.max_row, 2)

    def test_valor_gravado_como_numero_nao_texto(self):
        sheet = self.gerar([("45", Decimal("7.70"))])
        self.assertIsInstance(sheet.cell(row=2, column=5).value, float)

    def test_acha_as_colunas_por_nome_nao_por_posicao(self):
        """O layout é do Superlógica e pode mudar: as colunas são localizadas
        pelo nome normalizado, com acento e caixa ignorados."""
        modelo = montar_modelo(
            os.path.join(self.pasta, "outra_ordem.xlsx"),
            cabecalho=["VALOR", "fornecedor", "  Condomínio  "],
            molde=[None, "DINAMICA SERVICOS POSTAIS", None])
        app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        sheet = load_workbook(self.saida).active
        self.assertEqual(sheet.cell(row=2, column=1).value, 7.70)
        self.assertEqual(sheet.cell(row=2, column=3).value, "45")

    def test_modelo_sem_coluna_condominio_avisa(self):
        modelo = montar_modelo(os.path.join(self.pasta, "sem_cond.xlsx"),
                               cabecalho=["valor", "fornecedor"],
                               molde=[None, "DINAMICA"])
        with self.assertRaises(ValueError) as erro:
            app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        self.assertIn("condominio", str(erro.exception))

    def test_modelo_sem_coluna_valor_avisa(self):
        modelo = montar_modelo(os.path.join(self.pasta, "sem_valor.xlsx"),
                               cabecalho=["condomínio", "fornecedor"],
                               molde=[None, "DINAMICA"])
        with self.assertRaises(ValueError) as erro:
            app.gerar_planilha_despesas(modelo, self.saida, [("45", Decimal("7.70"))])
        self.assertIn("valor", str(erro.exception))

    def test_lista_vazia_nao_gera_arquivo_pela_metade(self):
        with self.assertRaises(ValueError):
            app.gerar_planilha_despesas(self.modelo, self.saida, [])
        self.assertFalse(os.path.isfile(self.saida))

    def test_estilo_do_molde_e_preservado_nas_linhas_novas(self):
        """Copiar em vez de reconstruir é o ponto do desenho: o importador do
        Superlógica pode depender de formato de célula."""
        wb = load_workbook(self.modelo)
        wb.active.cell(row=2, column=3).font = Font(bold=True, italic=True)
        wb.save(self.modelo)
        sheet = self.gerar([("45", Decimal("7.70")), ("48", Decimal("57.75"))])
        for linha in (2, 3):
            self.assertTrue(sheet.cell(row=linha, column=3).font.bold)
            self.assertTrue(sheet.cell(row=linha, column=3).font.italic)


class TestNormalizarCabecalho(unittest.TestCase):
    def test_tira_acento_caixa_e_espacos(self):
        self.assertEqual(app._normalizar_cabecalho("  Condomínio "), "condominio")
        self.assertEqual(app._normalizar_cabecalho("VALOR"), "valor")

    def test_celula_vazia_vira_string_vazia(self):
        self.assertEqual(app._normalizar_cabecalho(None), "")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_despesas_superlogica -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'gerar_planilha_despesas'`

- [ ] **Step 3: Write minimal implementation**

No fim de `logica.py`. O módulo já importa `copy`, `unicodedata` e
`load_workbook` — conferir antes de acrescentar imports.

```python
# ============================================================
#  PLANILHA DE DESPESAS DO SUPERLÓGICA
# ============================================================

#  Nomes das duas colunas que o programa preenche, já normalizados. O resto do
#  layout (32 colunas na versão atual) pertence ao Superlógica e é copiado do
#  modelo do usuário sem interpretação.
COLUNA_DESPESA_CONDOMINIO = "condominio"
COLUNA_DESPESA_VALOR = "valor"
LINHA_MOLDE_DESPESAS = 2


def _normalizar_cabecalho(texto):
    """Cabeçalho sem acento, minúsculo e sem espaços nas pontas. Serve para
    achar a coluna pelo NOME em vez da posição: o modelo é do Superlógica e
    pode ser reordenado ou reacentuado sem aviso."""
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return texto.strip().lower()


def gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos):
    """
    Gera a planilha de importação de despesas do Superlógica a partir do
    modelo do usuário. `lancamentos` é [(id_sl, valor), ...] na ordem de saída.

    O modelo é ABERTO E PREENCHIDO, nunca reconstruído: copiar preserva
    formatos de célula, validações e colunas ocultas que o importador do
    Superlógica pode exigir e que uma planilha montada do zero perderia sem
    aviso.

    A linha 2 do modelo é o molde — os campos que se repetem em todo
    lançamento (fornecedor, categoria, forma de pagamento...). Ela é
    SUBSTITUÍDA pela primeira linha real; nenhuma linha de exemplo pode
    sobrar no arquivo final.
    """
    if not lancamentos:
        raise ValueError(
            "Nenhum lançamento para gerar — a planilha não foi criada.")

    wb = load_workbook(caminho_modelo)
    sheet = wb.active

    colunas = {}
    for celula in sheet[1]:
        nome = _normalizar_cabecalho(celula.value)
        if nome:
            colunas.setdefault(nome, celula.column)

    faltando = [nome for nome in (COLUNA_DESPESA_CONDOMINIO, COLUNA_DESPESA_VALOR)
                if nome not in colunas]
    if faltando:
        raise ValueError(
            "O modelo não tem a(s) coluna(s): " + ", ".join(faltando) +
            ". Confira se o arquivo é o modelo de despesas do Superlógica.")

    coluna_condominio = colunas[COLUNA_DESPESA_CONDOMINIO]
    coluna_valor = colunas[COLUNA_DESPESA_VALOR]

    #  Valor e estilo de cada célula do molde, lidos ANTES de escrever — a
    #  primeira linha gerada sobrescreve a própria linha-molde.
    molde = []
    for coluna in range(1, sheet.max_column + 1):
        celula = sheet.cell(row=LINHA_MOLDE_DESPESAS, column=coluna)
        molde.append((celula.value, copy.copy(celula._style)))

    for indice, (id_sl, valor) in enumerate(lancamentos):
        numero_linha = LINHA_MOLDE_DESPESAS + indice
        for coluna, (valor_molde, estilo) in enumerate(molde, start=1):
            celula = sheet.cell(row=numero_linha, column=coluna)
            celula.value = valor_molde
            celula._style = copy.copy(estilo)
        sheet.cell(row=numero_linha, column=coluna_condominio).value = id_sl
        sheet.cell(row=numero_linha, column=coluna_valor).value = float(valor)

    #  Modelo salvo com mais de uma linha de exemplo não pode deixar resto
    #  depois do último lançamento — seria despesa fantasma na importação.
    primeira_sobra = LINHA_MOLDE_DESPESAS + len(lancamentos)
    if sheet.max_row >= primeira_sobra:
        sheet.delete_rows(primeira_sobra, sheet.max_row - primeira_sobra + 1)

    wb.save(caminho_saida)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_despesas_superlogica -v`
Expected: PASS (12 testes)

- [ ] **Step 5: Run the full suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (176 testes)

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_despesas_superlogica.py
git commit -m "Gera a planilha de despesas do Superlógica a partir do modelo"
```

---

### Task 2: Montar os lançamentos a partir do lote (`logica.py`)

**Files:**
- Modify: `logica.py` (depois de `gerar_planilha_despesas`), docstring de `carregar_cadastro`
- Modify: `tests/cadastro_teste.py`
- Test: `tests/test_despesas_superlogica.py`

**Interfaces:**
- Consumes: `_codigos_do_cadastro(cadastro)` (já existe; devolve dict código→cnpj).
- Produces: `lancamentos_de_despesa(resultado, cadastro) -> tuple[list, dict]` — devolve `(lancamentos, travas)`, onde `lancamentos` é `[(id_sl, valor), ...]` e `travas` é `{"sem_valor": [arquivo, ...], "sem_id_sl": [arquivo, ...]}`

**Contexto que o implementador precisa:** `resultado` é o dict que o painel da
aba 3 renderiza. Tem três listas: `processados` (protocolos com valor apurado,
cada item com `arquivo`, `codigo`, `condominio`, `valor` em `Decimal`),
`pendentes` (a contagem não fechou e ninguém informou o valor) e `ignorados`
(arquivos que nem são protocolo).

**Por que receber `resultado` e não as linhas da planilha:** nas linhas prontas,
um arquivo que não é protocolo e uma pendência do tipo "não foi possível ler o
documento" ficam idênticos — código, condomínio e valor todos vazios. Mas um
deve **travar** a geração e o outro deve ser **ignorado**. Só o `resultado`
separa os dois. (O spec descreve a assinatura antiga, `lancamentos_de_despesa(linhas, cadastro)`;
está corrigido aqui.)

- [ ] **Step 1: Dar `id_sl` ao cadastro de teste**

Em `tests/cadastro_teste.py`, acrescentar o campo a cada registro, e ajustar a
docstring (que hoje diz `Valor = {"codigo", "nome"}`). **LAGO MAGGIORE fica com
`id_sl` vazio de propósito**, para cobrir o condomínio cadastrado sem ID:

```python
"""
Cadastro fixo usado só pelos testes — NUNCA a planilha real
(cadastro_condominios.xlsx). Congelado de propósito: se os testes lessem a
planilha real, cadastrar um condomínio novo poderia quebrar um teste sem que
nada esteja errado. Os 9 condomínios abaixo são os que apareceram nos bugs
reais corrigidos nas versões 6.1/6.2.

Chave = CNPJ normalizado (14 dígitos). Valor = {"codigo", "nome", "id_sl"}.

"id_sl" é o código do condomínio no Superlógica, usado para gerar a planilha
de despesas. LAGO MAGGIORE está sem de propósito: na planilha real 10 dos 764
condomínios não têm esse campo, e o programa precisa lidar com isso.
"""

CADASTRO_TESTE = {
    "40338774000141": {"codigo": "10695", "nome": "ARAUJO LIMA", "id_sl": "701"},
    "01195716000154": {"codigo": "10004", "nome": "KLOSTERS", "id_sl": "44"},
    "08578541000103": {"codigo": "10002", "nome": "SAN REMO", "id_sl": "42"},
    "05695194000100": {"codigo": "10490", "nome": "CENTRO COM CONDE DE BONFIM RES", "id_sl": "487"},
    "29361458000158": {"codigo": "10491", "nome": "CENTRO COM CONDE DE BONFIM", "id_sl": "488"},
    "07448975000126": {"codigo": "11194", "nome": "MANHATTAN", "id_sl": "912"},
    "10864886000175": {"codigo": "10625", "nome": "ARGENTINA", "id_sl": "633"},
    "29273778000156": {"codigo": "11189", "nome": "VILLE DE BEAUVAIS", "id_sl": "907"},
    "07945453000130": {"codigo": "10590", "nome": "LAGO MAGGIORE", "id_sl": ""},
}
```

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (176 testes) — nenhum teste existente compara o dict inteiro, então acrescentar chave não quebra nada. Se quebrar, o teste que quebrou precisa ser ajustado e o motivo relatado.

- [ ] **Step 2: Write the failing test**

Acrescentar a `tests/test_despesas_superlogica.py`, antes do bloco `if __name__`:

```python
from cadastro_teste import CADASTRO_TESTE

KLOSTERS = {"arquivo": "a.pdf", "codigo": "10004", "condominio": "10004 KLOSTERS",
            "valor": Decimal("57.75")}
SAN_REMO = {"arquivo": "b.pdf", "codigo": "10002", "condominio": "10002 SAN REMO",
            "valor": Decimal("11.55")}
#  Cadastrado, mas sem ID no Superlógica — 10 dos 764 reais estão assim.
LAGO = {"arquivo": "c.pdf", "codigo": "10590", "condominio": "10590 LAGO MAGGIORE",
        "valor": Decimal("7.70")}
#  Resolvido à mão depois de "não foi possível ler": não tem código nenhum.
SEM_CODIGO = {"arquivo": "d.pdf", "codigo": None, "condominio": "",
              "valor": Decimal("300.30")}


def resultado_com(processados=(), pendentes=(), ignorados=()):
    return {"processados": list(processados), "pendentes": list(pendentes),
            "ignorados": list(ignorados)}


class TestLancamentosDeDespesa(unittest.TestCase):
    def test_um_lancamento_por_protocolo_na_ordem(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, SAN_REMO]), CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75")),
                                       ("42", Decimal("11.55"))])
        self.assertEqual(travas, {"sem_valor": [], "sem_id_sl": []})

    def test_mesmo_condominio_duas_vezes_gera_dois_lancamentos(self):
        """Cada protocolo físico vira um lançamento — somar quebraria a
        correspondência com o papel que originou cada um."""
        outro = dict(KLOSTERS, arquivo="a2.pdf", valor=Decimal("138.60"))
        lancamentos, _ = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, outro]), CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75")),
                                       ("44", Decimal("138.60"))])

    def test_pendente_sem_valor_trava(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS], pendentes=[{"arquivo": "p.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(travas["sem_valor"], ["p.pdf"])
        self.assertEqual(travas["sem_id_sl"], [])

    def test_condominio_sem_id_sl_trava(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS, LAGO]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["c.pdf"])
        self.assertEqual(travas["sem_valor"], [])

    def test_protocolo_sem_codigo_trava_por_falta_de_id(self):
        _, travas = app.lancamentos_de_despesa(
            resultado_com([SEM_CODIGO]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["d.pdf"])

    def test_codigo_fora_do_cadastro_trava_por_falta_de_id(self):
        fora = dict(KLOSTERS, arquivo="e.pdf", codigo="99999")
        _, travas = app.lancamentos_de_despesa(
            resultado_com([fora]), CADASTRO_TESTE)
        self.assertEqual(travas["sem_id_sl"], ["e.pdf"])

    def test_arquivo_que_nao_e_protocolo_nao_trava_nem_vira_lancamento(self):
        """Ignorados nunca deveriam virar despesa — não podem travar o lote."""
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS], ignorados=[{"arquivo": "outro.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(lancamentos, [("44", Decimal("57.75"))])
        self.assertEqual(travas, {"sem_valor": [], "sem_id_sl": []})

    def test_lote_inteiro_travado_devolve_lista_vazia(self):
        lancamentos, travas = app.lancamentos_de_despesa(
            resultado_com([LAGO], pendentes=[{"arquivo": "p.pdf"}]),
            CADASTRO_TESTE)
        self.assertEqual(lancamentos, [])
        self.assertEqual(travas["sem_id_sl"], ["c.pdf"])
        self.assertEqual(travas["sem_valor"], ["p.pdf"])

    def test_valor_continua_decimal(self):
        """Dinheiro só vira float na escrita da planilha."""
        lancamentos, _ = app.lancamentos_de_despesa(
            resultado_com([KLOSTERS]), CADASTRO_TESTE)
        self.assertIsInstance(lancamentos[0][1], Decimal)
```

- [ ] **Step 3: Run test to verify it fails**

Run: `python -m unittest tests.test_despesas_superlogica.TestLancamentosDeDespesa -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'lancamentos_de_despesa'`

- [ ] **Step 4: Write minimal implementation**

Em `logica.py`, depois de `gerar_planilha_despesas`:

```python
def lancamentos_de_despesa(resultado, cadastro):
    """
    Monta os lançamentos de despesa a partir do resultado de um lote de
    protocolos. Devolve `(lancamentos, travas)`:

    - `lancamentos`: [(id_sl, valor)] na ordem do painel, pronto para
      `gerar_planilha_despesas`;
    - `travas`: {"sem_valor": [...], "sem_id_sl": [...]} com os nomes dos
      arquivos que impedem a geração — a interface não gera nada enquanto
      houver qualquer um, e mostra os dois grupos separados porque a ação é
      diferente (pendência resolve no painel; ID SL resolve no cadastro).

    Recebe o `resultado` do painel, e não as linhas da planilha, porque nas
    linhas um arquivo que não é protocolo e uma pendência "não foi possível
    ler o documento" ficam idênticos (código, condomínio e valor vazios) — e
    um deve travar enquanto o outro deve ser ignorado.

    Mesmo condomínio em dois protocolos gera dois lançamentos: cada um
    continua rastreável até o papel que o originou.
    """
    codigos = _codigos_do_cadastro(cadastro)
    lancamentos = []
    travas = {"sem_valor": [], "sem_id_sl": []}

    #  Ignorados ficam de fora sem travar: não são protocolo, nunca deveriam
    #  virar despesa.
    for item in resultado.get("pendentes", []) or []:
        travas["sem_valor"].append(item.get("arquivo", ""))

    for item in resultado.get("processados", []) or []:
        codigo = item.get("codigo") or ""
        registro = cadastro.get(codigos.get(codigo, "")) if codigo else None
        id_sl = (registro or {}).get("id_sl", "")
        if not id_sl:
            travas["sem_id_sl"].append(item.get("arquivo", ""))
            continue
        lancamentos.append((id_sl, item["valor"]))

    return lancamentos, travas
```

- [ ] **Step 5: Corrigir a docstring de `carregar_cadastro`**

Ela afirma hoje que o ID SL é "Guardado só como referência: nada da
identificação nem dos carimbos usa esse campo." Deixou de ser verdade — a
geração da planilha de despesas usa. Trocar a frase por algo como:
"Nada da identificação nem dos carimbos usa esse campo; quem usa é a geração
da planilha de despesas do Superlógica (`lancamentos_de_despesa`)."

- [ ] **Step 6: Run the full suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (185 testes)

- [ ] **Step 7: Commit**

```bash
git add logica.py tests/cadastro_teste.py tests/test_despesas_superlogica.py
git commit -m "Monta os lançamentos de despesa a partir do lote de protocolos"
```

---

### Task 3: Botão no painel de resultado (`identificacao_por_cnpj_6_0.py`)

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` — o `from logica import (...)` (linha ~57), `_acao_informar_valor` (o dict `registro_processado`), o rodapé de `mostrar_resultado_protocolos` (linha ~3121), e um método novo
- Test: verificação por script com `App()` + `withdraw()`

**Interfaces:**
- Consumes: `lancamentos_de_despesa(resultado, cadastro)` e `gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos)` (Tasks 1 e 2); `self._resultado_protocolos`, `self._ctx_protocolos` (já existem).
- Produces: `_acao_gerar_despesas(self)`.

- [ ] **Step 1: Fazer o item resolvido à mão carregar o código**

Em `_acao_informar_valor`, o dict `registro_processado` monta o item que entra
em `resultado["processados"]` — e **não inclui `codigo`**. Sem ele,
`lancamentos_de_despesa` não acha o ID SL, e justamente os protocolos que
precisaram de atenção manual travariam o lote inteiro.

Acrescentar a chave, ao lado de `"condominio"`:

```python
            #  O código é o que liga este protocolo ao ID SL do cadastro na
            #  hora de gerar a planilha de despesas — sem ele, um protocolo
            #  resolvido à mão travaria a geração por "sem ID SL".
            "codigo": dados.get("codigo"),
```

Verificar com um script que resolva uma pendência à mão (como já é feito nos
testes manuais da aba 3) e confirme que o item em `resultado["processados"]`
tem `codigo` preenchido.

- [ ] **Step 2: Acrescentar o import**

No `from logica import (...)`, acrescentar:

```python
    lancamentos_de_despesa, gerar_planilha_despesas,
```

- [ ] **Step 3: Escrever a ação**

Inserir logo depois de `_acao_abrir_planilha_protocolos`:

```python
    def _acao_gerar_despesas(self):
        """
        Gera a planilha de importação de despesas do Superlógica a partir
        deste lote: uma linha por protocolo cobrável, com o condomínio
        preenchido pelo ID SL do cadastro.

        Recusa gerar enquanto houver protocolo cobrável incompleto — e diz
        quais são e onde resolver cada caso, porque "não gera nada" sem a
        lista vira adivinhação.
        """
        resultado = getattr(self, "_resultado_protocolos", None)
        if not resultado:
            return

        lancamentos, travas = lancamentos_de_despesa(resultado, self.cadastro)

        if travas["sem_valor"] or travas["sem_id_sl"]:
            partes = []
            if travas["sem_valor"]:
                partes.append(
                    "Sem valor — resolva no botão \"Informar valor\":\n  "
                    + "\n  ".join(travas["sem_valor"]))
            if travas["sem_id_sl"]:
                partes.append(
                    "Sem o código do Superlógica — preencha o \"ID SL\" na aba "
                    "de Cadastro:\n  " + "\n  ".join(travas["sem_id_sl"]))
            messagebox.showwarning(
                "Ainda não dá para gerar",
                "A planilha de despesas não foi gerada porque estes protocolos "
                "estão incompletos:\n\n" + "\n\n".join(partes),
                parent=self._janela_resultado)
            return

        if not lancamentos:
            messagebox.showinfo(
                "Nada para lançar",
                "Nenhum protocolo deste lote vira despesa.",
                parent=self._janela_resultado)
            return

        modelo = filedialog.askopenfilename(
            title="Escolha o modelo de despesas do Superlógica",
            filetypes=[("Excel", "*.xlsx")],
            parent=self._janela_resultado)
        if not modelo:
            return

        ctx = self._ctx_protocolos or {}
        pasta_sugerida = os.path.dirname(ctx.get("destino", "")) or None
        destino = filedialog.asksaveasfilename(
            title="Salvar planilha de despesas como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="despesas_superlogica.xlsx",
            initialdir=pasta_sugerida,
            parent=self._janela_resultado)
        if not destino:
            return

        try:
            gerar_planilha_despesas(modelo, destino, lancamentos)
        except Exception as e:
            messagebox.showerror(
                "Erro ao gerar",
                f"Não foi possível gerar a planilha de despesas:\n{e}",
                parent=self._janela_resultado)
            return

        plural = "lançamento" if len(lancamentos) == 1 else "lançamentos"
        if messagebox.askyesno(
            "Planilha de despesas pronta",
            f"{len(lancamentos)} {plural} gravado(s) em:\n{destino}\n\nAbrir agora?",
            parent=self._janela_resultado,
        ):
            try:
                os.startfile(destino)
            except Exception as e:
                messagebox.showerror(
                    "Erro", f"Não foi possível abrir a planilha:\n{e}",
                    parent=self._janela_resultado)
```

- [ ] **Step 4: Acrescentar o botão ao rodapé do painel**

Em `mostrar_resultado_protocolos`, o rodapé já tem o botão "Abrir planilha"
empacotado com `.pack(side="right")`. Acrescentar o novo **antes** dele no
código, para que apareça à esquerda dele, com o mesmo estilo neutro:

```python
        ctk.CTkButton(
            rodape, text="Gerar planilha do Superlógica", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=self._acao_gerar_despesas,
        ).pack(side="right", padx=(0, 8))
```

- [ ] **Step 5: Verificar**

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem saída

Run: `python -c "import identificacao_por_cnpj_6_0"`
Expected: sem erro

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (185 testes)

Escrever um script à parte que instancie `App()` com `withdraw()`, monte um
`resultado` sintético com processados, pendentes e ignorados, chame
`mostrar_resultado_protocolos` e depois `_acao_gerar_despesas` com os diálogos
de arquivo e o `messagebox` substituídos por dublês — confirmando que: com
pendente, recusa e a mensagem lista o arquivo certo no grupo certo; sem
pendente e com tudo cadastrado, gera o arquivo. **Nunca executar
`python identificacao_por_cnpj_6_0.py`.**

- [ ] **Step 6: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Botão que gera a planilha de despesas do Superlógica no painel"
```

---

### Task 4: Verificação com os arquivos reais e documentação

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Gerar contra o modelo real**

O modelo real está em `C:\Users\Dell\Downloads\Despesas.xlsx` (32 colunas,
`condomínio` na 1 e `valor` na 10; linha 2 com fornecedor, favorecido,
categoria, tipo de documento, forma de pagamento e `chave` = 46).

Escrever um script à parte (fora do repositório) que:
1. leia o cadastro real com `carregar_cadastro("cadastro_condominios.xlsx")`;
2. monte um `resultado` com os protocolos do lote multipágina — 11161 MARILIA
   duas vezes (R$ 138,60 e R$ 119,35) e 10380 ROXY (R$ 231,00);
3. chame `lancamentos_de_despesa` e depois `gerar_planilha_despesas`, gravando
   num arquivo temporário;
4. imprima, do arquivo gerado: o número de linhas, e para cada linha o
   condomínio, o valor e três dos campos repetidos.

Expected: 4 linhas no total (cabeçalho + 3 lançamentos); condomínios `496`,
`496`, `649`; valores 138.6, 119.35, 231.0; fornecedor, categoria e forma de
pagamento idênticos nas três linhas; nenhuma linha em branco.

Se algum número divergir, **investigar** — não ajustar o esperado.

- [ ] **Step 2: Verificação visual pelo usuário**

Não automatizável, e não cabe a um subagente. Pedir ao usuário que processe um
lote de protocolos, clique em "Gerar planilha do Superlógica", aponte o modelo
e confira o arquivo gerado — e, o teste que só ele pode fazer, **importe no
Superlógica** e veja se os lançamentos entram corretamente.

- [ ] **Step 3: Atualizar o CLAUDE.md**

Na seção "Contagem dos Protocolos dos Correios (aba 3)", acrescentar um bloco
sobre a geração da planilha de despesas, registrando:

- o que o botão faz e de onde vem cada campo (condomínio = ID SL do cadastro,
  valor = valor apurado, resto copiado da linha-molde do modelo);
- que o modelo é lido do arquivo do usuário e **copiado, não reconstruído**,
  porque o layout pertence ao Superlógica e copiar preserva formatos e
  validações que o importador pode exigir;
- que as colunas são achadas pelo nome normalizado, não pela posição;
- que mesmo condomínio em dois protocolos gera duas linhas, de propósito;
- que a geração **trava por completo** enquanto houver pendência sem valor ou
  condomínio sem ID SL, e que 10 dos 764 condomínios estão sem o campo — então
  isso vai acontecer, e o conserto é preencher na aba de Cadastro.

Atualizar também a menção ao "ID SL" na seção "O que o programa faz", que hoje
afirma que o campo é guardado "**só como referência**: nada da identificação nem
dos carimbos usa esse campo" — continua verdade para identificação e carimbos,
mas agora o campo tem um uso real.

Acrescentar a entrada da **v6.14.0** ao "Histórico de decisões".

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md
git commit -m "Documenta a geração da planilha de despesas do Superlógica"
```
