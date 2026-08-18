# Contagem e cobrança dos Protocolos dos Correios — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Aba nova que conta as unidades de cada Protocolo de Recebimento dos Correios, multiplica pela tarifa informada no lote, carimba o valor no topo direito do PDF (mantendo o carimbo lateral do código) e gera uma planilha com o total.

**Architecture:** Toda a lógica nova vive em `logica.py` (sem dependência de interface, como o resto do módulo) e é coberta por testes `unittest`. A interface ganha a aba 3 e uma thread de processamento no molde da aba 2. `criar_overlay` e `processar_pdf` ganham parâmetros opcionais com default igual ao comportamento atual — o mesmo movimento da v6.9.0 quando `angulo=0` foi adicionado.

**Tech Stack:** Python 3, CustomTkinter, pypdf, reportlab, openpyxl, pymupdf, winocr (Windows), unittest.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-08-17-contagem-protocolo-correio-design.md`.
- Branch de trabalho: `protocolo-correio-valor` (já criada, spec já commitado).
- Comentários, nomes de variáveis e mensagens em **português**.
- Nenhum texto visível ao usuário pode dizer "OCR" ou "DPI" — usar "protocolos escaneados", "qualidade de leitura".
- Todo widget CustomTkinter usa `corner_radius=0`; cores só via `self.tema_atual[chave]`, nunca hex solto.
- Cobalto (`acento`) só no botão primário da aba.
- Dinheiro em `Decimal` internamente; `float` só na escrita da planilha.
- Fixtures de teste **nunca** contêm nomes de moradores reais nem usam `cadastro_condominios.xlsx` — usar `tests/cadastro_teste.py`.
- Rodar a suíte inteira com: `python -m unittest discover -s tests -p "test_*.py"`
- Comportamento atual de `criar_overlay`, `processar_pdf` e `extrair_texto_ocr` preservado byte a byte quando os parâmetros novos não são passados.

---

## File Structure

| Arquivo | Responsabilidade |
|---|---|
| `logica.py` (modificar) | Contagem, conferência, cálculo, textos dos carimbos, planilha. Bloco novo logo abaixo do bloco "PROTOCOLO DE RECEBIMENTO DE DOCUMENTO" que já existe (linha ~528). |
| `identificacao_por_cnpj_6_0.py` (modificar) | Aba 3, popup da tarifa, thread de processamento. |
| `tests/test_protocolo_contagem.py` (criar) | Contagem, conferência, valor, textos. |
| `tests/test_planilha_protocolo.py` (criar) | Linha e gravação da planilha. |
| `tests/test_carimbo_lateral.py` (modificar) | Regressão de `criar_overlay` e cobertura do alinhamento à direita. |
| `CLAUDE.md` (modificar) | Seção da funcionalidade + histórico de versão. |

---

### Task 1: Contagem e conferência (`logica.py`)

**Files:**
- Modify: `logica.py` (inserir após a linha 527, fim de `montar_texto_protocolo_correio`)
- Test: `tests/test_protocolo_contagem.py`

**Interfaces:**
- Consumes: `extrair_codigo_protocolo_correio(texto)`, `MARCADOR_PROTOCOLO_CORREIO` (já existem em `logica.py`).
- Produces:
  - `extrair_dados_protocolo_correio(texto) -> dict | None` com as chaves `codigo` (str|None), `condominio` (str), `total_impresso` (int|None), `linhas_contadas` (int), `entregas_contadas` (int)
  - `conferir_contagem_protocolo(dados) -> tuple[bool, str]`

- [ ] **Step 1: Write the failing test**

Criar `tests/test_protocolo_contagem.py`. As constantes reproduzem os defeitos reais do `winocr` (página inteira numa linha só, colunas fora de ordem, traço duplicado) com nomes fictícios:

```python
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

# O winocr devolve a página inteira numa linha só e com as colunas fora de
# ordem — os "Correio" da coluna Entrega saem todos no fim, depois do rodapé.
# Nomes fictícios de propósito: fixture não guarda nome de morador real.
PROTOCOLO_OK = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Imodata - Condominios e Imoveis "
    "Rua Barata Ribeiro, 774 / 100 andar Copacabana } RJ -22.051-002 "
    "matriz@imodata.net - (21) 3816-7800 Assinatura Entrega "
    "Correio Correio Correio 1155 1 de 1"
)

# Caso real do protocolo de 15 unidades: o OCR duplicou o traço numa linha
# ("702 - - Enny"), então a regex de unidade conta 2 e o "Correio" conta 3.
PROTOCOLO_TRACO_DUPLO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega Correio Correio Correio 1 de 1"
)

PROTOCOLO_SEM_LISTANDO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva Assinatura Entrega "
    "Correio Correio 1 de 1"
)

PROTOCOLO_DIVERGENTE = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal Listando 3 unidades Assinatura Entrega Correio 1 de 1"
)

NAO_E_PROTOCOLO = "CO-ESTIPULANTE: KLOSTERS CNPJ: 01.195.716/0001-54"


class TestExtracaoProtocolo(unittest.TestCase):
    def test_documento_que_nao_e_protocolo(self):
        self.assertIsNone(app.extrair_dados_protocolo_correio(NAO_E_PROTOCOLO))

    def test_texto_vazio_nao_quebra(self):
        self.assertIsNone(app.extrair_dados_protocolo_correio(""))

    def test_le_codigo_e_condominio_do_cabecalho(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["codigo"], "10004")
        self.assertEqual(dados["condominio"], "W700A KLOSTERS")

    def test_conta_as_tres_fontes(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["linhas_contadas"], 3)
        self.assertEqual(dados["entregas_contadas"], 3)

    def test_rodape_nao_vira_unidade(self):
        """CEP, telefone e o manuscrito lido pelo OCR ficam depois do
        "Listando" e não podem entrar na contagem de linhas."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        self.assertEqual(dados["linhas_contadas"], 3)

    def test_traco_duplicado_derruba_a_regex_mas_nao_a_contagem(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_TRACO_DUPLO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["entregas_contadas"], 3)

    def test_sem_listando_o_total_e_none(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_SEM_LISTANDO)
        self.assertIsNone(dados["total_impresso"])


class TestConferenciaProtocolo(unittest.TestCase):
    def test_aceita_quando_tudo_bate(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_OK)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)
        self.assertEqual(motivo, "")

    def test_aceita_quando_so_um_conferidor_bate(self):
        """Traço duplicado: a regex de unidade erra, o "Correio" salva."""
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_TRACO_DUPLO)
        aceito, _ = app.conferir_contagem_protocolo(dados)
        self.assertTrue(aceito)

    def test_recusa_sem_total_impresso(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_SEM_LISTANDO)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        self.assertIn("total impresso", motivo)

    def test_recusa_quando_nenhum_conferidor_bate(self):
        dados = app.extrair_dados_protocolo_correio(PROTOCOLO_DIVERGENTE)
        aceito, motivo = app.conferir_contagem_protocolo(dados)
        self.assertFalse(aceito)
        self.assertIn("3", motivo)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_protocolo_contagem -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'extrair_dados_protocolo_correio'`

- [ ] **Step 3: Write minimal implementation**

Em `logica.py`, logo após `montar_texto_protocolo_correio` (linha 527):

```python
RE_LISTANDO_PROTOCOLO = re.compile(r"Listando\s+(\d+)\s+unidade", re.IGNORECASE)
#  O winocr devolve a página numa linha só, então nada de (?m)^ aqui. O
#  (?:\s*[-–—])+ cobre o traço duplicado que o OCR produz às vezes
#  ("702 - - Enny Marins de Lima", visto no protocolo real do ASTORIA).
RE_UNIDADE_PROTOCOLO = re.compile(r"\b\d{1,4}(?:\s*[-–—])+\s*[A-Za-zÀ-ÿ]")
RE_ENTREGA_PROTOCOLO = re.compile(r"\bCorreio\b", re.IGNORECASE)
RE_CABECALHO_PROTOCOLO = re.compile(
    r"([^()\n]{0,60}?)\s*\(\d+\)\s*" + re.escape(MARCADOR_PROTOCOLO_CORREIO),
    re.IGNORECASE,
)


def extrair_dados_protocolo_correio(texto):
    """
    Lê um Protocolo de Recebimento de Documento e devolve o que é preciso
    para cobrar por ele. `None` se o documento não for um protocolo.

    Três contagens independentes porque cada uma falha de um jeito: o
    "Listando N unidades" impresso é a fonte do valor, e as outras duas
    servem para confirmá-lo (ver conferir_contagem_protocolo).
    """
    texto = texto or ""
    if MARCADOR_PROTOCOLO_CORREIO.lower() not in texto.lower():
        return None

    cabecalho = RE_CABECALHO_PROTOCOLO.search(texto)
    listando = RE_LISTANDO_PROTOCOLO.search(texto)

    #  A contagem de linhas olha só o que vem ANTES do "Listando": depois
    #  dele só há rodapé (CEP, telefone) e números que o OCR inventa lendo
    #  o valor manuscrito — nada disso é unidade.
    corpo = texto[:listando.start()] if listando else texto

    return {
        "codigo": extrair_codigo_protocolo_correio(texto),
        "condominio": cabecalho.group(1).strip() if cabecalho else "",
        "total_impresso": int(listando.group(1)) if listando else None,
        "linhas_contadas": len(RE_UNIDADE_PROTOCOLO.findall(corpo)),
        "entregas_contadas": len(RE_ENTREGA_PROTOCOLO.findall(texto)),
    }


def conferir_contagem_protocolo(dados):
    """
    Decide se dá para confiar na contagem. Devolve (aceito, motivo).

    O total impresso manda; basta que UM dos dois conferidores concorde com
    ele. Sem o total impresso não se aceita nada, mesmo que os conferidores
    concordem entre si — contar linhas por OCR sozinho é chute com cara de
    precisão, e o resultado aqui vira dinheiro cobrado.
    """
    total = dados.get("total_impresso")
    if total is None:
        return False, 'Não foi possível ler o total impresso ("Listando N unidades")'

    linhas = dados.get("linhas_contadas", 0)
    entregas = dados.get("entregas_contadas", 0)
    if total == linhas or total == entregas:
        return True, ""
    return False, f"Listando {total}, mas foram contadas {linhas} e {entregas} unidades"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_protocolo_contagem -v`
Expected: PASS (11 testes)

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_contagem.py
git commit -m "Conta e confere as unidades do protocolo dos Correios"
```

---

### Task 2: Cálculo do valor e textos do carimbo (`logica.py`)

**Files:**
- Modify: `logica.py` (após `conferir_contagem_protocolo`)
- Test: `tests/test_protocolo_contagem.py` (acrescentar classe)

**Interfaces:**
- Consumes: nada da Task 1 em runtime.
- Produces:
  - `valor_protocolo(unidades: int, tarifa) -> Decimal`
  - `formatar_reais(valor) -> str`
  - `montar_texto_valor_protocolo(unidades, tarifa, valor) -> str`

- [ ] **Step 1: Write the failing test**

Acrescentar ao fim de `tests/test_protocolo_contagem.py`, antes do `if __name__`:

```python
from decimal import Decimal


class TestValorProtocolo(unittest.TestCase):
    def test_multiplicacao_simples(self):
        self.assertEqual(app.valor_protocolo(15, Decimal("3.85")), Decimal("57.75"))

    def test_aceita_tarifa_float_sem_perder_centavo(self):
        self.assertEqual(app.valor_protocolo(12, 3.85), Decimal("46.20"))

    def test_lote_de_referencia_fecha(self):
        """Os quatro protocolos reais de 24/07/2026, conferidos contra os
        valores escritos à mão no papel."""
        total = sum(app.valor_protocolo(u, Decimal("3.85")) for u in (2, 15, 3, 12))
        self.assertEqual(total, Decimal("123.20"))

    def test_arredonda_meio_centavo_para_cima(self):
        self.assertEqual(app.valor_protocolo(1, Decimal("3.855")), Decimal("3.86"))

    def test_zero_unidades(self):
        self.assertEqual(app.valor_protocolo(0, Decimal("3.85")), Decimal("0.00"))

    def test_formata_em_reais(self):
        self.assertEqual(app.formatar_reais(Decimal("57.75")), "R$ 57,75")
        self.assertEqual(app.formatar_reais(Decimal("1234.50")), "R$ 1.234,50")

    def test_texto_do_carimbo_mostra_a_conta(self):
        texto = app.montar_texto_valor_protocolo(15, Decimal("3.85"), Decimal("57.75"))
        self.assertEqual(texto, "15 un × R$ 3,85 = R$ 57,75")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_protocolo_contagem.TestValorProtocolo -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'valor_protocolo'`

- [ ] **Step 3: Write minimal implementation**

Garantir no topo de `logica.py`, junto dos outros imports:

```python
from decimal import Decimal, ROUND_HALF_UP
```

E após `conferir_contagem_protocolo`:

```python
def valor_protocolo(unidades, tarifa):
    """
    unidades × tarifa em Decimal, duas casas. Dinheiro não passa por float:
    a tarifa vira Decimal a partir da string para não herdar o erro de
    representação binária (3.85 float não é exatamente 3,85).
    """
    tarifa_decimal = tarifa if isinstance(tarifa, Decimal) else Decimal(str(tarifa))
    bruto = Decimal(int(unidades)) * tarifa_decimal
    return bruto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def formatar_reais(valor):
    """1234.5 -> "R$ 1.234,50" (formato brasileiro)."""
    texto = f"{Decimal(str(valor)):,.2f}"
    return "R$ " + texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def montar_texto_valor_protocolo(unidades, tarifa, valor):
    """Linha carimbada no topo direito: "15 un × R$ 3,85 = R$ 57,75".
    Mostra a conta, não só o resultado, para conferir no papel sem
    precisar refazer a multiplicação."""
    return f"{unidades} un × {formatar_reais(tarifa)} = {formatar_reais(valor)}"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_protocolo_contagem -v`
Expected: PASS (18 testes)

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_contagem.py
git commit -m "Calcula o valor do protocolo em Decimal e monta o texto do carimbo"
```

---

### Task 3: Dois carimbos na mesma página (`logica.py`)

**Files:**
- Modify: `logica.py:731-784` (`criar_overlay` e `processar_pdf`)
- Test: `tests/test_carimbo_lateral.py`

**Interfaces:**
- Consumes: nada das tasks anteriores.
- Produces:
  - `criar_overlay(largura, altura, texto, fonte, tamanho, cor, x, y, centralizado, angulo=0, alinhamento="esquerda")`
  - `processar_pdf(caminho_entrada, caminho_saida, texto, config, carimbos_extras=None)` — cada item de `carimbos_extras` é um dict com as chaves `texto`, `fonte`, `tamanho`, `cor`, `x`, `y`, `centralizado`, `angulo`, `alinhamento`.

- [ ] **Step 1: Write the failing test**

Acrescentar a `tests/test_carimbo_lateral.py` (o arquivo já importa `logica as app` e usa `app.criar_overlay`; manter o estilo existente):

```python
import re
from pypdf import PdfReader
from reportlab.pdfbase.pdfmetrics import stringWidth

TEXTO_VALOR = "15 un × R$ 3,85 = R$ 57,75"


def posicao_do_texto(buffer):
    """Lê a coordenada em que o texto foi realmente desenhado, direto do
    content stream do PDF (reportlab grava "1 0 0 1 <x> <y> Tm" sem
    compressão). Comparar bytes do PDF não serviria: o cabeçalho é igual
    qualquer que seja o alinhamento."""
    dados = PdfReader(buffer).pages[0].get_contents().get_data()
    m = re.search(rb"1 0 0 1 ([-\d.]+) ([-\d.]+) Tm", dados)
    assert m, f"nenhum operador Tm encontrado em {dados!r}"
    return float(m.group(1)), float(m.group(2))


class TestAlinhamentoDireita(unittest.TestCase):
    def test_alinhado_a_esquerda_comeca_em_x(self):
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 567, 814,
            False, 0, "esquerda")
        x, y = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, 567, places=1)
        self.assertAlmostEqual(y, 814, places=1)

    def test_alinhado_a_direita_termina_em_x(self):
        """Alinhado à direita o texto TERMINA em x, então começa em
        x - largura. É isso que encosta o carimbo na margem direita."""
        largura_texto = stringWidth(TEXTO_VALOR, "Helvetica", 10)
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 567, 814,
            False, 0, "direita")
        x, y = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, 567 - largura_texto, places=1)
        self.assertAlmostEqual(y, 814, places=1)
        self.assertLess(x, 567)

    def test_padrao_continua_alinhando_a_esquerda(self):
        """Sem o parâmetro novo, o texto sai na mesma posição de sempre."""
        sem_parametro = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, False)
        com_parametro = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, False,
            0, "esquerda")
        self.assertEqual(posicao_do_texto(sem_parametro), (40.0, 40.0))
        self.assertEqual(posicao_do_texto(sem_parametro),
                         posicao_do_texto(com_parametro))

    def test_centralizado_continua_centralizando(self):
        """centralizado=True precisa seguir equivalendo a "centro"."""
        largura_texto = stringWidth(TEXTO_VALOR, "Helvetica", 10)
        buffer = app.criar_overlay(
            595.0, 842.0, TEXTO_VALOR, "Helvetica", 10, "#000000", 40, 40, True)
        x, _ = posicao_do_texto(buffer)
        self.assertAlmostEqual(x, (595.0 - largura_texto) / 2, places=1)


class TestCarimbosExtras(unittest.TestCase):
    def setUp(self):
        import tempfile
        from pypdf import PdfWriter
        self.pasta = tempfile.mkdtemp()
        self.entrada = os.path.join(self.pasta, "entrada.pdf")
        self.saida = os.path.join(self.pasta, "saida.pdf")
        escritor = PdfWriter()
        escritor.add_blank_page(width=595, height=842)
        escritor.add_blank_page(width=595, height=842)
        with open(self.entrada, "wb") as f:
            escritor.write(f)
        self.config = {
            "fonte": "Helvetica", "tamanho": 10, "cor": "#000000",
            "x": 40, "y": 40, "centralizado": False, "angulo": 90,
        }

    def tearDown(self):
        import shutil
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_sem_extras_carimba_so_o_principal(self):
        app.processar_pdf(self.entrada, self.saida, "10004 KLOSTERS", self.config)
        texto = app.extrair_texto_pdf(self.saida)
        self.assertIn("KLOSTERS", texto)
        self.assertNotIn("57,75", texto)

    def test_os_dois_carimbos_saem_em_todas_as_paginas(self):
        extras = [{
            "texto": "15 un × R$ 3,85 = R$ 57,75", "fonte": "Helvetica",
            "tamanho": 10, "cor": "#000000", "x": 567, "y": 814,
            "centralizado": False, "angulo": 0, "alinhamento": "direita",
        }]
        app.processar_pdf(self.entrada, self.saida, "10004 KLOSTERS",
                          self.config, extras)
        from pypdf import PdfReader
        paginas = PdfReader(self.saida).pages
        self.assertEqual(len(paginas), 2)
        for pagina in paginas:
            texto = pagina.extract_text()
            self.assertIn("KLOSTERS", texto)
            self.assertIn("57,75", texto)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_carimbo_lateral -v`
Expected: FAIL — `TypeError: criar_overlay() takes 9 positional arguments but 11 were given`

- [ ] **Step 3: Write minimal implementation**

Substituir a assinatura de `criar_overlay` (linha 731) e o ramo `else` (linhas 753-760):

```python
def criar_overlay(largura, altura, texto, fonte, tamanho, cor, x, y, centralizado,
                  angulo=0, alinhamento="esquerda"):
    """Desenha `texto` no PDF. Se tiver quebras de linha ("\n"), cada linha é
    desenhada empilhada, a primeira em cima e as seguintes abaixo dela.

    `angulo=90` é um modo especial (protocolo dos Correios, ver
    montar_texto_protocolo_correio): ignora x/y/centralizado e desenha uma
    linha única rotacionada 90° (sentido anti-horário — lê de baixo pra
    cima), colada perto da borda direita e verticalmente centralizada.

    `alinhamento` ("esquerda", "centro", "direita") vale para o modo normal:
    "direita" faz o texto TERMINAR em x, usado pelo carimbo do valor no topo
    direito do protocolo. `centralizado=True` continua equivalendo a
    "centro", para não quebrar quem já chamava a função."""
```

E dentro do `else`:

```python
    else:
        altura_linha = tamanho * 1.2
        for i, linha in enumerate(texto.split("\n")):
            x_linha = x
            if centralizado or alinhamento == "centro":
                largura_texto = c.stringWidth(linha, fonte, tamanho)
                x_linha = (largura - largura_texto) / 2
            elif alinhamento == "direita":
                largura_texto = c.stringWidth(linha, fonte, tamanho)
                x_linha = x - largura_texto
            c.drawString(x_linha, y - i * altura_linha, linha)
```

Substituir `processar_pdf` inteira (linhas 767-784):

```python
def processar_pdf(caminho_entrada, caminho_saida, texto, config, carimbos_extras=None):
    """
    Carimba o PDF e grava a saída. `carimbos_extras` permite mais de um
    carimbo por página num único passe de escrita — usado pelo protocolo dos
    Correios, que leva o código na lateral e o valor no topo direito. Sem
    ele, o comportamento é o de sempre: um carimbo só, vindo de `config`.
    """
    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()
    for pagina in reader.pages:
        largura = float(pagina.mediabox.width)
        altura = float(pagina.mediabox.height)

        carimbos = [{
            "texto": texto,
            "fonte": config["fonte"],
            "tamanho": config["tamanho"],
            "cor": config["cor"],
            "x": config["x"],
            "y": config["y"],
            "centralizado": config["centralizado"],
            "angulo": config.get("angulo", 0),
            "alinhamento": config.get("alinhamento", "esquerda"),
        }]
        carimbos.extend(carimbos_extras or [])

        for carimbo in carimbos:
            overlay_buffer = criar_overlay(
                largura, altura, carimbo["texto"],
                carimbo["fonte"], carimbo["tamanho"], carimbo["cor"],
                carimbo["x"], carimbo["y"], carimbo["centralizado"],
                carimbo.get("angulo", 0), carimbo.get("alinhamento", "esquerda"),
            )
            pagina.merge_page(PdfReader(overlay_buffer).pages[0])

        writer.add_page(pagina)

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "wb") as f:
        writer.write(f)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS — inclusive os testes antigos de `criar_overlay`, que não passam os parâmetros novos.

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_carimbo_lateral.py
git commit -m "Permite mais de um carimbo por página e alinhamento à direita"
```

---

### Task 4: OCR do documento inteiro (`logica.py`)

**Files:**
- Modify: `logica.py:314-348` (`extrair_texto_ocr`)
- Test: `tests/test_protocolo_contagem.py` (acrescentar classe)

**Interfaces:**
- Produces:
  - `paginas_para_ocr(total_paginas: int, max_paginas: int | None) -> int`
  - `extrair_texto_ocr(caminho, max_paginas=2, dpi=300)` passa a aceitar `max_paginas=None` (todas as páginas).

- [ ] **Step 1: Write the failing test**

Acrescentar a `tests/test_protocolo_contagem.py`:

```python
class TestPaginasParaOcr(unittest.TestCase):
    """A regra de quantas páginas ler foi extraída para uma função pura
    justamente para ser testável sem winocr (que só existe no Windows)."""

    def test_none_le_todas(self):
        self.assertEqual(app.paginas_para_ocr(5, None), 5)

    def test_limite_menor_que_o_documento(self):
        self.assertEqual(app.paginas_para_ocr(5, 2), 2)

    def test_limite_maior_que_o_documento(self):
        self.assertEqual(app.paginas_para_ocr(1, 2), 1)

    def test_limite_zero_ou_negativo_nao_le_nada(self):
        self.assertEqual(app.paginas_para_ocr(5, 0), 0)
        self.assertEqual(app.paginas_para_ocr(5, -1), 0)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_protocolo_contagem.TestPaginasParaOcr -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'paginas_para_ocr'`

- [ ] **Step 3: Write minimal implementation**

Em `logica.py`, imediatamente antes de `extrair_texto_ocr` (linha 314):

```python
def paginas_para_ocr(total_paginas, max_paginas):
    """Quantas páginas o OCR deve ler. `max_paginas=None` significa todas —
    usado pelos protocolos dos Correios, onde parar na 2ª página perderia
    unidades em silêncio."""
    if max_paginas is None:
        return total_paginas
    return min(total_paginas, max(0, max_paginas))
```

E dentro de `extrair_texto_ocr`, trocar o corte por página:

```python
    textos = []
    doc = fitz.open(caminho)
    try:
        limite = paginas_para_ocr(doc.page_count, max_paginas)
        for i, pagina in enumerate(doc):
            if i >= limite:
                break
```

O resto do corpo da função fica igual. Atualizar a docstring, trocando "as primeiras páginas" por:

```python
    """
    Renderiza páginas do PDF como imagem e roda OCR usando o motor nativo do
    Windows (winocr) — sem programas externos instalados. Lê no máximo
    `max_paginas` páginas; `max_paginas=None` lê o documento inteiro.
    Tenta português (pt-BR) primeiro; se não disponível, usa inglês (en-US).
    """
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_contagem.py
git commit -m "Permite OCR do documento inteiro (max_paginas=None)"
```

---

### Task 4b: RapidOCR como reserva opcional (`logica.py`)

**Files:**
- Modify: `logica.py` (após `extrair_texto_ocr`)
- Test: `tests/test_protocolo_contagem.py` (acrescentar classe)

**Interfaces:**
- Consumes: `extrair_texto_ocr`, `OCR_DISPONIVEL` (já existem).
- Produces:
  - `RAPIDOCR_DISPONIVEL: bool`
  - `extrair_texto_rapidocr(caminho, dpi=300) -> str`
  - `extrair_texto_escaneado(caminho, dpi=300) -> str` — winocr com RapidOCR de reserva

**Regra:** a reserva entra **só** quando o winocr não existe ou levanta erro.
Nunca quando ele lê e a contagem não confere — aí o caminho continua sendo
qualidade maior e, persistindo, pendente. Dependência opcional: não entra no
`requirements.txt` nem no `.spec`, e a ausência dela não pode quebrar nada.

- [ ] **Step 1: Write the failing test**

Acrescentar a `tests/test_protocolo_contagem.py`:

```python
class TestReservaDeOcr(unittest.TestCase):
    """A reserva é escolhida por uma função pura, testável sem nenhum dos
    dois motores instalados."""

    def test_usa_winocr_quando_disponivel(self):
        self.assertEqual(app.motor_de_ocr(tem_winocr=True, tem_rapidocr=True), "winocr")
        self.assertEqual(app.motor_de_ocr(tem_winocr=True, tem_rapidocr=False), "winocr")

    def test_cai_para_rapidocr_sem_winocr(self):
        self.assertEqual(app.motor_de_ocr(tem_winocr=False, tem_rapidocr=True), "rapidocr")

    def test_sem_nenhum_motor(self):
        self.assertIsNone(app.motor_de_ocr(tem_winocr=False, tem_rapidocr=False))

    def test_flag_de_disponibilidade_existe(self):
        self.assertIsInstance(app.RAPIDOCR_DISPONIVEL, bool)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_protocolo_contagem.TestReservaDeOcr -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'motor_de_ocr'`

- [ ] **Step 3: Write minimal implementation**

Junto do bloco onde `OCR_DISPONIVEL` é definido, acrescentar o import opcional:

```python
#  Reserva opcional: modelos PP-OCR (do PaddleOCR) rodando em ONNX. NÃO entra
#  no requirements.txt nem no .spec — quem tiver instalado na própria máquina
#  ganha a reserva, e o CODIFICADOR.zip continua do tamanho de hoje.
try:
    from rapidocr import RapidOCR
    RAPIDOCR_DISPONIVEL = True
except Exception:
    RapidOCR = None
    RAPIDOCR_DISPONIVEL = False

_rapidocr_motor = None
```

E após `extrair_texto_ocr`:

```python
def motor_de_ocr(tem_winocr=None, tem_rapidocr=None):
    """
    Qual motor usar: "winocr", "rapidocr" ou None se nenhum existir. O winocr
    sempre ganha quando está disponível — a reserva cobre a máquina onde o
    motor nativo não existe, não a leitura que deu resultado ruim.

    Os parâmetros existem para o teste; em produção ficam None e a função
    consulta as flags do módulo.
    """
    if tem_winocr is None:
        tem_winocr = OCR_DISPONIVEL
    if tem_rapidocr is None:
        tem_rapidocr = RAPIDOCR_DISPONIVEL
    if tem_winocr:
        return "winocr"
    if tem_rapidocr:
        return "rapidocr"
    return None


def extrair_texto_rapidocr(caminho, dpi=300):
    """
    OCR pelos modelos PP-OCR em ONNX. Devolve um bloco de texto por região
    detectada, separados por quebra de linha — é a estrutura real da tabela
    do protocolo, e as regexes de contagem funcionam igual.
    """
    global _rapidocr_motor
    if not RAPIDOCR_DISPONIVEL:
        raise RuntimeError("RapidOCR não disponível.")

    import numpy as np
    if _rapidocr_motor is None:
        _rapidocr_motor = RapidOCR()   # carregar os modelos é caro; reaproveita

    pedacos = []
    doc = fitz.open(caminho)
    try:
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            resultado = _rapidocr_motor(np.array(img))
            textos = getattr(resultado, "txts", None)
            if textos is None:                      # API antiga: (lista, tempo)
                textos = [linha[1] for linha in (resultado[0] or [])]
            pedacos.extend(textos or [])
    finally:
        doc.close()
    return "\n".join(pedacos)


def extrair_texto_escaneado(caminho, dpi=300):
    """
    Texto de um PDF escaneado, com o winocr na frente e o RapidOCR de reserva.
    A reserva só entra quando o winocr não existe ou quebra — uma leitura que
    funcionou nunca é substituída.
    """
    if OCR_DISPONIVEL:
        try:
            return extrair_texto_ocr(caminho, max_paginas=None, dpi=dpi)
        except Exception:
            if not RAPIDOCR_DISPONIVEL:
                raise
    if RAPIDOCR_DISPONIVEL:
        return extrair_texto_rapidocr(caminho, dpi=dpi)
    raise RuntimeError(
        "Nenhum leitor de documentos escaneados disponível. "
        "Instale: pip install pymupdf winocr")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_contagem.py
git commit -m "RapidOCR como reserva opcional quando o winocr não existe"
```

---

### Task 5: Planilha dos protocolos (`logica.py`)

**Files:**
- Modify: `logica.py` (após `salvar_planilha_nfse`, fim do arquivo)
- Test: `tests/test_planilha_protocolo.py`

**Interfaces:**
- Consumes: `valor_protocolo` (Task 2), `_codigos_do_cadastro(cadastro)` (já existe).
- Produces:
  - `COLUNAS_PROTOCOLO` — lista de `(titulo, largura, formato)`
  - `linha_planilha_protocolo(nome_arquivo, dados, cadastro, tarifa=None, unidades=None, observacao="") -> list`
  - `salvar_planilha_protocolo(caminho, linhas) -> None`

- [ ] **Step 1: Write the failing test**

Criar `tests/test_planilha_protocolo.py`:

```python
import os
import shutil
import sys
import tempfile
import unittest
from decimal import Decimal

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)
_AQUI = os.path.dirname(os.path.abspath(__file__))
if _AQUI not in sys.path:
    sys.path.insert(0, _AQUI)

import logica as app
from cadastro_teste import CADASTRO_TESTE

DADOS_CADASTRADO = {
    "codigo": "10004", "condominio": "W700A KLOSTERS",
    "total_impresso": 15, "linhas_contadas": 15, "entregas_contadas": 15,
}
DADOS_NAO_CADASTRADO = {
    "codigo": "99999", "condominio": "W700A DESCONHECIDO",
    "total_impresso": 3, "linhas_contadas": 3, "entregas_contadas": 3,
}


class TestLinhaPlanilhaProtocolo(unittest.TestCase):
    def test_codigo_cadastrado_usa_o_nome_do_cadastro(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=15)
        self.assertEqual(linha[1], "KLOSTERS")
        self.assertEqual(linha[2], "10004")
        self.assertEqual(linha[3], 15)
        self.assertEqual(linha[5], 57.75)
        self.assertEqual(linha[6], "")

    def test_codigo_nao_cadastrado_mantem_o_nome_do_documento(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
            tarifa=Decimal("3.85"), unidades=3)
        self.assertEqual(linha[1], "W700A DESCONHECIDO")
        self.assertEqual(linha[6], "Código não cadastrado")
        self.assertEqual(linha[5], 11.55)

    def test_pendente_deixa_valores_vazios_nunca_zero(self):
        """0 significaria "entregou zero unidades" — mesma regra das
        retenções federais na planilha das NFS-e."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            observacao="Listando 15, mas foram contadas 14 e 14 unidades")
        self.assertIsNone(linha[3])
        self.assertIsNone(linha[4])
        self.assertIsNone(linha[5])
        self.assertIn("Listando 15", linha[6])

    def test_nao_e_protocolo(self):
        linha = app.linha_planilha_protocolo(
            "outro.pdf", None, CADASTRO_TESTE,
            observacao="Não é um protocolo dos Correios")
        self.assertEqual(linha[0], "outro.pdf")
        self.assertEqual(len(linha), len(app.COLUNAS_PROTOCOLO))
        self.assertIsNone(linha[5])


class TestSalvarPlanilhaProtocolo(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.destino = os.path.join(self.pasta, "protocolos.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_grava_numeros_e_linha_de_total(self):
        from openpyxl import load_workbook
        linhas = [
            app.linha_planilha_protocolo("a.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                         Decimal("3.85"), 15),
            app.linha_planilha_protocolo("b.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
                                         Decimal("3.85"), 3),
            app.linha_planilha_protocolo("c.pdf", None, CADASTRO_TESTE,
                                         observacao="Não é um protocolo dos Correios"),
        ]
        app.salvar_planilha_protocolo(self.destino, linhas)

        sheet = load_workbook(self.destino).active
        self.assertEqual(sheet.cell(row=1, column=1).value, "Arquivo")
        # Valor gravado como número, não texto (dá para somar no Excel)
        self.assertEqual(sheet.cell(row=2, column=6).value, 57.75)
        self.assertIsInstance(sheet.cell(row=2, column=6).value, float)
        # Linha de total logo abaixo dos dados
        linha_total = sheet.max_row
        self.assertEqual(sheet.cell(row=linha_total, column=1).value, "TOTAL")
        self.assertEqual(sheet.cell(row=linha_total, column=4).value, 18)
        self.assertEqual(sheet.cell(row=linha_total, column=6).value, 69.30)

    def test_planilha_vazia_nao_quebra(self):
        from openpyxl import load_workbook
        app.salvar_planilha_protocolo(self.destino, [])
        sheet = load_workbook(self.destino).active
        self.assertEqual(sheet.cell(row=sheet.max_row, column=1).value, "TOTAL")
        self.assertEqual(sheet.cell(row=sheet.max_row, column=6).value, 0)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_planilha_protocolo -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'linha_planilha_protocolo'`

- [ ] **Step 3: Write minimal implementation**

No fim de `logica.py`:

```python
# ============================================================
#  PLANILHA DOS PROTOCOLOS DOS CORREIOS
# ============================================================

COLUNAS_PROTOCOLO = [
    ("Arquivo", 38, None),
    ("Condomínio", 30, None),
    ("Código", 10, None),
    ("Unidades", 10, "0"),
    ("Tarifa", 12, "R$ #,##0.00"),
    ("Valor", 14, "R$ #,##0.00"),
    ("Observação", 44, None),
]


def linha_planilha_protocolo(nome_arquivo, dados, cadastro, tarifa=None,
                             unidades=None, observacao=""):
    """
    Monta a linha da planilha. O nome do condomínio vem do cadastro quando o
    código está lá; senão fica o que o próprio documento traz no cabeçalho.

    `unidades=None` é o caso pendente (contagem recusada) ou o de um arquivo
    que nem é protocolo: Unidades, Tarifa e Valor saem VAZIOS, nunca 0 — 0
    significaria "entregou zero unidades".
    """
    if dados is None:
        return [nome_arquivo, "", "", None, None, None, observacao]

    codigo = dados.get("codigo") or ""
    registro = None
    if codigo:
        cnpj = _codigos_do_cadastro(cadastro).get(codigo)
        registro = cadastro.get(cnpj) if cnpj else None

    condominio = registro["nome"] if registro else dados.get("condominio", "")
    if not registro and not observacao:
        observacao = "Código não cadastrado"

    if unidades is None:
        return [nome_arquivo, condominio, codigo, None, None, None, observacao]

    valor = valor_protocolo(unidades, tarifa)
    return [nome_arquivo, condominio, codigo, int(unidades),
            float(tarifa), float(valor), observacao]


def salvar_planilha_protocolo(caminho, linhas):
    """
    Grava a planilha dos protocolos com linha de TOTAL no rodapé. Os totais
    são calculados aqui em Python (não como fórmula do Excel) para que o
    arquivo já chegue com o número pronto, sem depender de o Excel abrir e
    recalcular.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Protocolos"

    sheet.append([c[0] for c in COLUNAS_PROTOCOLO])
    for celula in sheet[1]:
        celula.font = Font(bold=True)

    for linha in linhas:
        sheet.append(linha)

    ultima_dados = sheet.max_row

    for indice, (_, largura, formato) in enumerate(COLUNAS_PROTOCOLO, start=1):
        letra = sheet.cell(row=1, column=indice).column_letter
        sheet.column_dimensions[letra].width = largura
        if formato:
            for numero_linha in range(2, ultima_dados + 1):
                sheet.cell(row=numero_linha, column=indice).number_format = formato

    sheet.freeze_panes = "A2"
    ultima_coluna = sheet.cell(row=1, column=len(COLUNAS_PROTOCOLO)).column_letter
    sheet.auto_filter.ref = f"A1:{ultima_coluna}{ultima_dados}"

    #  Total depois do autofiltro, para não virar uma linha filtrável
    linha_total = ultima_dados + 1
    celula_rotulo = sheet.cell(row=linha_total, column=1, value="TOTAL")
    celula_rotulo.font = Font(bold=True)
    for coluna in (4, 6):   # Unidades e Valor
        total = sum(
            linha[coluna - 1] for linha in linhas
            if isinstance(linha[coluna - 1], (int, float))
        )
        celula = sheet.cell(row=linha_total, column=coluna, value=round(total, 2))
        celula.font = Font(bold=True)
        celula.number_format = COLUNAS_PROTOCOLO[coluna - 1][2]

    wb.save(caminho)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_planilha_protocolo.py
git commit -m "Planilha dos protocolos com linha de total"
```

---

### Task 6: Aba 3 na interface (`identificacao_por_cnpj_6_0.py`)

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py:55-62` (import), `:178-179` (StringVars), `:272-297` (`_montar_interface`), e um bloco novo de métodos após `_concluir_extracao` (linha ~2211)
- Test: verificação manual (interface não tem teste automatizado neste projeto)

**Interfaces:**
- Consumes: `extrair_dados_protocolo_correio`, `conferir_contagem_protocolo`, `valor_protocolo`, `montar_texto_valor_protocolo`, `formatar_reais`, `linha_planilha_protocolo`, `salvar_planilha_protocolo`, `COLUNAS_PROTOCOLO`, `paginas_para_ocr`, `extrair_texto_escaneado` (Tasks 1, 2, 4, 4b, 5); `montar_texto_protocolo_correio`, `_codigos_do_cadastro`, `extrair_texto_pdf`, `extrair_texto_ocr`, `proximo_dpi_maior`, `processar_pdf`, `formatar_cnpj`, `LIMITE_TEXTO_MINIMO` (já existiam).
- Produces: nada consumido por tasks posteriores.

- [ ] **Step 1: Ampliar o import e declarar as variáveis da aba**

Em `identificacao_por_cnpj_6_0.py`, acrescentar ao `from logica import (...)` (linha ~55):

```python
    extrair_dados_protocolo_correio, conferir_contagem_protocolo,
    valor_protocolo, montar_texto_valor_protocolo, formatar_reais,
    linha_planilha_protocolo, salvar_planilha_protocolo,
    extrair_texto_escaneado,
```

Após a linha 179 (`self.arquivo_planilha_saida = tk.StringVar()`):

```python
        self.pasta_protocolos = tk.StringVar()
        self.pasta_protocolos_saida = tk.StringVar()
        self.arquivo_planilha_protocolos = tk.StringVar()
```

- [ ] **Step 2: Registrar a aba e renumerar as outras**

Substituir o bloco de criação das abas em `_montar_interface` (linhas 277-291):

```python
        self.aba_processar = ttk.Frame(notebook)
        self.aba_extracao = ttk.Frame(notebook)
        self.aba_protocolos = ttk.Frame(notebook)
        self.aba_cadastro = ttk.Frame(notebook)
        self.aba_logs = ttk.Frame(notebook)
        notebook.add(self.aba_processar, text="1. Processamento")
        notebook.add(self.aba_extracao, text="2. Extrair dados")
        notebook.add(self.aba_protocolos, text="3. Protocolos dos Correios")
        notebook.add(self.aba_cadastro, text="4. Cadastro de Condomínios")
        notebook.add(self.aba_logs, text="5. Logs")

        # A ordem importa: _montar_aba_processar cria self._widgets_tema, que
        # as abas seguintes usam para registrar os widgets delas no tema.
        self._montar_aba_processar(self.aba_processar)
        self._montar_aba_extracao(self.aba_extracao)
        self._montar_aba_protocolos(self.aba_protocolos)
        self._montar_aba_cadastro(self.aba_cadastro)
        self._montar_aba_logs(self.aba_logs)
```

Corrigir também o comentário obsoleto `#  ABA 2 — CADASTRO` (linha 300) para `#  ABA 4 — CADASTRO`.

- [ ] **Step 3: Montar a aba**

Inserir após `_concluir_extracao` (linha ~2211):

```python
    # --------------------------------------------------------
    #  ABA 3 — PROTOCOLOS DOS CORREIOS
    # --------------------------------------------------------
    def _montar_aba_protocolos(self, parent_externo):
        """
        Conta as unidades de cada protocolo, calcula o valor pela tarifa do
        lote, carimba o PDF e gera a planilha. Mesma linguagem visual das
        outras abas (Swiss, cantos retos, cobalto só no botão primário).
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        def registrar(widget, mapa):
            self._widgets_tema.append((widget, mapa))
            for prop, chave in mapa.items():
                try:
                    widget.configure(**{prop: tema[chave]})
                except Exception:
                    pass
            return widget

        container = registrar(
            ctk.CTkFrame(parent_externo, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        container.pack(fill="both", expand=True)

        bloco_titulo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_titulo.pack(fill="x", padx=24, pady=(24, 0))

        titulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="Protocolos dos Correios", font=(fonte, 20),
                         text_color=tema["texto"], anchor="w"),
            {"text_color": "texto"},
        )
        titulo.pack(anchor="w")

        subtitulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="CONTAGEM DE UNIDADES E VALOR", font=(fonte, 11),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        subtitulo.pack(anchor="w")

        hairline = registrar(
            ctk.CTkFrame(container, height=1, corner_radius=0, fg_color=tema["borda"]),
            {"fg_color": "borda"},
        )
        hairline.pack(fill="x", padx=24, pady=(16, 24))

        corpo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        corpo.pack(fill="both", expand=True, padx=24)
        corpo.columnconfigure(0, weight=1)

        def montar_campo(linha_grid, rotulo, variavel, comando_trocar):
            label = registrar(
                ctk.CTkLabel(corpo, text=rotulo, font=(fonte, 11),
                             text_color=tema["texto_secundario"], anchor="w"),
                {"text_color": "texto_secundario"},
            )
            label.grid(row=linha_grid, column=0, sticky="w", pady=(0, 4))

            linha = registrar(
                ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"]),
                {"fg_color": "fundo"},
            )
            linha.grid(row=linha_grid + 1, column=0, sticky="ew", pady=(0, 16))
            linha.columnconfigure(0, weight=1)

            entry = registrar(
                ctk.CTkEntry(linha, textvariable=variavel, corner_radius=0,
                             fg_color=tema["superficie"], border_width=1,
                             border_color=tema["borda"], text_color=tema["texto"],
                             font=(fonte, 13)),
                {"fg_color": "superficie", "border_color": "borda", "text_color": "texto"},
            )
            entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

            botao = registrar(
                ctk.CTkButton(
                    linha, text="Trocar", corner_radius=0, width=90,
                    fg_color="transparent", hover_color=tema["superficie"],
                    border_width=1, border_color=tema["borda_forte"],
                    text_color=tema["texto"], font=(fonte, 13), command=comando_trocar,
                ),
                {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
            )
            botao.grid(row=0, column=1)

        montar_campo(0, "PASTA COM OS PROTOCOLOS", self.pasta_protocolos,
                     self._escolher_pasta_protocolos)
        montar_campo(2, "SALVAR OS PDFS CARIMBADOS EM", self.pasta_protocolos_saida,
                     self._escolher_pasta_protocolos_saida)
        montar_campo(4, "SALVAR PLANILHA EM", self.arquivo_planilha_protocolos,
                     self._escolher_planilha_protocolos)

        self.botao_protocolos = registrar(
            ctk.CTkButton(
                corpo, text="Calcular protocolos", corner_radius=0, height=44,
                font=(fonte, 15), fg_color=tema["acento"],
                hover_color=tema["acento_hover"], text_color=tema["sobre_acento"],
                border_width=0, command=self._iniciar_protocolos,
            ),
            {"fg_color": "acento", "hover_color": "acento_hover", "text_color": "sobre_acento"},
        )
        self.botao_protocolos.grid(row=6, column=0, sticky="ew", pady=(8, 8))

        explicacao = registrar(
            ctk.CTkLabel(
                corpo,
                text=("Conta quantas unidades cada protocolo entregou, multiplica pelo "
                      "valor por linha que você informar e escreve o total no canto "
                      "superior direito do PDF, junto do código do condomínio. "
                      "Protocolos em que a contagem não confere ficam sem carimbo e "
                      "aparecem na planilha com o motivo."),
                font=(fonte, 13), text_color=tema["texto_terciario"],
                justify="left", anchor="w", wraplength=640,
            ),
            {"text_color": "texto_terciario"},
        )
        explicacao.grid(row=7, column=0, sticky="w", pady=(0, 16))

        self.barra_protocolos = ttk.Progressbar(corpo, mode="determinate")
        self.barra_protocolos.grid(row=8, column=0, sticky="ew", pady=(0, 8))

        self.label_status_protocolos = registrar(
            ctk.CTkLabel(corpo, text="", font=(fonte, 13),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        self.label_status_protocolos.grid(row=9, column=0, sticky="w", pady=(0, 24))

        self._atualizar_botao_protocolos()

    def _escolher_pasta_protocolos(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os protocolos")
        if pasta:
            self.pasta_protocolos.set(pasta)
            nome_pasta = os.path.basename(os.path.normpath(pasta)) or "protocolos"
            self.pasta_protocolos_saida.set(os.path.join(pasta, "carimbados"))
            self.arquivo_planilha_protocolos.set(
                os.path.join(pasta, f"protocolos_{nome_pasta}.xlsx"))
        self._atualizar_botao_protocolos()

    def _escolher_pasta_protocolos_saida(self):
        pasta = filedialog.askdirectory(title="Onde salvar os PDFs carimbados")
        if pasta:
            self.pasta_protocolos_saida.set(pasta)
        self._atualizar_botao_protocolos()

    def _escolher_planilha_protocolos(self):
        atual = self.arquivo_planilha_protocolos.get().strip()
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=os.path.basename(atual) if atual else "protocolos.xlsx",
            initialdir=os.path.dirname(atual) if atual else None,
        )
        if caminho:
            self.arquivo_planilha_protocolos.set(caminho)
        self._atualizar_botao_protocolos()

    def _atualizar_botao_protocolos(self):
        """Conta os *.pdf da pasta escolhida e ajusta texto/estado do botão."""
        pasta = self.pasta_protocolos.get().strip()
        quantidade = 0
        if pasta and os.path.isdir(pasta):
            try:
                quantidade = sum(1 for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
            except Exception:
                quantidade = 0

        pronto = (quantidade > 0
                  and self.pasta_protocolos_saida.get().strip()
                  and self.arquivo_planilha_protocolos.get().strip())
        if pronto:
            plural = "protocolo" if quantidade == 1 else "protocolos"
            self.botao_protocolos.configure(
                text=f"Calcular {quantidade} {plural}", state="normal")
        else:
            self.botao_protocolos.configure(text="Calcular protocolos", state="disabled")
```

- [ ] **Step 4: Verificar que a aba aparece**

Run: `python identificacao_por_cnpj_6_0.py`
Expected: cinco abas, a terceira chamando-se "3. Protocolos dos Correios", com os três campos e o botão desabilitado até escolher pasta e destinos. Fechar o app.

- [ ] **Step 5: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Aba 3 dos protocolos dos Correios (campos e estado do botão)"
```

- [ ] **Step 6: Popup da tarifa**

Inserir após `_atualizar_botao_protocolos`:

```python
    def _pedir_tarifa(self):
        """
        Pede o valor por linha do lote. Campo vazio de propósito, sem valor
        padrão: a tarifa muda com o tempo (nos protocolos de referência
        aparecem 3,45 e 3,85) e um padrão herdado passaria batido.
        Devolve Decimal, ou None se o usuário cancelar.
        """
        from decimal import Decimal, InvalidOperation

        tema = self.tema_atual
        fonte = familia_fonte()
        janela = ctk.CTkToplevel(self)
        janela.title("Valor por linha")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(self)
        janela.grab_set()

        resultado = {"valor": None}

        ctk.CTkLabel(janela, text="Quanto custa cada linha entregue?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text="Ex.: 3,85", font=(fonte, 12),
                     text_color=tema["texto_terciario"]).pack(padx=24, anchor="w")

        entrada = ctk.CTkEntry(janela, corner_radius=0, width=200,
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(padx=24, pady=(12, 4), anchor="w")
        entrada.focus_set()

        aviso = ctk.CTkLabel(janela, text="", font=(fonte, 12),
                             text_color=tema["acento"])
        aviso.pack(padx=24, pady=(0, 8), anchor="w")

        def confirmar():
            texto = entrada.get().strip().replace("R$", "").replace(" ", "")
            texto = texto.replace(".", "").replace(",", ".") if "," in texto else texto
            try:
                valor = Decimal(texto)
            except (InvalidOperation, ValueError):
                aviso.configure(text="Digite um número, como 3,85.")
                return
            if valor <= 0:
                aviso.configure(text="O valor precisa ser maior que zero.")
                return
            resultado["valor"] = valor
            janela.destroy()

        linha_botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        linha_botoes.pack(padx=24, pady=(0, 24), anchor="e")

        ctk.CTkButton(linha_botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(linha_botoes, text="Continuar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return resultado["valor"]
```

- [ ] **Step 7: Verificar o popup**

Run: `python identificacao_por_cnpj_6_0.py`
Expected: ainda não há como abrir o popup pela interface (o botão chama `_iniciar_protocolos`, criado no passo seguinte). Só conferir que o app abre sem erro de sintaxe. Fechar.

- [ ] **Step 8: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Popup que pede a tarifa por linha do lote"
```

- [ ] **Step 9: Disparo e thread de processamento**

Inserir após `_pedir_tarifa`:

```python
    def _iniciar_protocolos(self):
        pasta = self.pasta_protocolos.get().strip()
        pasta_saida = self.pasta_protocolos_saida.get().strip()
        destino = self.arquivo_planilha_protocolos.get().strip()

        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione a pasta com os protocolos.")
            return
        if not pasta_saida:
            messagebox.showerror("Erro", "Escolha onde salvar os PDFs carimbados.")
            return
        if os.path.normpath(pasta_saida) == os.path.normpath(pasta):
            messagebox.showerror(
                "Erro", "A pasta de saída precisa ser diferente da pasta de origem, "
                        "para não sobrescrever os protocolos originais.")
            return
        if not destino:
            messagebox.showerror("Erro", "Escolha onde salvar a planilha.")
            return
        if os.path.isfile(destino) and not messagebox.askyesno(
            "Substituir planilha",
            f"Este arquivo já existe e será substituído:\n\n{destino}\n\nContinuar?"
        ):
            return

        tarifa = self._pedir_tarifa()
        if tarifa is None:
            return

        self.botao_protocolos.configure(state="disabled")
        self.label_status_protocolos.configure(text="Lendo os protocolos...")

        thread = threading.Thread(
            target=self._processar_protocolos_em_thread,
            args=(pasta, pasta_saida, destino, tarifa), daemon=True)
        thread.start()

    def _ler_texto_protocolo(self, caminho, dpi):
        """Texto nativo se houver; senão o documento inteiro pelo leitor de
        escaneados (winocr, com RapidOCR de reserva se estiver instalado).
        Todas as páginas — parar na 2ª perderia unidades."""
        texto = extrair_texto_pdf(caminho)
        if len(texto.strip()) >= LIMITE_TEXTO_MINIMO:
            return texto
        return extrair_texto_escaneado(caminho, dpi=dpi)

    def _processar_protocolos_em_thread(self, pasta, pasta_saida, destino, tarifa):
        arquivos = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
        total = len(arquivos)
        self.after(0, lambda: self.barra_protocolos.configure(maximum=total, value=0))

        config = self.montar_config_atual()
        codigos = _codigos_do_cadastro(self.cadastro)

        linhas = []
        carimbados = 0
        pendentes = 0
        ignorados = 0

        for indice, nome in enumerate(arquivos, 1):
            caminho = os.path.join(pasta, nome)
            self.after(0, lambda i=indice, n=nome:
                       self.label_status_protocolos.configure(
                           text=f"Lendo {i} de {total}: {n}"))

            dados = None
            unidades = None
            observacao = ""
            try:
                dpi = config.get("dpi", 300)
                texto = self._ler_texto_protocolo(caminho, dpi)
                dados = extrair_dados_protocolo_correio(texto)

                if dados is None:
                    observacao = "Não é um protocolo dos Correios"
                else:
                    aceito, motivo = conferir_contagem_protocolo(dados)
                    if not aceito:
                        #  Uma re-tentativa em qualidade maior, como já se faz
                        #  quando o CNPJ sai com checksum inválido.
                        dpi_maior = proximo_dpi_maior(dpi)
                        if dpi_maior and dpi_maior != dpi:
                            texto = extrair_texto_ocr(caminho, max_paginas=None,
                                                      dpi=dpi_maior)
                            novos = extrair_dados_protocolo_correio(texto)
                            if novos:
                                dados = novos
                                aceito, motivo = conferir_contagem_protocolo(dados)
                    if aceito:
                        unidades = dados["total_impresso"]
                    else:
                        observacao = motivo
            except Exception as e:
                observacao = f"Erro ao ler: {e}"

            if unidades is not None:
                try:
                    self._carimbar_protocolo(caminho, pasta_saida, nome, dados,
                                             unidades, tarifa, config, codigos)
                    carimbados += 1
                except Exception as e:
                    observacao = f"Erro ao carimbar: {e}"
            elif dados is None:
                ignorados += 1
            else:
                pendentes += 1

            linhas.append(linha_planilha_protocolo(
                nome, dados, self.cadastro, tarifa, unidades, observacao))
            self.after(0, lambda v=indice: self.barra_protocolos.configure(value=v))

        try:
            salvar_planilha_protocolo(destino, linhas)
        except Exception as e:
            self.after(0, lambda: self.label_status_protocolos.configure(
                text="Não foi possível salvar a planilha."))
            self.after(0, self._atualizar_botao_protocolos)
            self.after(0, lambda: messagebox.showerror(
                "Erro ao salvar",
                f"Não foi possível gravar a planilha:\n{e}\n\n"
                "Se ela estiver aberta no Excel, feche e tente de novo."))
            return

        total_valor = sum(l[5] for l in linhas if isinstance(l[5], float))
        resumo = f"{carimbados} protocolo(s) · {formatar_reais(total_valor)}"
        if pendentes:
            resumo += f" · {pendentes} sem conferir"
        if ignorados:
            resumo += f" · {ignorados} ignorado(s)"

        self.after(0, lambda: self.label_status_protocolos.configure(text=resumo))
        self.after(0, self._atualizar_botao_protocolos)
        self.after(0, lambda: self._concluir_extracao(destino, resumo))

    def _carimbar_protocolo(self, caminho, pasta_saida, nome, dados, unidades,
                            tarifa, config, codigos):
        """
        Dois carimbos num passe só: o lateral rotacionado com o código (igual
        ao da aba 1, para a IA do Superlógica continuar lendo) e o valor no
        topo direito. Código fora do cadastro carimba só o valor — o valor não
        depende do cadastro, e perder a cobrança por isso seria pior.
        """
        valor = valor_protocolo(unidades, tarifa)
        texto_valor = montar_texto_valor_protocolo(unidades, tarifa, valor)

        cnpj = codigos.get(dados.get("codigo") or "")
        registro = self.cadastro.get(cnpj) if cnpj else None

        config_carimbo = dict(config)
        config_carimbo["angulo"] = 90
        if registro:
            texto_lateral = montar_texto_protocolo_correio(
                dados["codigo"], registro["nome"], cnpj)
        else:
            #  Sem cadastro não há nome nem CNPJ para a linha lateral; o
            #  carimbo principal vira o próprio valor, no topo direito.
            texto_lateral = texto_valor
            config_carimbo["angulo"] = 0
            config_carimbo["centralizado"] = False
            config_carimbo["alinhamento"] = "direita"
            config_carimbo["x"] = MARGEM_VALOR_PROTOCOLO_X
            config_carimbo["y"] = MARGEM_VALOR_PROTOCOLO_Y

        extras = []
        if registro:
            extras.append({
                "texto": texto_valor,
                "fonte": config["fonte"],
                "tamanho": config["tamanho"],
                "cor": config["cor"],
                "x": MARGEM_VALOR_PROTOCOLO_X,
                "y": MARGEM_VALOR_PROTOCOLO_Y,
                "centralizado": False,
                "angulo": 0,
                "alinhamento": "direita",
            })

        caminho_saida = os.path.join(pasta_saida, nome)
        processar_pdf(caminho, caminho_saida, texto_lateral, config_carimbo, extras)
```

O carimbo do valor precisa de duas constantes. O `x` é a borda direita menos a margem, e depende da largura da página — por isso é resolvido em `criar_overlay` a partir do `x` informado. Definir no topo de `identificacao_por_cnpj_6_0.py`, junto das outras constantes de módulo:

```python
#  Posição do carimbo de valor no protocolo dos Correios: canto superior
#  direito, no espaço em branco do documento. O rodapé foi descartado porque
#  protocolos de duas páginas têm conteúdo lá embaixo.
MARGEM_VALOR_PROTOCOLO_X = 567   # 595pt (A4) - 28pt de margem
MARGEM_VALOR_PROTOCOLO_Y = 814   # 842pt (A4) - 28pt de margem
```

Conferir que `montar_config_atual`, `_codigos_do_cadastro`, `montar_texto_protocolo_correio`, `proximo_dpi_maior` e `extrair_texto_pdf` estão importados/definidos no arquivo; acrescentar ao `from logica import (...)` os que faltarem.

- [ ] **Step 10: Rodar a suíte e conferir que nada quebrou**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS

- [ ] **Step 11: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Processa os protocolos: conta, carimba código e valor, gera planilha"
```

---

### Task 7: Aceitação contra os PDFs reais e documentação

**Files:**
- Create: `scratch/aceitacao_protocolos.py` (script temporário, **não commitar**)
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: tudo das Tasks 1–5.

- [ ] **Step 1: Escrever o script de aceitação**

O lote real está em `C:\Users\Dell\Downloads\TESTE CORREIO` (quatro PDFs escaneados). Os valores esperados foram conferidos contra o que está escrito à mão em cada folha.

```python
"""Aceitação: o lote real precisa fechar nos valores manuscritos."""
import glob, os, sys
from decimal import Decimal
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import logica as app

TARIFA = Decimal("3.85")
ESPERADO = {  # unidades, valor manuscrito no papel
    "20260813183654301-001.pdf": (2, Decimal("7.70")),
    "20260813183654301-002.pdf": (15, Decimal("57.75")),
    "20260813183654301-003.pdf": (3, Decimal("11.55")),
    "20260813183654301-004.pdf": (12, Decimal("46.20")),
}

total = Decimal("0.00")
falhas = 0
for caminho in sorted(glob.glob(r"C:\Users\Dell\Downloads\TESTE CORREIO\*.pdf")):
    nome = os.path.basename(caminho)
    texto = app.extrair_texto_pdf(caminho)
    if len(texto.strip()) < app.LIMITE_TEXTO_MINIMO:
        texto = app.extrair_texto_ocr(caminho, max_paginas=None, dpi=300)
    dados = app.extrair_dados_protocolo_correio(texto)
    aceito, motivo = app.conferir_contagem_protocolo(dados)
    unidades = dados["total_impresso"] if aceito else None
    valor = app.valor_protocolo(unidades, TARIFA) if aceito else None
    esperado_un, esperado_valor = ESPERADO[nome]
    ok = aceito and unidades == esperado_un and valor == esperado_valor
    falhas += 0 if ok else 1
    total += valor or Decimal("0.00")
    print(f"{nome} | {dados['codigo']} {dados['condominio']} | "
          f"{unidades} un | {valor} | esperado {esperado_un}/{esperado_valor} | "
          f"{'OK' if ok else 'FALHOU: ' + motivo}")

print(f"TOTAL {total} (esperado 123.20) | falhas: {falhas}")
sys.exit(1 if falhas or total != Decimal("123.20") else 0)
```

- [ ] **Step 2: Rodar a aceitação**

Run: `python scratch/aceitacao_protocolos.py`
Expected: quatro linhas `OK` e `TOTAL 123.20 (esperado 123.20) | falhas: 0`, saída 0.

Se falhar, **não ajustar o esperado** — os valores vieram do papel. Investigar a leitura.

- [ ] **Step 3: Conferir o carimbo a olho**

Rodar o app, processar a pasta `TESTE CORREIO` com tarifa `3,85`, abrir um PDF de saída e confirmar:
- o valor aparece no canto superior direito, sem invadir o cabeçalho;
- o carimbo lateral do código continua legível e na mesma posição de antes;
- em protocolo de duas páginas (pedir um ao usuário — não há nenhum no lote de referência), os dois carimbos saem nas duas páginas e a contagem soma as unidades das duas.

- [ ] **Step 4: Atualizar o CLAUDE.md**

Acrescentar ao "Histórico de decisões", depois da entrada da v6.9.0:

```markdown
- **v6.10.0 — contagem e cobrança dos Protocolos dos Correios (aba 3)**: aba
  nova que conta as unidades entregues em cada protocolo, multiplica pela
  tarifa informada no lote (perguntada a cada processamento — ela muda com o
  tempo: nos protocolos de referência aparecem 3,45 e 3,85), carimba o valor
  no canto superior direito e gera uma planilha com o total. Ao contrário da
  aba 2, **aqui o OCR é usado** — o que torna isso aceitável é o
  `Listando N unidades` impresso no documento, que funciona como dígito
  verificador da contagem, com a contagem de linhas e a de "Correio" servindo
  de conferência. Divergiu, relê em qualidade maior e, persistindo, vira
  pendente sem carimbo. Abas de Cadastro e Logs renumeradas para 4 e 5. Ver
  spec `docs/superpowers/specs/2026-08-17-contagem-protocolo-correio-design.md`.
```

Acrescentar uma seção "Contagem dos Protocolos dos Correios (aba 3)" depois da seção da aba 2, registrando: o formato de saída do `winocr` (linha única, colunas fora de ordem), o traço duplicado que motivou o `(?:\s*[-–—])+`, a regra de aceite (basta um conferidor bater; sem `Listando` não aceita), e o lote de aceitação com os valores 7,70 / 57,75 / 11,55 / 46,20.

Atualizar também, na seção "O que o programa faz", a numeração das abas de Cadastro e Logs.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md
git commit -m "Documenta a aba de contagem dos protocolos dos Correios (v6.10.0)"
```

---

## Notas de execução

- `scratch/` não vai para o repositório. Conferir o `.gitignore` antes do primeiro commit da Task 7; se `scratch/` não estiver lá, usar a pasta de scratchpad da sessão.
- Os PDFs reais de `TESTE CORREIO` **não** entram no repositório nem viram fixture — as fixtures são sintéticas e sem nome de morador, seguindo o que já é feito em `tests/test_protocolo_correio.py`.
- O protocolo de duas páginas mencionado pelo usuário não existe no lote de referência; o suporte está implementado (`max_paginas=None`, carimbo em todas as páginas) mas só é verificado no passo 3 da Task 7, com um arquivo que o usuário precisa fornecer.
