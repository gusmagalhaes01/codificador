# Painel de resultado da aba 3 — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dar à aba 3 (Protocolos dos Correios) um painel de encerramento que lista os protocolos pendentes e deixa resolvê-los ali mesmo, digitando o valor em reais — que então é carimbado no PDF e entra na cobrança.

**Architecture:** Três helpers de painel saem do `mostrar_resultado` da aba 1 e passam a servir as duas abas. A aba 3 ganha `mostrar_resultado_protocolos`, um contexto guardado do último processamento e uma ação "Informar valor" que carimba o PDF e regrava a planilha. Em `logica.py` entram a leitura do valor digitado e o suporte a valor manual na linha da planilha.

**Tech Stack:** Python 3, CustomTkinter, ttk.Treeview, pypdf, reportlab, openpyxl, unittest.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-08-18-painel-resultado-protocolos-design.md`.
- Branch de trabalho: `painel-resultado-protocolos` (já criada, spec já commitado).
- Comentários, nomes de variáveis e mensagens em **português**.
- Nenhum texto visível ao usuário pode dizer "OCR" ou "DPI" — usar "protocolos escaneados", "qualidade de leitura".
- Todo widget CustomTkinter usa `corner_radius=0`; cores só via `self.tema_atual[chave]`, nunca hex solto.
- Cobalto (`acento`) só no que **resolve** uma pendência: o botão "Informar valor" e o cartão/rótulo "Pendentes". Botões neutros (Abrir PDF, Cancelar) são contornados, com fundo transparente.
- Dinheiro em `Decimal` internamente; `float` só na escrita da planilha.
- Pendência nunca recebe 0 no lugar de célula vazia — 0 significaria "entregou zero unidades".
- O comportamento do painel da aba 1 tem que ficar idêntico após a refatoração.
- A suíte roda com: `python -m unittest discover -s tests -p "test_*.py"` (hoje 131 testes).
- Nunca executar `python identificacao_por_cnpj_6_0.py` num subagente — janela bloqueante.

---

## File Structure

| Arquivo | Responsabilidade |
|---|---|
| `logica.py` (modificar) | `converter_valor_digitado` (leitura/validação do valor em reais) e `valor_manual` em `linha_planilha_protocolo`. |
| `identificacao_por_cnpj_6_0.py` (modificar) | Três helpers de painel, `mostrar_resultado_protocolos`, `_ctx_protocolos`, ação "Informar valor". |
| `tests/test_valor_digitado.py` (criar) | Leitura do valor digitado. |
| `tests/test_planilha_protocolo.py` (modificar) | Valor manual na linha e no TOTAL. |
| `CLAUDE.md` (modificar) | Seção da aba 3 + histórico. |

---

### Task 1: Leitura do valor digitado (`logica.py`)

**Files:**
- Modify: `logica.py` (após `formatar_reais`)
- Modify: `identificacao_por_cnpj_6_0.py` (`_pedir_tarifa`, para passar a usar a função nova)
- Test: `tests/test_valor_digitado.py`

**Interfaces:**
- Produces: `converter_valor_digitado(texto) -> tuple[Decimal | None, str]` — devolve `(valor, "")` quando válido e `(None, mensagem)` quando não. A mensagem vai direto para a tela, então é frase em português, sem jargão.

- [ ] **Step 1: Write the failing test**

Criar `tests/test_valor_digitado.py`:

```python
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
        """Entradas gigantes não podem levantar InvalidOperation crua — o
        usuário precisa ver o aviso normal, não um popup de erro técnico."""
        for entrada in ("1" * 40, "1e30"):
            valor, mensagem = app.converter_valor_digitado(entrada)
            self.assertIsNone(valor, entrada)
            self.assertTrue(mensagem, entrada)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_valor_digitado -v`
Expected: FAIL — `AttributeError: module 'logica' has no attribute 'converter_valor_digitado'`

- [ ] **Step 3: Write minimal implementation**

Conferir que o topo de `logica.py` importa `InvalidOperation`; se não, ajustar
a linha existente para:

```python
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
```

Em `logica.py`, logo após `formatar_reais`:

```python
LIMITE_VALOR_DIGITADO = Decimal("1000000")  # teto de sanidade para um lote


def converter_valor_digitado(texto):
    """
    Lê um valor em reais digitado por uma pessoa e devolve `(valor, mensagem)`:
    `(Decimal, "")` quando válido, `(None, aviso)` quando não. O aviso vai
    direto para a tela, então é frase em português, sem jargão.

    Aceita "300,30", "300.30", "1.234,56" e "R$ 57,75". Recusa mais de duas
    casas decimais pelo mesmo motivo que a tarifa recusa: um valor com três
    casas produz carimbo e planilha que não fecham quando alguém confere no
    papel.
    """
    bruto = (texto or "").strip().replace("R$", "").replace(" ", "")
    if not bruto:
        return None, "Digite um número, como 300,30."

    #  Formato brasileiro: o ponto é separador de milhar e a vírgula é o
    #  decimal. Sem vírgula, o ponto é tratado como decimal ("300.30").
    if "," in bruto:
        bruto = bruto.replace(".", "").replace(",", ".")

    try:
        valor = Decimal(bruto)
    except (InvalidOperation, ValueError):
        return None, "Digite um número, como 300,30."

    if not valor.is_finite():
        return None, "Digite um número, como 300,30."
    if valor <= 0:
        return None, "O valor precisa ser maior que zero."
    if valor > LIMITE_VALOR_DIGITADO:
        return None, "Esse valor parece alto demais. Confira o que foi digitado."
    if -valor.as_tuple().exponent > 2:
        return None, "Use no máximo duas casas decimais, como 300,30."

    return valor, ""
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_valor_digitado -v`
Expected: PASS (10 testes)

- [ ] **Step 5: Refazer `_pedir_tarifa` em cima da função nova**

Em `identificacao_por_cnpj_6_0.py`, acrescentar `converter_valor_digitado` ao
`from logica import (...)` e substituir o corpo da função interna `confirmar()`
de `_pedir_tarifa` por:

```python
        def confirmar():
            valor, mensagem = converter_valor_digitado(entrada.get())
            if valor is None:
                aviso.configure(text=mensagem)
                return
            resultado["valor"] = valor
            janela.destroy()
```

O nome do label de aviso e o do dict de resultado são os que já existem na
função — conferir no arquivo antes de trocar. A validação da tarifa passa a ser
literalmente a mesma do valor informado, que é o que o spec pede, e o
tratamento de entrada absurda deixa de ser um caminho separado.

- [ ] **Step 6: Rodar a suíte inteira**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (141 testes)

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem saída (compila)

- [ ] **Step 7: Commit**

```bash
git add logica.py identificacao_por_cnpj_6_0.py tests/test_valor_digitado.py
git commit -m "Lê e valida o valor em reais digitado, compartilhado com a tarifa"
```

---

### Task 2: Valor manual na linha da planilha (`logica.py`)

**Files:**
- Modify: `logica.py` (`linha_planilha_protocolo`)
- Test: `tests/test_planilha_protocolo.py`

**Interfaces:**
- Consumes: `COLUNAS_PROTOCOLO`, `salvar_planilha_protocolo` (já existem).
- Produces: `linha_planilha_protocolo(nome_arquivo, dados, cadastro, tarifa=None, unidades=None, observacao="", valor_manual=None) -> list`

- [ ] **Step 1: Write the failing test**

Acrescentar a `tests/test_planilha_protocolo.py`, antes do bloco
`if __name__ == "__main__":`. O arquivo já importa `unittest`, `Decimal`,
`logica as app` e `CADASTRO_TESTE`, e já define `DADOS_CADASTRADO` e
`DADOS_NAO_CADASTRADO` no topo:

```python
class TestValorManual(unittest.TestCase):
    def test_valor_manual_preenche_valor_e_deixa_unidades_vazias(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("300.30"))
        self.assertEqual(linha[1], "KLOSTERS")
        self.assertIsNone(linha[3])          # Unidades
        self.assertIsNone(linha[4])          # Tarifa
        self.assertEqual(linha[5], 300.30)   # Valor
        self.assertEqual(linha[6], "Valor informado manualmente")

    def test_valor_manual_nunca_grava_zero(self):
        """0 significaria "entregou zero unidades" — mesma regra das
        pendências e das retenções federais na planilha das NFS-e."""
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("10.00"))
        self.assertNotEqual(linha[3], 0)
        self.assertNotEqual(linha[4], 0)

    def test_valor_manual_grava_como_numero_nao_texto(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("300.30"))
        self.assertIsInstance(linha[5], float)

    def test_codigo_nao_cadastrado_mantem_os_dois_avisos(self):
        linha = app.linha_planilha_protocolo(
            "p.pdf", DADOS_NAO_CADASTRADO, CADASTRO_TESTE,
            valor_manual=Decimal("11.55"))
        self.assertIn("Valor informado manualmente", linha[6])
        self.assertIn("Código não cadastrado", linha[6])

    def test_total_soma_manuais_junto_com_calculados(self):
        from openpyxl import load_workbook
        pasta = tempfile.mkdtemp()
        try:
            destino = os.path.join(pasta, "p.xlsx")
            linhas = [
                app.linha_planilha_protocolo("a.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                             Decimal("3.85"), 15),
                app.linha_planilha_protocolo("b.pdf", DADOS_CADASTRADO, CADASTRO_TESTE,
                                             valor_manual=Decimal("300.30")),
            ]
            app.salvar_planilha_protocolo(destino, linhas)
            sheet = load_workbook(destino).active
            ultima = sheet.max_row
            self.assertEqual(sheet.cell(row=ultima, column=1).value, "TOTAL")
            #  Só as unidades efetivamente contadas entram no total de unidades
            self.assertEqual(sheet.cell(row=ultima, column=4).value, 15)
            #  57,75 calculado + 300,30 informado
            self.assertEqual(sheet.cell(row=ultima, column=6).value, 358.05)
        finally:
            shutil.rmtree(pasta, ignore_errors=True)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_planilha_protocolo.TestValorManual -v`
Expected: FAIL — `TypeError: linha_planilha_protocolo() got an unexpected keyword argument 'valor_manual'`

- [ ] **Step 3: Write minimal implementation**

Em `logica.py`, trocar a assinatura de `linha_planilha_protocolo` por:

```python
def linha_planilha_protocolo(nome_arquivo, dados, cadastro, tarifa=None,
                             unidades=None, observacao="", valor_manual=None):
```

O trecho que resolve `codigo`, `registro`, `condominio` e a observação de
cadastro **continua exatamente como está**. Inserir o ramo novo logo antes do
`if unidades is None:`:

```python
    if valor_manual is not None:
        #  Valor digitado por uma pessoa no painel de resultado: Unidades e
        #  Tarifa ficam VAZIAS, porque não houve contagem nem multiplicação —
        #  0 ali significaria "entregou zero unidades". A observação é o único
        #  rastro de que o número não foi calculado pelo programa; o carimbo no
        #  PDF mostra só o valor.
        aviso = "Valor informado manualmente"
        observacao = f"{observacao} · {aviso}" if observacao else aviso
        return [nome_arquivo, condominio, codigo, None, None,
                float(valor_manual), observacao]

    if unidades is None:
```

Acrescentar o parâmetro novo à docstring da função.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (146 testes)

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_planilha_protocolo.py
git commit -m "Aceita valor informado à mão na linha da planilha de protocolos"
```

---

### Task 3: Extrair os três helpers de painel

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (`mostrar_resultado`, a partir da linha ~1615)
- Test: script de regressão de árvore de widgets (fora do repo)

**Interfaces:**
- Produces:
  - `_montar_painel_resultado(self, titulo, geometria, minimo) -> ctk.CTkToplevel`
  - `_montar_faixa_cartoes(self, parent, cartoes) -> ctk.CTkFrame` — `cartoes` é uma lista de `(valor: str, rotulo: str, destaque: bool)`
  - `_montar_tabela_resultado(self, parent, colunas, larguras, altura=6) -> ttk.Treeview` — `colunas` é uma lista de `(chave, titulo)`

**Esta task é refatoração pura: nenhum comportamento pode mudar.** O painel da
aba 1 é o fluxo principal em produção, e a prova disso é o diff vazio do
Step 4 — não a leitura do código.

- [ ] **Step 1: Capturar a árvore de widgets ANTES da mudança**

Criar `dump_painel.py` na pasta de scratch da sessão (não commitar):

```python
"""Dump da árvore de widgets do painel da aba 1, para comparar antes/depois."""
import sys
sys.path.insert(0, r"C:\Users\Dell\Documents\p")
import identificacao_por_cnpj_6_0 as ui

RESULTADO = {
    "total": 3,
    "tempo_segundos": 72,
    "processados": [
        {"arquivo": "a.pdf", "codigo": "10004", "origem": "Nome do arquivo"},
        {"arquivo": "b.pdf", "codigo": "10002", "origem": "CNPJ no conteúdo"},
    ],
    "pendentes": [
        {"arquivo": "c.pdf", "motivo": "CNPJ não cadastrado",
         "cnpj": "40338774000141", "nome_sugerido": "ARAUJO LIMA",
         "candidatos": [], "caminho": r"C:\x\c.pdf", "tipo": "nao_cadastrado"},
    ],
}


def dump(widget, profundidade=0):
    linhas = []
    for filho in widget.winfo_children():
        descricao = filho.__class__.__name__
        for chave in ("text", "width", "height"):
            try:
                valor = filho.cget(chave)
                if valor not in (None, "", 0):
                    descricao += f" {chave}={valor!r}"
            except Exception:
                pass
        linhas.append("  " * profundidade + descricao)
        linhas.extend(dump(filho, profundidade + 1))
    return linhas


app = ui.App()
app.withdraw()
app.mostrar_resultado(RESULTADO)
app.update_idletasks()
print("\n".join(dump(app._janela_resultado)))
app.destroy()
```

Run: `python <scratch>/dump_painel.py > <scratch>/painel_antes.txt`
Expected: um arquivo com dezenas de linhas descrevendo a árvore, sem exceção.

Se o script não rodar, **pare e reporte**: sem a captura "antes" não existe a
prova que esta task exige, e seguir sem ela é refatorar às cegas código de
produção.

- [ ] **Step 2: Escrever os três helpers**

Inserir logo antes de `def mostrar_resultado` (linha ~1615):

```python
    def _montar_painel_resultado(self, titulo, geometria, minimo):
        """
        Cria (ou reaproveita, limpando o conteúdo) a janela de um painel de
        resultado. Compartilhado pelas abas 1 e 3 — as duas têm o mesmo
        formato de janela, só muda o que vai dentro.
        """
        existente = getattr(self, "_janela_resultado", None)
        if existente is not None and existente.winfo_exists():
            for widget in existente.winfo_children():
                widget.destroy()
            janela = existente
            janela.focus_force()
        else:
            janela = ctk.CTkToplevel(self)
            self._janela_resultado = janela

        janela.title(titulo)
        janela.geometry(geometria)
        janela.minsize(*minimo)
        janela.resizable(True, True)
        janela.configure(fg_color=self.tema_atual["fundo"])
        janela.transient(self)
        return janela

    def _montar_faixa_cartoes(self, parent, cartoes):
        """
        Linha de caixinhas de resumo no topo do painel: rótulo pequeno em cima,
        número grande embaixo, separados por hairlines verticais. `cartoes` é
        uma lista de (valor, rotulo, destaque); o destaque sai em cobalto,
        reservado ao que exige ação do usuário.
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        faixa = ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"])
        faixa.pack(fill="x", padx=24, pady=(24, 0))

        for indice, (valor, rotulo, destaque) in enumerate(cartoes):
            coluna = indice * 2
            faixa.columnconfigure(coluna, weight=1)

            cor_rotulo = tema["acento"] if destaque else tema["texto_secundario"]
            cor_valor = tema["acento"] if destaque else tema["texto"]

            bloco = ctk.CTkFrame(faixa, corner_radius=0, fg_color=tema["fundo"])
            bloco.grid(row=0, column=coluna, sticky="nsew", padx=16)
            ctk.CTkLabel(bloco, text=rotulo, font=(fonte, 11),
                         text_color=cor_rotulo, anchor="w").pack(fill="x", anchor="w")
            ctk.CTkLabel(bloco, text=valor, font=(fonte, 40),
                         text_color=cor_valor, anchor="w").pack(fill="x", anchor="w")

            if indice < len(cartoes) - 1:
                linha = ctk.CTkFrame(faixa, width=1, corner_radius=0,
                                     fg_color=tema["borda"])
                linha.grid(row=0, column=coluna + 1, sticky="ns")

        return faixa

    def _montar_tabela_resultado(self, parent, colunas, larguras, altura=6):
        """
        Treeview estilizada com a aparência do projeto, com barra de rolagem.
        `colunas` é uma lista de (chave, titulo); `larguras` casa por posição.
        Devolve a tabela; quem chama insere as linhas.
        """
        frame = ctk.CTkFrame(parent, corner_radius=0, fg_color=self.tema_atual["fundo"])
        frame.pack(fill="both", expand=False, pady=(0, 8))

        chaves = [c[0] for c in colunas]
        tabela = ttk.Treeview(frame, columns=chaves, show="headings", height=altura)
        for (chave, titulo), largura in zip(colunas, larguras):
            tabela.heading(chave, text=titulo)
            tabela.column(chave, width=largura, anchor="w")
        tabela.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
        tabela.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")
        return tabela
```

- [ ] **Step 3: Fazer a aba 1 usar os helpers**

Em `mostrar_resultado`, três substituições, sem tocar em mais nada:

1. O bloco que cria ou reaproveita a janela — de `if getattr(self, "_janela_resultado", None) is not None...` até `janela.transient(self)` — vira:

```python
        janela = self._montar_painel_resultado(
            "Resultado do processamento", "820x680", (640, 480))
```

2. Todo o bloco `# ===== FAIXA DE CARTÕES =====`, incluindo as funções internas `montar_cartao` e `hairline_vertical` e as cinco chamadas a elas, vira:

```python
        self._montar_faixa_cartoes(janela, [
            (str(len(processados)), "PROCESSADOS", False),
            (str(len(pendentes)), "PENDENTES", True),
            (tempo_str, "TEMPO", False),
        ])
```

3. Cada construção de `ttk.Treeview` com seu frame e scrollbar (a dos pendentes e a dos processados) vira uma chamada a `_montar_tabela_resultado`, com **exatamente** as mesmas colunas, títulos, larguras e `height` que estão hoje no arquivo. Conferir os valores no código antes de trocar; a tabela de pendentes hoje é `("arquivo", "motivo")` com larguras 280 e 420 e `height=6`.

Não mudar textos, ordem dos widgets, paddings, nem a lógica das ações.

- [ ] **Step 4: Provar que nada mudou**

Run: `python <scratch>/dump_painel.py > <scratch>/painel_depois.txt`
Run: `diff <scratch>/painel_antes.txt <scratch>/painel_depois.txt`
Expected: **saída vazia**. Qualquer diferença significa que a refatoração mudou
o painel — corrigir até o diff zerar e relatar o que tinha causado.

- [ ] **Step 5: Rodar a suíte inteira**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (146 testes)

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem saída

- [ ] **Step 6: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Extrai os helpers de painel usados pelo resultado da aba 1"
```

---

### Task 4: Painel de resultado da aba 3

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (`_processar_protocolos_em_thread` e métodos novos)

**Interfaces:**
- Consumes: `_montar_painel_resultado`, `_montar_faixa_cartoes`, `_montar_tabela_resultado` (Task 3); `formatar_reais` (já existe em `logica.py`).
- Produces:
  - `self._ctx_protocolos` — dict com `pasta_saida`, `config`, `destino`, `codigos`, `linhas`
  - `mostrar_resultado_protocolos(self, resultado)` — `resultado` tem `total` (int), `total_valor` (`Decimal`), `processados` (lista de dicts com `arquivo`, `condominio`, `unidades`, `valor`), `pendentes` (lista de dicts com `arquivo`, `condominio`, `codigo`, `motivo`, `caminho`, `indice_linha`) e `ignorados` (lista de dicts com `arquivo`, `motivo`)
  - `self._resultado_protocolos` — o dict renderizado, atualizado in-place pelas ações
  - `_protocolo_selecionado(self, tabela) -> dict | None`
  - `_acao_abrir_pdf_protocolo(self, dados)`

- [ ] **Step 1: Montar as listas do painel no laço de processamento**

Em `_processar_protocolos_em_thread`, antes do laço, junto dos contadores que já
existem, inicializar:

```python
        processados_painel = []
        pendentes_painel = []
        ignorados_painel = []
```

Dentro do laço, depois que `linhas.append(...)` já rodou (para que
`len(linhas) - 1` seja o índice da linha recém-criada), acrescentar:

```python
            registro_cadastro = self.cadastro.get(
                codigos.get((dados or {}).get("codigo") or "", "")) or None
            if registro_cadastro:
                condominio_painel = f"{dados['codigo']} {registro_cadastro['nome']}"
            elif dados:
                condominio_painel = dados.get("condominio", "")
            else:
                condominio_painel = ""

            registro_painel = {
                "arquivo": nome,
                "condominio": condominio_painel,
                "codigo": (dados or {}).get("codigo"),
                "caminho": caminho,
                "indice_linha": len(linhas) - 1,
            }
            if unidades is not None:
                registro_painel["unidades"] = unidades
                registro_painel["valor"] = float(valor_protocolo(unidades, tarifa))
                processados_painel.append(registro_painel)
            elif nao_e_protocolo:
                ignorados_painel.append({"arquivo": nome, "motivo": observacao})
            else:
                registro_painel["motivo"] = observacao
                pendentes_painel.append(registro_painel)
```

- [ ] **Step 2: Guardar o contexto e abrir o painel**

No fim de `_processar_protocolos_em_thread`, substituir a linha
`self.after(0, lambda: self._concluir_extracao(destino, resumo))` por:

```python
        self._ctx_protocolos = {
            "pasta_saida": pasta_saida,
            "config": config,
            "destino": destino,
            "codigos": codigos,
            "linhas": linhas,
        }
        resultado = {
            "total": total,
            "total_valor": total_valor,
            "processados": processados_painel,
            "pendentes": pendentes_painel,
            "ignorados": ignorados_painel,
        }
        self.after(0, lambda: self.mostrar_resultado_protocolos(resultado))
```

`total_valor` já é calculado logo acima, para o resumo. Manter o `resumo` na
label de status como está — o painel não substitui essa informação, soma a ela.

- [ ] **Step 3: Escrever o painel**

Inserir após `_carimbar_protocolo`:

```python
    def mostrar_resultado_protocolos(self, resultado):
        """
        Painel de encerramento da aba 3. Mesma linguagem do painel da aba 1:
        cartões no topo, pendentes primeiro, ações por linha. A diferença é o
        que resolve uma pendência aqui — informar o valor em reais.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        self._resultado_protocolos = resultado

        processados = resultado.get("processados", []) or []
        pendentes = resultado.get("pendentes", []) or []
        ignorados = resultado.get("ignorados", []) or []
        total_valor = resultado.get("total_valor", Decimal("0.00"))

        janela = self._montar_painel_resultado(
            "Resultado dos protocolos", "860x680", (680, 480))

        self._montar_faixa_cartoes(janela, [
            (str(len(processados)), "PROTOCOLOS", False),
            (str(len(pendentes)), "PENDENTES", True),
            (formatar_reais(total_valor), "TOTAL", False),
        ])

        ctk.CTkFrame(janela, height=1, corner_radius=0,
                     fg_color=tema["borda"]).pack(fill="x", padx=24, pady=(24, 0))

        ctk.CTkLabel(
            janela, text=f"{resultado.get('total', 0)} arquivo(s) no total.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w",
        ).pack(side="bottom", fill="x", padx=24, pady=16)

        area = ctk.CTkScrollableFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        area.pack(side="top", fill="both", expand=True, padx=24, pady=(16, 0))

        # --- PENDENTES ---
        ctk.CTkLabel(
            area, text="PENDENTES — PRECISAM DE AÇÃO", font=(fonte, 11, "bold"),
            text_color=tema["acento"], anchor="w",
        ).pack(fill="x", pady=(0, 8))

        self._pend_protocolo_por_iid = {}
        if not pendentes:
            ctk.CTkLabel(
                area, text="Nenhum pendente — todos os protocolos foram calculados.",
                font=(fonte, 13), text_color=tema["texto_secundario"], anchor="w",
            ).pack(fill="x", pady=(0, 16))
        else:
            tabela = self._montar_tabela_resultado(
                area,
                [("arquivo", "Arquivo"), ("condominio", "Condomínio"), ("motivo", "Motivo")],
                [240, 200, 320],
            )
            for i, dados in enumerate(pendentes):
                iid = f"prot{i}"
                self._pend_protocolo_por_iid[iid] = dados
                tabela.insert("", "end", iid=iid, values=(
                    dados.get("arquivo", ""), dados.get("condominio", ""),
                    dados.get("motivo", "")))

            acoes = ctk.CTkFrame(area, corner_radius=0, fg_color=tema["fundo"])
            acoes.pack(fill="x", pady=(0, 16))

            botao_valor = ctk.CTkButton(
                acoes, text="Informar valor", corner_radius=0, state="disabled",
                fg_color=tema["borda"], hover_color=tema["acento_hover"],
                text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
                command=lambda: self._acao_informar_valor(
                    self._protocolo_selecionado(tabela)),
            )
            botao_valor.pack(side="left", padx=(0, 8))

            botao_abrir = ctk.CTkButton(
                acoes, text="Abrir PDF", corner_radius=0, state="disabled",
                fg_color="transparent", hover_color=tema["superficie"],
                border_width=1, border_color=tema["borda_forte"],
                text_color=tema["texto"], font=(fonte, 13),
                command=lambda: self._acao_abrir_pdf_protocolo(
                    self._protocolo_selecionado(tabela)),
            )
            botao_abrir.pack(side="left", padx=(0, 8))

            def ao_selecionar(_evento=None):
                #  Botão cobalto desabilitado tem que apagar o fg_color, senão
                #  fica azul e parece clicável.
                tem_linha = bool(tabela.selection())
                botao_valor.configure(
                    state="normal" if tem_linha else "disabled",
                    fg_color=tema["acento"] if tem_linha else tema["borda"])
                botao_abrir.configure(state="normal" if tem_linha else "disabled")

            tabela.bind("<<TreeviewSelect>>", ao_selecionar)
            tabela.bind("<Double-1>", lambda _e: self._acao_informar_valor(
                self._protocolo_selecionado(tabela)))

        # --- CALCULADOS ---
        ctk.CTkLabel(
            area, text="CALCULADOS", font=(fonte, 11, "bold"),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", pady=(8, 8))

        if processados:
            tabela_ok = self._montar_tabela_resultado(
                area,
                [("arquivo", "Arquivo"), ("condominio", "Condomínio"),
                 ("unidades", "Unidades"), ("valor", "Valor")],
                [240, 200, 90, 120],
            )
            for dados in processados:
                unidades = dados.get("unidades")
                tabela_ok.insert("", "end", values=(
                    dados.get("arquivo", ""), dados.get("condominio", ""),
                    "" if unidades is None else unidades,
                    formatar_reais(dados.get("valor", 0))))
        else:
            ctk.CTkLabel(
                area, text="Nenhum protocolo calculado neste lote.",
                font=(fonte, 13), text_color=tema["texto_secundario"], anchor="w",
            ).pack(fill="x", pady=(0, 16))

        # --- IGNORADOS (sem ação: não são protocolos) ---
        if ignorados:
            ctk.CTkLabel(
                area, text="IGNORADOS — NÃO SÃO PROTOCOLOS", font=(fonte, 11, "bold"),
                text_color=tema["texto_secundario"], anchor="w",
            ).pack(fill="x", pady=(16, 8))
            tabela_ign = self._montar_tabela_resultado(
                area, [("arquivo", "Arquivo"), ("motivo", "Motivo")],
                [300, 360], altura=3)
            for dados in ignorados:
                tabela_ign.insert("", "end", values=(
                    dados.get("arquivo", ""), dados.get("motivo", "")))

    def _protocolo_selecionado(self, tabela):
        """Dict do pendente na linha selecionada, ou None."""
        selecao = tabela.selection()
        if not selecao:
            return None
        return self._pend_protocolo_por_iid.get(selecao[0])

    def _acao_abrir_pdf_protocolo(self, dados):
        if not dados:
            return
        try:
            os.startfile(dados["caminho"])
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Não foi possível abrir o PDF:\n{e}",
                parent=self._janela_resultado)

    def _acao_informar_valor(self, dados):
        #  Implementado na tarefa seguinte; o painel já liga o botão a ele para
        #  que esta tarefa feche com a tela navegável.
        if not dados:
            return
        messagebox.showinfo(
            "Em construção", "Informar valor chega na próxima etapa.",
            parent=self._janela_resultado)
```

- [ ] **Step 4: Verificar**

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem saída

Run: `python -c "import identificacao_por_cnpj_6_0"`
Expected: sem erro

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (146 testes)

**Não abrir o app** — janela bloqueante. A verificação visual é da Task 6.

- [ ] **Step 5: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Painel de resultado da aba 3 com pendentes, calculados e ignorados"
```

---

### Task 5: Informar o valor e resolver a pendência

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (`_carimbar_protocolo`, `_acao_informar_valor`, método novo `_pedir_valor_protocolo`)

**Interfaces:**
- Consumes: `converter_valor_digitado` (Task 1); `linha_planilha_protocolo(..., valor_manual=)` (Task 2); `self._ctx_protocolos`, `self._resultado_protocolos`, `mostrar_resultado_protocolos` (Task 4).
- Produces: `_carimbar_protocolo(..., valor_manual=None)` — quando informado, carimba só o valor, sem a conta.

- [ ] **Step 1: Permitir carimbo só com o valor**

Em `_carimbar_protocolo`, trocar a assinatura por:

```python
    def _carimbar_protocolo(self, caminho, pasta_saida, nome, dados, unidades,
                            tarifa, config, codigos, valor_manual=None):
```

E substituir as duas linhas que hoje calculam `valor` e `texto_valor` por:

```python
        if valor_manual is not None:
            #  Valor digitado por uma pessoa: não há conta a mostrar, então o
            #  carimbo traz só o número — que é o que se escreveria à caneta.
            valor = valor_manual
            texto_valor = formatar_reais(valor)
        else:
            valor = valor_protocolo(unidades, tarifa)
            texto_valor = montar_texto_valor_protocolo(unidades, tarifa, valor)
```

O resto da função — carimbo lateral, montagem dos extras, chamada a
`processar_pdf` — fica igual.

- [ ] **Step 2: Escrever a janelinha do valor**

Inserir junto das outras ações do painel:

```python
    def _pedir_valor_protocolo(self, dados):
        """
        Janelinha que pede o valor em reais de um protocolo pendente. Devolve
        Decimal, ou None se o usuário cancelar. Mesma validação da tarifa,
        pela mesma função.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        janela = ctk.CTkToplevel(self._janela_resultado)
        janela.title("Informar valor")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(self._janela_resultado)
        janela.grab_set()

        escolha = {"valor": None}

        ctk.CTkLabel(janela, text="Qual o valor deste protocolo?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text=dados.get("arquivo", ""), font=(fonte, 12),
                     text_color=tema["texto_secundario"]).pack(padx=24, anchor="w")
        ctk.CTkLabel(janela, text=dados.get("condominio", ""), font=(fonte, 12),
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
            valor, mensagem = converter_valor_digitado(entrada.get())
            if valor is None:
                aviso.configure(text=mensagem)
                return
            escolha["valor"] = valor
            janela.destroy()

        botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        botoes.pack(padx=24, pady=(0, 24), anchor="e")
        ctk.CTkButton(botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(botoes, text="Confirmar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return escolha["valor"]
```

- [ ] **Step 3: Escrever a ação**

Substituir o esqueleto de `_acao_informar_valor` por:

```python
    def _acao_informar_valor(self, dados):
        """
        Pede o valor em reais e resolve a pendência: carimba o PDF, move a
        linha para os calculados e regrava a planilha.
        """
        if not dados:
            return
        ctx = getattr(self, "_ctx_protocolos", None)
        if not ctx:
            messagebox.showerror(
                "Erro",
                "O contexto do processamento se perdeu. Rode o lote de novo.",
                parent=self._janela_resultado)
            return

        valor = self._pedir_valor_protocolo(dados)
        if valor is None:
            return

        registro_dados = {"codigo": dados.get("codigo"),
                          "condominio": dados.get("condominio", "")}

        #  1. Carimba o PDF: código na lateral (se cadastrado) e valor no topo
        #     direito. Falhou aqui, nada mais acontece — não faz sentido marcar
        #     como resolvido um arquivo que não saiu carimbado.
        try:
            self._carimbar_protocolo(
                dados["caminho"], ctx["pasta_saida"], dados["arquivo"],
                registro_dados, None, None, ctx["config"], ctx["codigos"],
                valor_manual=valor)
        except Exception as e:
            messagebox.showerror(
                "Erro ao carimbar", f"Não foi possível carimbar o PDF:\n{e}",
                parent=self._janela_resultado)
            return

        #  2. Atualiza a linha da planilha em memória
        ctx["linhas"][dados["indice_linha"]] = linha_planilha_protocolo(
            dados["arquivo"], registro_dados, self.cadastro, valor_manual=valor)

        #  3. Move do painel de pendentes para os calculados
        resultado = self._resultado_protocolos
        resultado["pendentes"] = [p for p in resultado["pendentes"] if p is not dados]
        resultado["processados"].append({
            "arquivo": dados["arquivo"],
            "condominio": dados.get("condominio", ""),
            "unidades": None,
            "valor": float(valor),
        })
        resultado["total_valor"] = resultado.get("total_valor", Decimal("0.00")) + valor

        #  4. Regrava a planilha inteira. Se falhar (tipicamente porque está
        #     aberta no Excel), o valor NÃO se perde: já está em ctx["linhas"]
        #     e no resultado, e vai junto na próxima regravação.
        try:
            salvar_planilha_protocolo(ctx["destino"], ctx["linhas"])
        except Exception as e:
            messagebox.showwarning(
                "Planilha não atualizada",
                f"O valor foi carimbado no PDF, mas a planilha não pôde ser "
                f"regravada:\n{e}\n\nSe ela estiver aberta no Excel, feche e "
                f"informe o próximo valor — a planilha é regravada inteira a "
                f"cada correção.",
                parent=self._janela_resultado)

        self.mostrar_resultado_protocolos(resultado)
```

Conferir que `salvar_planilha_protocolo`, `linha_planilha_protocolo`,
`formatar_reais` e `converter_valor_digitado` estão no `from logica import (...)`.

- [ ] **Step 4: Verificar**

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem saída

Run: `python -c "import identificacao_por_cnpj_6_0"`
Expected: sem erro

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: PASS (146 testes)

- [ ] **Step 5: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Resolve pendência do protocolo informando o valor à mão"
```

---

### Task 6: Verificação com os protocolos reais e documentação

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Confirmar que o lote real gera a pendência esperada**

A pasta `C:\Users\Dell\Downloads\TESTE CORREIO\Nova pasta` tem quatro protocolos
multipágina. O do `11049 APART HOTEL` traz o `Listando 76 unidades` impresso
riscado a caneta, com `78` escrito à mão embaixo — ele fica pendente de
propósito, e é o caso que motivou este recurso.

Criar `pendencia.py` na pasta de scratch da sessão (não commitar):

```python
"""Reproduz a classificação da aba 3 sobre o lote multipágina, sem interface."""
import sys, glob, os
from decimal import Decimal
sys.path.insert(0, r"C:\Users\Dell\Documents\p")
import logica as app
import identificacao_por_cnpj_6_0 as ui

TARIFA = Decimal("3.85")
cad = app.carregar_cadastro("cadastro_condominios.xlsx")
codigos = app._codigos_do_cadastro(cad)

calculados, pendentes, total = [], [], Decimal("0.00")
for caminho in sorted(glob.glob(
        r"C:\Users\Dell\Downloads\TESTE CORREIO\Nova pasta\*.pdf")):
    nome = os.path.basename(caminho)
    texto = app.extrair_texto_pdf(caminho)
    if len(texto.strip()) < app.LIMITE_TEXTO_MINIMO:
        texto = app.extrair_texto_escaneado(
            caminho, dpi=ui.QUALIDADE_LEITURA_PROTOCOLO)
    dados = app.extrair_dados_protocolo_correio(texto)
    if dados is None:
        print(f"{nome} | NAO E PROTOCOLO")
        continue
    aceito, motivo = app.conferir_contagem_protocolo(dados)
    if aceito:
        valor = app.valor_protocolo(dados["total_impresso"], TARIFA)
        total += valor
        calculados.append(nome)
        print(f"{nome} | cod {dados['codigo']} | "
              f"{dados['total_impresso']} un | R$ {valor}")
    else:
        pendentes.append(nome)
        print(f"{nome} | cod {dados['codigo']} | PENDENTE: {motivo}")

print(f"\n{len(calculados)} calculado(s), {len(pendentes)} pendente(s), "
      f"total R$ {total}")
sys.exit(0 if len(calculados) == 3 and len(pendentes) == 1
         and total == Decimal("488.95") else 1)
```

Run: `python <scratch>/pendencia.py`
Expected: três calculados — 36, 31 e 60 unidades, somando **R$ 488,95** à
tarifa de R$ 3,85 — e um pendente, o do 11049, com motivo de contagem não lida.

- [ ] **Step 2: Verificação visual pelo usuário**

Esta parte não é automatizável e não deve ser tentada por um subagente. Pedir
ao usuário que abra o app, processe essa pasta com tarifa `3,85` e confirme:

- o painel abre com os cartões Protocolos / Pendentes / Total;
- o protocolo do 11049 aparece em Pendentes, com motivo legível;
- "Informar valor" com `300,30` carimba o PDF, move a linha para Calculados e
  o cartão Total sobe para R$ 789,25;
- o PDF de saída tem o código na lateral e `R$ 300,30` no topo direito, sem a
  conta;
- a planilha tem essa linha com Unidades vazia, Valor 300,30 e observação
  `Valor informado manualmente`, e o TOTAL inclui esse valor;
- com a planilha aberta no Excel, informar um valor mostra o aviso e o número
  não se perde — fechando o Excel e informando o próximo, os dois entram.

- [ ] **Step 3: Atualizar o CLAUDE.md**

Na seção "Contagem dos Protocolos dos Correios (aba 3)", acrescentar um bloco
sobre o painel de resultado, registrando:

- que a aba passou a encerrar com painel, no molde da aba 1, e que os três
  helpers de painel são compartilhados;
- que a pendência se resolve informando o **valor em reais**, não a quantidade,
  porque é o que a pessoa já escreve à caneta na folha;
- que a linha resolvida sai com Unidades vazia e observação
  `Valor informado manualmente`, e que o carimbo mostra só o valor — logo, o
  papel sozinho não distingue valor calculado de valor digitado;
- o caso que motivou: no protocolo real do 11049, o `Listando 76 unidades`
  impresso foi riscado à mão e trocado por 78. A correção está fora do que a
  máquina lê, e a contagem de "Correio" concordava com o número riscado — se o
  risco estivesse mais leve, o programa teria cobrado 76 com os dois
  conferidores de acordo. Não há conferidor automático que resolva isso.

Acrescentar a entrada correspondente ao "Histórico de decisões", com o número
de versão que o usuário definir.

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md
git commit -m "Documenta o painel de resultado da aba 3"
```
