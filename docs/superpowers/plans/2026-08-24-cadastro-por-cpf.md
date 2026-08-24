# Condomínios identificados por CPF — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Permitir cadastrar e identificar condomínios cujo documento é um CPF (11 dígitos) em vez de um CNPJ (14), com o CPF sendo extraído dos boletos e notas como o CNPJ já é.

**Architecture:** A chave do cadastro passa de "CNPJ" para "documento normalizado" — 11 ou 14 dígitos, distinguidos pelo comprimento, cada um com seu próprio dígito verificador. A extração aceita CPF sempre que ele vier depois de um rótulo (`PAGADOR`, `TOMADOR`...) e, na varredura genérica sem rótulo, só quando o CPF já estiver no cadastro.

**Tech Stack:** Python 3, `unittest`, `openpyxl`, `pypdf`, `reportlab`. Sem dependência nova.

## Global Constraints

- Comentários, nomes de função e nomes de variável em **português**, seguindo o resto do código.
- Nenhuma dependência nova em `requirements.txt`.
- Documento sempre normalizado para só dígitos internamente (`normalizar_cnpj`), formatado apenas na exibição e na planilha.
- Toda a lógica em `logica.py`, sem nenhum import de interface. `identificacao_por_cnpj_6_0.py` só consome.
- Vocabulário visível ao usuário é para leigos: nunca expor "OCR" ou "DPI" em texto de interface.
- Suíte completa: `python -m unittest discover -s tests -p "test_*.py"`. Ela precisa passar ao fim de **toda** tarefa, não só da última.
- Um commit por tarefa, mensagem em português no imperativo, como as do histórico (`Permite dizer o condomínio de um protocolo não identificado`).

## Estrutura de arquivos

| Arquivo | Responsabilidade nesta mudança |
|---|---|
| `logica.py` | `cpf_valido`, `formatar_documento`, `CPF_REGEX`/`CPF_FLEX`, extração de CPF, carimbo, planilha do cadastro |
| `identificacao_por_cnpj_6_0.py` | Formulário da aba de Cadastro, tabela, chamadas de `extrair_cnpj_tomador`, rótulos de tela |
| `tests/cadastro_teste.py` | Ganha um condomínio identificado por CPF |
| `tests/test_cpf_valido.py` | **Criar** — validação de CPF |
| `tests/test_formatar_documento.py` | **Criar** — formatação por comprimento |
| `tests/test_cadastro_cpf.py` | **Criar** — ida e volta pela planilha |
| `tests/test_extracao_cnpj.py` | Ganha os casos de CPF (campo rotulado e fallback) |
| `tests/test_carimbo_lateral.py` | Ganha o caso de CPF omitido no carimbo |
| `tests/test_extracao_nfse.py` | Ganha o caso de tomador pessoa física |
| `CLAUDE.md` | Seção nova documentando a mudança |

**CPFs usados nos testes** (fixos, deriváveis pelo algoritmo — não são de pessoas reais):
- Válido: `52998224725`
- DV errado: `52998224724`
- Todos iguais: `11111111111`

---

### Task 1: `cpf_valido`

**Files:**
- Modify: `logica.py` (logo depois de `cnpj_valido`, que termina em `logica.py:371`)
- Test: `tests/test_cpf_valido.py` (criar)

**Interfaces:**
- Consumes: nada.
- Produces: `cpf_valido(cpf_normalizado: str) -> bool` — recebe só dígitos, devolve `True` se os dois DVs baterem.

- [ ] **Step 1: Escrever o teste que falha**

Criar `tests/test_cpf_valido.py`:

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


class TestCpfValido(unittest.TestCase):
    def test_cpf_valido(self):
        self.assertTrue(app.cpf_valido("52998224725"))

    def test_digito_verificador_errado(self):
        self.assertFalse(app.cpf_valido("52998224724"))

    def test_todos_digitos_iguais_e_invalido(self):
        # Passa no cálculo dos DVs, mas não é CPF de ninguém — mesma
        # armadilha que cnpj_valido já cobre para os 14 dígitos iguais.
        self.assertFalse(app.cpf_valido("11111111111"))

    def test_comprimento_errado_e_invalido(self):
        self.assertFalse(app.cpf_valido("529982247"))
        self.assertFalse(app.cpf_valido(""))

    def test_cnpj_nao_passa_por_cpf(self):
        self.assertFalse(app.cpf_valido("01195716000154"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_cpf_valido -v
```

Esperado: FAIL com `AttributeError: module 'logica' has no attribute 'cpf_valido'`.

- [ ] **Step 3: Implementar**

Em `logica.py`, imediatamente depois do `return d[12] == dv1 and d[13] == dv2` de `cnpj_valido`:

```python
def cpf_valido(cpf_normalizado):
    """
    Valida os dígitos verificadores de um CPF (11 dígitos). Existe pelo mesmo
    motivo de `cnpj_valido`: descartar leituras de OCR que "parecem" um
    documento mas têm algum caractere errado.

    O algoritmo é diferente do CNPJ (pesos decrescentes 10..2 e 11..2), mas o
    contrato é o mesmo — recebe só dígitos, devolve bool.
    """
    d = cpf_normalizado
    if len(d) != 11 or d == d[0] * 11:
        return False

    def _digito(nums, peso_inicial):
        soma = sum(int(n) * p for n, p in zip(nums, range(peso_inicial, 1, -1)))
        resto = soma % 11
        return "0" if resto < 2 else str(11 - resto)

    dv1 = _digito(d[:9], 10)
    dv2 = _digito(d[:9] + dv1, 11)
    return d[9] == dv1 and d[10] == dv2
```

- [ ] **Step 4: Rodar o teste e confirmar que passa**

```bash
python -m unittest tests.test_cpf_valido -v
```

Esperado: `OK`, 5 testes.

- [ ] **Step 5: Rodar a suíte inteira**

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK`.

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_cpf_valido.py
git commit -m "Valida o dígito verificador do CPF"
```

---

### Task 2: `formatar_documento` no lugar de `formatar_cnpj`

**Files:**
- Modify: `logica.py:337-343` (a função), `logica.py:700`, `logica.py:1204`, `logica.py:1428`
- Modify: `identificacao_por_cnpj_6_0.py:68` (import), `:450`, `:461`, `:1912`, `:1960`, `:4125`, `:4137`
- Test: `tests/test_formatar_documento.py` (criar)

**Interfaces:**
- Consumes: nada.
- Produces: `formatar_documento(documento_normalizado: str) -> str` — 11 dígitos saem `000.000.000-00`, 14 saem `00.000.000/0000-00`, qualquer outra coisa sai crua. **`formatar_cnpj` deixa de existir**; nenhum código novo deve chamá-la.

- [ ] **Step 1: Escrever o teste que falha**

Criar `tests/test_formatar_documento.py`:

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


class TestFormatarDocumento(unittest.TestCase):
    def test_catorze_digitos_saem_como_cnpj(self):
        self.assertEqual(app.formatar_documento("01195716000154"),
                         "01.195.716/0001-54")

    def test_onze_digitos_saem_como_cpf(self):
        self.assertEqual(app.formatar_documento("52998224725"),
                         "529.982.247-25")

    def test_comprimento_desconhecido_sai_cru(self):
        # Mesmo contrato que formatar_cnpj tinha: não inventa formatação
        # para o que não reconhece.
        self.assertEqual(app.formatar_documento("123"), "123")
        self.assertEqual(app.formatar_documento(""), "")

    def test_aceita_entrada_ja_formatada(self):
        self.assertEqual(app.formatar_documento("529.982.247-25"),
                         "529.982.247-25")

    def test_formatar_cnpj_nao_existe_mais(self):
        self.assertFalse(hasattr(app, "formatar_cnpj"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_formatar_documento -v
```

Esperado: FAIL com `AttributeError: module 'logica' has no attribute 'formatar_documento'`.

- [ ] **Step 3: Substituir a função em `logica.py`**

Trocar o corpo atual de `formatar_cnpj` (`logica.py:337-343`) por:

```python
def formatar_documento(documento_normalizado):
    """
    Formata para exibição um documento de condomínio: 14 dígitos viram CNPJ,
    11 viram CPF. O comprimento é o que distingue os dois — nenhum CPF pode
    ser confundido com um CNPJ, nem o contrário.

    Qualquer outro comprimento sai cru, sem formatação inventada.
    """
    d = re.sub(r"\D", "", documento_normalizado)
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return documento_normalizado
```

- [ ] **Step 4: Atualizar todos os pontos de chamada**

```bash
grep -rn "formatar_cnpj" logica.py identificacao_por_cnpj_6_0.py tests/
```

Trocar **todas** as ocorrências por `formatar_documento`, incluindo a linha de
import em `identificacao_por_cnpj_6_0.py:68`. São 4 em `logica.py` (a definição
e as linhas 700, 1204, 1428) e 7 em `identificacao_por_cnpj_6_0.py` (68, 450,
461, 1912, 1960, 4125, 4137). É uma troca mecânica de nome — nenhum
comportamento muda para CNPJ.

Depois, confirmar que não sobrou nenhuma:

```bash
grep -rn "formatar_cnpj" logica.py identificacao_por_cnpj_6_0.py tests/
```

Esperado: nenhuma saída.

- [ ] **Step 5: Rodar o teste e a suíte**

```bash
python -m unittest tests.test_formatar_documento -v
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK` nos dois.

- [ ] **Step 6: Conferir que a interface ainda carrega**

```bash
python -c "import ast; ast.parse(open('identificacao_por_cnpj_6_0.py',encoding='utf-8').read()); print('sintaxe ok')"
```

Esperado: `sintaxe ok`. Nenhum teste cobre a interface, então essa checagem mais
o grep abaixo são o que pega um `formatar_cnpj` esquecido em
`identificacao_por_cnpj_6_0.py` que viraria `NameError` só na hora do clique.

```bash
grep -c "formatar_documento" identificacao_por_cnpj_6_0.py
```

Esperado: `7`.

- [ ] **Step 7: Commit**

```bash
git add logica.py identificacao_por_cnpj_6_0.py tests/test_formatar_documento.py
git commit -m "Formata CPF e CNPJ pela mesma função de documento"
```

---

### Task 3: O cadastro guarda CPF na planilha

**Files:**
- Modify: `logica.py:1192-1210` (`salvar_cadastro`)
- Modify: `tests/cadastro_teste.py`
- Modify: `tests/test_cnpj_valido.py:17-19`
- Test: `tests/test_cadastro_cpf.py` (criar)

**Interfaces:**
- Consumes: `formatar_documento` (Task 2).
- Produces: `CADASTRO_TESTE` passa a conter a chave `"52998224725"` → `{"codigo": "11300", "nome": "VILA MARINA", "id_sl": "930"}`. As tarefas 5 e 6 usam esse condomínio.

**Nota para quem implementa:** `carregar_cadastro` **não precisa mudar**. Ela já
faz `cnpj_norm = normalizar_cnpj(str(cnpj))` e aceita qualquer sequência de
dígitos não vazia como chave — um CPF de 11 dígitos já carrega hoje. O que
falta é a gravação: `salvar_cadastro` precisa formatar pelo comprimento e
escrever o cabeçalho novo. **Não acrescente validação de comprimento em
`carregar_cadastro`**: isso passaria a descartar em silêncio linhas que hoje
carregam, que é exatamente o defeito que esta mudança existe para corrigir.

- [ ] **Step 1: Escrever o teste que falha**

Criar `tests/test_cadastro_cpf.py`:

```python
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

import logica as app
from openpyxl import Workbook, load_workbook

CPF_VILA_MARINA = "52998224725"
KLOSTERS = "01195716000154"


class TestCadastroPorCpf(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        self.caminho = os.path.join(self.pasta, "cadastro.xlsx")

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_ida_e_volta_preserva_o_cpf(self):
        original = {
            CPF_VILA_MARINA: {"codigo": "11300", "nome": "VILA MARINA", "id_sl": "930"},
            KLOSTERS: {"codigo": "10004", "nome": "KLOSTERS", "id_sl": "44"},
        }
        app.salvar_cadastro(self.caminho, original)
        self.assertEqual(app.carregar_cadastro(self.caminho), original)

    def test_cpf_gravado_formatado_como_cpf(self):
        app.salvar_cadastro(self.caminho, {
            CPF_VILA_MARINA: {"codigo": "11300", "nome": "VILA MARINA", "id_sl": "930"},
        })
        sheet = load_workbook(self.caminho).active
        self.assertEqual(sheet.cell(row=2, column=1).value, "529.982.247-25")

    def test_cabecalho_menciona_os_dois_documentos(self):
        app.salvar_cadastro(self.caminho, {})
        sheet = load_workbook(self.caminho).active
        self.assertEqual(sheet.cell(row=1, column=1).value, "CNPJ / CPF")

    def test_planilha_antiga_so_com_cnpj_continua_carregando(self):
        """O cabeçalho mudou de nome, mas carregar_cadastro lê por posição —
        planilha gravada por uma versão anterior precisa continuar abrindo."""
        wb = Workbook()
        sheet = wb.active
        sheet.append(["CNPJ", "Código", "Nome do Condomínio", "ID SL"])
        sheet.append(["01.195.716/0001-54", "10004", "KLOSTERS", "44"])
        wb.save(self.caminho)
        cadastro = app.carregar_cadastro(self.caminho)
        self.assertEqual(cadastro[KLOSTERS]["nome"], "KLOSTERS")

    def test_cpf_e_cnpj_nunca_colidem_como_chave(self):
        cadastro = {
            CPF_VILA_MARINA: {"codigo": "11300", "nome": "VILA MARINA", "id_sl": "930"},
            KLOSTERS: {"codigo": "10004", "nome": "KLOSTERS", "id_sl": "44"},
        }
        app.salvar_cadastro(self.caminho, cadastro)
        self.assertEqual(len(app.carregar_cadastro(self.caminho)), 2)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_cadastro_cpf -v
```

Esperado: FAIL em `test_cabecalho_menciona_os_dois_documentos` (`'CNPJ' != 'CNPJ / CPF'`) e em `test_cpf_gravado_formatado_como_cpf` (o CPF sai cru, sem pontuação).

- [ ] **Step 3: Ajustar `salvar_cadastro`**

Em `logica.py`, dentro de `salvar_cadastro`, trocar a linha do cabeçalho e o corpo do laço:

```python
    sheet.append(["CNPJ / CPF", "Código", "Nome do Condomínio", "ID SL"])
    for documento, dados in sorted(cadastro.items(), key=lambda kv: kv[1]["nome"]):
        sheet.append([formatar_documento(documento), dados["codigo"], dados["nome"],
                       dados.get("id_sl", "")])
```

E acrescentar ao final da docstring da função:

```
    A coluna A guarda CNPJ (14 dígitos) ou CPF (11) — alguns condomínios não
    têm CNPJ próprio e são identificados pelo CPF do síndico, que é o que sai
    impresso no documento. `carregar_cadastro` lê por posição e nunca pelo nome
    do cabeçalho, então planilha gravada por versão anterior (cabeçalho "CNPJ")
    continua abrindo.
```

- [ ] **Step 4: Rodar o teste e confirmar que passa**

```bash
python -m unittest tests.test_cadastro_cpf -v
```

Esperado: `OK`, 5 testes.

- [ ] **Step 5: Acrescentar o condomínio por CPF ao cadastro de teste**

Em `tests/cadastro_teste.py`, acrescentar como última entrada do dict `CADASTRO_TESTE`:

```python
    "52998224725": {"codigo": "11300", "nome": "VILA MARINA", "id_sl": "930"},
```

E acrescentar ao final da docstring do módulo:

```
VILA MARINA está identificada por CPF (11 dígitos) de propósito: alguns
condomínios não têm CNPJ e são identificados pelo CPF do síndico, que é o que
sai impresso no boleto. O programa precisa lidar com os dois comprimentos.
```

- [ ] **Step 6: Corrigir o teste que varre o cadastro inteiro**

`tests/test_cnpj_valido.py:17-19` itera `CADASTRO_TESTE` afirmando que **toda**
chave é um CNPJ válido — com a VILA MARINA no dict ele passa a falhar. Trocar
por:

```python
    def test_cnpjs_reais_do_cadastro_sao_validos(self):
        for documento in CADASTRO_TESTE:
            if len(documento) == 11:
                continue  # VILA MARINA é identificada por CPF, não por CNPJ
            self.assertTrue(app.cnpj_valido(documento), f"deveria ser válido: {documento}")
```

- [ ] **Step 7: Rodar a suíte inteira**

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK`. Se algum outro teste quebrar por causa da entrada nova no
`CADASTRO_TESTE`, **leia a falha antes de mexer**: um teste que quebra porque
passou a existir um condomínio de 11 dígitos está apontando um lugar que
assume 14 dígitos, e provavelmente é código de produção que precisa da
correção, não o teste.

- [ ] **Step 8: Commit**

```bash
git add logica.py tests/cadastro_teste.py tests/test_cadastro_cpf.py tests/test_cnpj_valido.py
git commit -m "Guarda condomínio identificado por CPF no cadastro"
```

---

### Task 4: O formulário da aba de Cadastro aceita CPF

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py:474-490` (`_adicionar_ou_atualizar`)
- Modify: `identificacao_por_cnpj_6_0.py:389` (rótulo "CNPJ" do formulário) e `:423` (`heading` da coluna da tabela)

**Interfaces:**
- Consumes: `formatar_documento` (Task 2).
- Produces: nada consumido por outras tarefas.

**Nota:** não há teste automatizado de interface neste projeto — nenhum arquivo
em `tests/` importa `identificacao_por_cnpj_6_0`. A verificação desta tarefa é
manual, pela própria janela, descrita no Step 4. Não invente um teste de widget:
seria o primeiro do repositório e traria um aparato de mock que nada mais usa.

- [ ] **Step 1: Trocar a validação**

Em `identificacao_por_cnpj_6_0.py`, dentro de `_adicionar_ou_atualizar`,
substituir

```python
        cnpj_norm = normalizar_cnpj(cnpj_raw)
        if len(cnpj_norm) != 14:
            messagebox.showerror("Erro", "CNPJ inválido — deve ter 14 dígitos.")
            return
        if not codigo:
            messagebox.showerror("Erro", "Informe o código.")
            return

        self.cadastro[cnpj_norm] = {"codigo": codigo, "nome": nome, "id_sl": id_sl}
```

por

```python
        documento = normalizar_cnpj(cnpj_raw)
        if len(documento) not in (11, 14):
            messagebox.showerror(
                "Erro",
                "Documento inválido — informe um CNPJ (14 dígitos) ou um CPF "
                "(11 dígitos).",
            )
            return
        if not codigo:
            messagebox.showerror("Erro", "Informe o código.")
            return

        #  O documento é a chave do cadastro: trocá-lo num registro que já
        #  existe precisa REMOVER a chave antiga, senão o condomínio passa a
        #  existir duas vezes. É o caso de quem estava cadastrado por CPF e
        #  depois obteve CNPJ próprio.
        selecionado = self.tabela.selection()
        if selecionado and selecionado[0] != documento:
            self.cadastro.pop(selecionado[0], None)

        self.cadastro[documento] = {"codigo": codigo, "nome": nome, "id_sl": id_sl}
```

Conferir que não sobrou nenhum `cnpj_norm` na função:

```bash
grep -n "cnpj_norm" identificacao_por_cnpj_6_0.py
```

As ocorrências restantes devem ser de **outras** funções (`_selecionar_linha`,
`_atualizar_tabela_cadastro`, `_acao_escolher_pendente`), que não mudam aqui.

- [ ] **Step 2: Atualizar os rótulos visíveis**

Em `_montar_aba_cadastro`, trocar

```python
        caption_in(frame_form, "CNPJ").grid(row=0, column=0, sticky="w", padx=(0, 16))
```

por

```python
        caption_in(frame_form, "CNPJ / CPF").grid(row=0, column=0, sticky="w", padx=(0, 16))
```

e o cabeçalho da tabela

```python
        self.tabela.heading("cnpj", text="CNPJ")
```

por

```python
        self.tabela.heading("cnpj", text="CNPJ / CPF")
```

O identificador interno da coluna (`"cnpj"`) e a variável `self.form_cnpj` ficam
como estão: renomeá-los espalharia a mudança por toda a classe sem ganho nenhum.

- [ ] **Step 3: Conferir a sintaxe e a suíte**

```bash
python -c "import ast; ast.parse(open('identificacao_por_cnpj_6_0.py',encoding='utf-8').read()); print('sintaxe ok')"
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `sintaxe ok` e `OK`.

- [ ] **Step 4: Verificar na janela**

```bash
python identificacao_por_cnpj_6_0.py
```

Na aba de Condomínios, conferir os quatro casos, nesta ordem:

1. Documento `529.982.247-25`, código `11300`, nome `VILA MARINA` → salva, e a
   linha aparece na tabela com o CPF formatado como `529.982.247-25`.
2. Documento `123` → recusa, com a mensagem citando CNPJ **e** CPF.
3. Selecionar a linha da VILA MARINA, trocar o documento para um CNPJ válido
   (ex: `10.864.886/0001-75`) e salvar → a tabela continua com **uma** linha,
   não duas. É esse o ponto da correção do Step 1.
4. Fechar e reabrir o programa → a VILA MARINA continua lá, lida da planilha.

Ao terminar, desfazer o que os testes manuais escreveram na planilha real:

```bash
git checkout -- cadastro_condominios.xlsx
```

- [ ] **Step 5: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Aceita CPF no cadastro de condomínios"
```

---

### Task 5: Extrair CPF do campo rotulado

**Files:**
- Modify: `logica.py:63` (constantes de regex) e `logica.py:548-607` (`extrair_cnpj_tomador`)
- Test: `tests/test_extracao_cnpj.py`

**Interfaces:**
- Consumes: `cpf_valido` (Task 1).
- Produces: `CPF_REGEX` no módulo `logica` — CPF pontuado, com lookarounds que impedem casar dentro de uma sequência maior de dígitos. `extrair_cnpj_tomador` passa a devolver candidatos de 11 **ou** 14 dígitos.

- [ ] **Step 1: Escrever o teste que falha**

Em `tests/test_extracao_cnpj.py`, acrescentar junto das outras constantes do
topo do arquivo:

```python
CPF_VILA_MARINA = "52998224725"
```

e acrescentar esta classe ao final, imediatamente antes do
`if __name__ == "__main__":`:

```python
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
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_extracao_cnpj -v
```

Esperado: FAIL nos dois primeiros da classe nova (`[] != ['52998224725']`). O
terceiro e o quarto já passam hoje e precisam continuar passando.

- [ ] **Step 3: Acrescentar a constante de regex**

Em `logica.py`, logo abaixo de `CNPJ_REGEX` (linha 63):

```python
#  CPF pontuado. Os lookarounds impedem que os 11 primeiros (ou últimos)
#  dígitos de um CNPJ colado sem separador sejam lidos como um CPF — sem
#  eles, uma coincidência de checksum viraria carimbo com o código errado.
CPF_REGEX = re.compile(r"(?<!\d)\d{3}\.\d{3}\.\d{3}-\d{2}(?!\d)")
```

- [ ] **Step 4: Reescrever a montagem dos padrões de campo**

Em `extrair_cnpj_tomador`, substituir tudo que vai da linha `CNPJ_FLEX = ...`
até o fechamento da lista `PADROES_CAMPO` por:

```python
    # Tolerante a variações de separador: alguns recibos FedCorp escrevem o CNPJ do
    # co-estipulante com "-" (ou ".") no lugar da "/" — ex: 08.578.541-0001-03. O
    # separador antes do bloco 0001 e antes dos 2 dígitos finais aceita /, -, . ou espaço.
    CNPJ_FLEX = r"(\d{2}[\s.]?\d{3}[\s.]?\d{3}[\s/.\-]?\d{4}[\s.\-]?\d{2})"
    #  Mesma tolerância para o CPF do síndico, quando o condomínio não tem
    #  CNPJ próprio. Os lookarounds impedem casar dentro de um CNPJ.
    CPF_FLEX = r"(?<!\d)(\d{3}[\s.]?\d{3}[\s.]?\d{3}[\s.\-]?\d{2})(?![\d\-/])"

    documentos_a_ignorar = set(CNPJS_INTERMEDIARIOS.keys())
    documentos_a_ignorar.add(cnpj_emitente_normalizado)

    #  Um molde por rótulo, aplicado às duas formas de documento. O caminho do
    #  CNPJ continua idêntico ao que sempre foi; o do CPF só acrescenta
    #  candidatos, nunca remove.
    MOLDES_CAMPO = [
        # Boleto/recibo FedCorp: CO-ESTIPULANTE ... CNPJ: xx.xxx.xxx/xxxx-xx
        r"CO[-\s]?ESTIPULANTE[:\s]+.{0,120}?CNPJ[:\s]*{DOC}",
        # Boleto genérico: linha Pagador ... CNPJ/CPF: xxxxxxxxxxxxxxx
        r"PAGADOR[:\s]+.{0,150}?CNPJ[/\s]?CPF[:\s]*{DOC}",
        # NFS-e: TOMADOR ... CNPJ: xx.xxx.xxx/xxxx-xx
        r"TOMADOR[:\s]+.{0,120}?CNPJ[:\s]*{DOC}",
        # Detalhamento de faturamento: EMPREGADOR: nome (CNPJ xx...)
        r"EMPREGADOR[:\s]+.{0,120}?\(CNPJ[:\s]*{DOC}\)",
        # Contratos genéricos: CONTRATANTE / CLIENTE ... CNPJ
        r"(?:CONTRATANTE|CLIENTE)[:\s]+.{0,100}?CNPJ[:\s]*{DOC}",
    ]
    #  `.replace` e não `.format`: os moldes têm chaves de quantificador
    #  ({0,120}) que o format tentaria interpretar como campo de formatação.
    PADROES_CAMPO = ([molde.replace("{DOC}", CNPJ_FLEX) for molde in MOLDES_CAMPO]
                     + [molde.replace("{DOC}", CPF_FLEX) for molde in MOLDES_CAMPO])
```

Em seguida, no laço que valida os achados, substituir

```python
            cnpj_norm = normalizar_cnpj(m.group(1))
            if (len(cnpj_norm) == 14 and cnpj_valido(cnpj_norm)
                    and cnpj_norm not in cnpjs_a_ignorar):
                if cnpj_norm not in candidatos_campo:
                    candidatos_campo.append(cnpj_norm)
```

por

```python
            documento = normalizar_cnpj(m.group(1))
            valido = ((len(documento) == 14 and cnpj_valido(documento))
                      or (len(documento) == 11 and cpf_valido(documento)))
            if valido and documento not in documentos_a_ignorar:
                if documento not in candidatos_campo:
                    candidatos_campo.append(documento)
```

E, no fallback genérico logo abaixo, trocar o nome antigo do conjunto na linha

```python
    candidatos = [c for c in normalizados if cnpj_valido(c) and c not in cnpjs_a_ignorar]
```

por `documentos_a_ignorar`. Conferir que o nome antigo sumiu:

```bash
grep -n "cnpjs_a_ignorar" logica.py
```

Esperado: nenhuma saída.

Por fim, trocar o primeiro parágrafo da docstring da função por:

```
    Retorna lista de documentos candidatos (normalizados, só dígitos) — CNPJ de
    14 dígitos ou CPF de 11 — encontrados no texto, excluindo o documento da
    empresa emitente (que se repete em todo documento) e os intermediários
    conhecidos (ex: Imodata).
```

- [ ] **Step 5: Rodar os testes**

```bash
python -m unittest tests.test_extracao_cnpj -v
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK` nos dois. Se algum teste **antigo** de `test_extracao_cnpj.py`
quebrar, pare e investigue: significa que a montagem por moldes não reproduziu
fielmente algum padrão de CNPJ, e o caminho do CNPJ não pode mudar de
comportamento nesta tarefa.

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_extracao_cnpj.py
git commit -m "Extrai CPF do campo do pagador nos documentos"
```

---

### Task 6: CPF na varredura genérica, só se cadastrado

**Files:**
- Modify: `logica.py:548` (assinatura de `extrair_cnpj_tomador`) e o bloco de fallback genérico
- Modify: `identificacao_por_cnpj_6_0.py:4040`, `:4074`, `:4093`
- Test: `tests/test_extracao_cnpj.py`

**Interfaces:**
- Consumes: `cpf_valido` (Task 1), `CPF_REGEX` (Task 5).
- Produces: `extrair_cnpj_tomador(texto, cnpj_emitente_normalizado, cadastro=None) -> list[str]`. O terceiro parâmetro é **opcional** e só afeta o fallback genérico: sem ele, nenhum CPF entra por ali. As chamadas de dois argumentos continuam válidas.

- [ ] **Step 1: Escrever o teste que falha**

Em `tests/test_extracao_cnpj.py`, acrescentar ao bloco de imports do topo:

```python
from cadastro_teste import CADASTRO_TESTE
```

e acrescentar esta classe ao final, antes do `if __name__ == "__main__":`:

```python
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
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_extracao_cnpj -v
```

Esperado: FAIL com `TypeError: extrair_cnpj_tomador() takes 2 positional arguments but 3 were given`.

- [ ] **Step 3: Implementar**

Trocar a assinatura em `logica.py`:

```python
def extrair_cnpj_tomador(texto, cnpj_emitente_normalizado, cadastro=None):
```

e substituir o bloco do fallback genérico (a partir do comentário
`# --- Fallback: varredura genérica (comportamento original) ---` até o `return`
final da função) por:

```python
    # --- Fallback: varredura genérica (comportamento original) ---
    encontrados = CNPJ_REGEX.findall(texto)
    normalizados = [normalizar_cnpj(c) for c in encontrados]
    candidatos = [c for c in normalizados
                  if cnpj_valido(c) and c not in documentos_a_ignorar]

    #  CPF é tratado aqui de forma deliberadamente mais restrita que o CNPJ:
    #  sem rótulo, um CPF solto na página costuma ser de uma pessoa qualquer
    #  (síndico, avalista, quem assinou), enquanto um CNPJ solto tende a ser de
    #  alguma empresa envolvida na cobrança. Só entra quem já é condomínio
    #  conhecido. Custo assumido: condomínio novo por CPF não vira pendente
    #  "não cadastrado" por este caminho — ele cai no match por código/nome do
    #  arquivo, que continua valendo como rede.
    if cadastro:
        for achado in CPF_REGEX.findall(texto):
            documento = normalizar_cnpj(achado)
            if (cpf_valido(documento) and documento in cadastro
                    and documento not in documentos_a_ignorar):
                candidatos.append(documento)

    vistos = set()
    unicos = []
    for c in candidatos:
        if c not in vistos:
            vistos.add(c)
            unicos.append(c)
    return unicos
```

Acrescentar ao final da docstring da função:

```
    `cadastro` é opcional e só afeta a varredura genérica: sem ele, nenhum CPF
    entra por ali. Ver o comentário no bloco do fallback.
```

- [ ] **Step 4: Passar o cadastro nas três chamadas da interface**

Em `identificacao_por_cnpj_6_0.py`, acrescentar `self.cadastro` como terceiro
argumento nas linhas 4040, 4074 e 4093:

```python
                            candidatos_regiao = extrair_cnpj_tomador(texto_regiao, cnpj_emitente_norm, self.cadastro)
```

```python
                        candidatos = extrair_cnpj_tomador(texto, cnpj_emitente_norm, self.cadastro)
```

```python
                                candidatos_retry = extrair_cnpj_tomador(texto_retry, cnpj_emitente_norm, self.cadastro)
```

Conferir:

```bash
grep -n "extrair_cnpj_tomador" identificacao_por_cnpj_6_0.py
```

Esperado: 4 linhas — o import na 70 e as três chamadas, as três com `self.cadastro`.

- [ ] **Step 5: Rodar os testes**

```bash
python -m unittest tests.test_extracao_cnpj -v
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

```bash
python -c "import ast; ast.parse(open('identificacao_por_cnpj_6_0.py',encoding='utf-8').read()); print('sintaxe ok')"
```

Esperado: `OK`, `OK`, `sintaxe ok`.

- [ ] **Step 6: Commit**

```bash
git add logica.py identificacao_por_cnpj_6_0.py tests/test_extracao_cnpj.py
git commit -m "Aceita CPF já cadastrado na varredura sem rótulo"
```

---

### Task 7: Tomador pessoa física na NFS-e

**Files:**
- Modify: `logica.py:1349-1354` (dentro de `extrair_dados_nfse`)
- Test: `tests/test_extracao_nfse.py`

**Interfaces:**
- Consumes: `cpf_valido` (Task 1), `CPF_REGEX` (Task 5).
- Produces: a chave `"cnpj_tomador"` do dict devolvido por `extrair_dados_nfse` passa a poder conter 11 dígitos. **O nome da chave não muda** — renomeá-la espalharia a mudança por `linha_planilha_nfse` e por toda a suíte de NFS-e sem nenhum ganho de comportamento.

- [ ] **Step 1: Escrever o teste que falha**

A fixture sai da nota real que já existe, com o CNPJ do tomador trocado por um
CPF. `01.195.716/0001-54` aparece **uma única vez** em `tests/dados/nfse_ff.txt`
(no bloco do tomador; o emitente tem outro CNPJ), então o `replace` atinge só o
campo certo — e a nota continua sendo uma nota inteira, que passa por todos os
recortes de seção de `extrair_dados_nfse`.

Em `tests/test_extracao_nfse.py`, acrescentar esta classe antes do
`if __name__ == "__main__":`:

```python
class TestTomadorPessoaFisica(unittest.TestCase):
    """Condomínio sem CNPJ próprio aparece na nota com o CPF do síndico no
    bloco do tomador. Sem ler os dois formatos, a planilha de notas desses
    condomínios sairia sistematicamente sem código."""

    def setUp(self):
        #  Mesma nota da fixture, com o documento do tomador trocado por um
        #  CPF. O rótulo do DANFSe já é "CNPJ / CPF / NIF" — só o valor muda.
        self.texto = ler("nfse_ff.txt").replace("01.195.716/0001-54",
                                                 "529.982.247-25")

    def test_acha_cpf_no_bloco_do_tomador(self):
        dados = app.extrair_dados_nfse(self.texto)
        self.assertIsNotNone(dados)
        self.assertEqual(dados["cnpj_tomador"], "52998224725")

    def test_cpf_com_dv_errado_nao_entra(self):
        texto = self.texto.replace("529.982.247-25", "529.982.247-24")
        dados = app.extrair_dados_nfse(texto)
        self.assertIsNotNone(dados)
        self.assertEqual(dados["cnpj_tomador"], "")

    def test_nota_com_cnpj_continua_lendo_o_cnpj(self):
        """A fixture original não pode mudar de resultado."""
        dados = app.extrair_dados_nfse(ler("nfse_ff.txt"))
        self.assertEqual(dados["cnpj_tomador"], KLOSTERS)
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_extracao_nfse -v
```

Esperado: FAIL em `test_acha_cpf_no_bloco_do_tomador` com `'' != '52998224725'`.
Os outros dois já passam e precisam continuar passando.

- [ ] **Step 3: Implementar**

Em `logica.py`, dentro de `extrair_dados_nfse`, substituir

```python
    cnpj_tomador = ""
    m_cnpj = CNPJ_REGEX.search(bloco_tomador)
    if m_cnpj:
        candidato = normalizar_cnpj(m_cnpj.group(0))
        if cnpj_valido(candidato):
            cnpj_tomador = candidato
```

por

```python
    #  O tomador pode ser pessoa jurídica (CNPJ) ou física (CPF do síndico,
    #  nos condomínios sem CNPJ próprio). A variável e a chave do dict
    #  continuam se chamando "cnpj_tomador" para não espalhar a renomeação
    #  por toda a planilha de notas — o conteúdo é o documento, seja qual for.
    cnpj_tomador = ""
    m_cnpj = CNPJ_REGEX.search(bloco_tomador)
    if m_cnpj:
        candidato = normalizar_cnpj(m_cnpj.group(0))
        if cnpj_valido(candidato):
            cnpj_tomador = candidato
    if not cnpj_tomador:
        m_cpf = CPF_REGEX.search(bloco_tomador)
        if m_cpf:
            candidato = normalizar_cnpj(m_cpf.group(0))
            if cpf_valido(candidato):
                cnpj_tomador = candidato
```

O CNPJ é tentado primeiro e o CPF só entra se ele não achou nada, então nenhuma
nota que hoje funciona muda de resultado.

- [ ] **Step 4: Rodar os testes**

```bash
python -m unittest tests.test_extracao_nfse -v
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK` nos dois.

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_extracao_nfse.py
git commit -m "Lê tomador pessoa física na NFS-e"
```

---

### Task 8: O carimbo do protocolo omite o CPF

**Files:**
- Modify: `logica.py:697-700` (`montar_texto_protocolo_correio`)
- Test: `tests/test_carimbo_lateral.py`

**Interfaces:**
- Consumes: `formatar_documento` (Task 2).
- Produces: `montar_texto_protocolo_correio(codigo, nome, documento_normalizado) -> str` — com CNPJ devolve `"{código} {nome} - {CNPJ formatado}"` como sempre; com CPF (ou documento vazio) devolve `"{código} {nome}"`.

- [ ] **Step 1: Escrever o teste que falha**

Em `tests/test_carimbo_lateral.py`, dentro da classe
`TestMontarTextoProtocoloCorreio` que já existe, acrescentar:

```python
    def test_omite_o_cpf(self):
        """O PDF carimbado circula e vai para o Superlógica. Estampar o CPF de
        uma pessoa física nele é diferente de estampar o CNPJ de um
        condomínio, então o documento fica de fora — e o traço junto."""
        texto = app.montar_texto_protocolo_correio("11300", "VILA MARINA",
                                                    "52998224725")
        self.assertEqual(texto, "11300 VILA MARINA")

    def test_sem_documento_tambem_omite(self):
        texto = app.montar_texto_protocolo_correio("11300", "VILA MARINA", "")
        self.assertEqual(texto, "11300 VILA MARINA")
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_carimbo_lateral -v
```

Esperado: FAIL com `'11300 VILA MARINA - 529.982.247-25' != '11300 VILA MARINA'`.

- [ ] **Step 3: Implementar**

Em `logica.py`, substituir `montar_texto_protocolo_correio` inteira por:

```python
def montar_texto_protocolo_correio(codigo, nome, documento_normalizado):
    """
    Linha única do carimbo lateral do Protocolo de Recebimento de Documento.

    Com CNPJ sai "{código} {nome} - {CNPJ}", como sempre. Com CPF o documento é
    OMITIDO: o PDF carimbado circula e vai para o Superlógica, e estampar o CPF
    de uma pessoa física nele é diferente de estampar o CNPJ de um condomínio.
    """
    documento = re.sub(r"\D", "", documento_normalizado or "")
    if len(documento) != 14:
        return f"{codigo} {nome}"
    return f"{codigo} {nome} - {formatar_documento(documento)}"
```

- [ ] **Step 4: Rodar os testes**

```bash
python -m unittest tests.test_carimbo_lateral -v
```

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK` nos dois. `test_formata_codigo_nome_cnpj`, que já existia, precisa
continuar passando — é ele que garante que o caso do CNPJ não mudou.

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_carimbo_lateral.py
git commit -m "Deixa o CPF fora do carimbo do protocolo"
```

---

### Task 9: Documentar no CLAUDE.md

**Files:**
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: tudo das tarefas 1 a 8.
- Produces: nada.

- [ ] **Step 1: Acrescentar a entrada no histórico de decisões**

Em `CLAUDE.md`, na lista da seção "Histórico de decisões", acrescentar como
primeiro item:

```markdown
- **v6.15.0 — condomínios identificados por CPF**: alguns condomínios não têm
  CNPJ próprio e aparecem nos boletos e notas com o **CPF do síndico**. A chave
  do cadastro deixou de ser "CNPJ" e passou a ser **documento normalizado**: 14
  dígitos (CNPJ) ou 11 (CPF). Funciona sem ambiguidade porque o comprimento
  distingue os dois sozinho, e cada um tem seu próprio dígito verificador
  (`cnpj_valido` e `cpf_valido`, `logica.py`). `formatar_cnpj` virou
  `formatar_documento`, formatando pelo comprimento. A coluna A da planilha
  passou a se chamar `CNPJ / CPF`; `carregar_cadastro` lê por posição e nunca
  pelo nome do cabeçalho, então planilha antiga continua abrindo. Junto veio
  uma correção: editar o documento de um registro existente **criava um
  segundo** em vez de atualizar o primeiro, porque o documento é a chave — é
  exatamente o fluxo de quem estava cadastrado por CPF e depois obteve CNPJ.
  Ver o spec `docs/superpowers/specs/2026-08-24-cadastro-por-cpf-design.md`.
```

- [ ] **Step 2: Documentar a assimetria da extração**

Ainda em `CLAUDE.md`, acrescentar uma seção nova logo depois de "Regra de
negócio importante — CNPJs que NUNCA são o condomínio tomador":

```markdown
## CPF do síndico: por que a regra é diferente da do CNPJ

Condomínio sem CNPJ próprio é identificado pelo **CPF do síndico**, que sai
impresso no campo do tomador/pagador. A extração trata os dois em dois níveis,
com pesos deliberadamente diferentes:

- **Depois de um rótulo** (`PAGADOR ... CNPJ/CPF`, `TOMADOR`,
  `CO-ESTIPULANTE`...), o CPF é sempre aceito: o rótulo é a garantia de que
  aquele documento é o do pagador.
- **Na varredura genérica**, sem rótulo legível, o CPF **só vira candidato se
  já estiver no cadastro** (parâmetro opcional `cadastro` de
  `extrair_cnpj_tomador`). Num boleto, um CNPJ solto tende a ser de alguma
  empresa envolvida na cobrança, mas um CPF solto costuma ser de uma pessoa
  qualquer — síndico, sacador avalista, quem assinou. Aceitar todo CPF válido
  da página encheria o processo de falso positivo, e falso positivo aqui é PDF
  carimbado com o código do condomínio errado.

**Custo assumido:** um condomínio novo identificado por CPF **não** vira
pendente "não cadastrado" por esse caminho — ele cai no match por código/nome
do arquivo, que continua valendo como rede. Quem for afrouxar a regra depois
precisa saber que o ganho é esse e o risco é o carimbo errado.

`CPF_REGEX` e `CPF_FLEX` (`logica.py`) começam com o lookbehind `(?<!\d)` e
terminam com um lookahead que barra a continuação do número — de propósito: sem
eles, os 11 primeiros dígitos de um CNPJ colado sem separador (OCR) poderiam
casar como CPF, e uma coincidência de checksum viraria carimbo errado.

**O CPF fica fora do carimbo do protocolo dos Correios**
(`montar_texto_protocolo_correio`): com CNPJ o carimbo é
`"{código} {nome} - {CNPJ}"`, com CPF sai só `"{código} {nome}"`. O PDF
carimbado circula e vai para o Superlógica, e estampar o CPF de uma pessoa
física nele é diferente de estampar o CNPJ de um condomínio.
```

- [ ] **Step 3: Atualizar as convenções**

Na seção "Convenções do código" de `CLAUDE.md`, substituir a linha

```markdown
- CNPJs sempre normalizados para 14 dígitos internamente (`normalizar_cnpj`),
  formatados só na exibição/planilha (`formatar_cnpj`).
```

por

```markdown
- Documentos sempre normalizados para só dígitos internamente
  (`normalizar_cnpj`) — 14 para CNPJ, 11 para CPF —, formatados só na
  exibição/planilha (`formatar_documento`). O comprimento é o que distingue os
  dois; nunca inferir o tipo de outro jeito.
```

- [ ] **Step 4: Conferir a suíte uma última vez**

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: `OK`.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md
git commit -m "Documenta a identificação de condomínio por CPF"
```
