# Testes automatizados da identificação (v6.3.0) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Congelar em testes automatizados os casos reais de identificação já
corrigidos, de forma que qualquer regressão que passe a carimbar código errado
apareça em segundos.

**Architecture:** Testes `unittest` numa pasta `tests/`, lendo fixtures de texto
extraído dos PDFs reais e um cadastro de teste fixo de 9 condomínios (nunca a
planilha real). A maioria são *testes de caracterização*: travam o comportamento
atual de funções que já funcionam. Uma única alteração no app extrai a regra de
desempate para uma função de módulo testável.

**Tech Stack:** Python 3, biblioteca padrão `unittest` (sem dependência nova),
`.bat` para rodar por duplo clique.

## Global Constraints

- **Sem dependência nova:** só `unittest` (stdlib). Nada de `pip install`.
- **Não tocar em `cadastro_condominios.xlsx`** (planilha real). Testes usam
  `tests/cadastro_teste.py` (9 entradas fixas).
- **Não mudar o formato de `processamento.log`** — o refactor do desempate
  (Task 5) deve manter o sufixo `" (desempate: CNPJ cadastrado)"` idêntico.
- **Única alteração permitida no app:** extrair `desempatar_por_cadastro()` em
  `identificacao_por_cnpj_6_0.py`. Nenhuma outra função do app muda.
- **Plataforma:** Windows. Comando de teste roda da raiz do repo:
  `python -m unittest discover -s tests -p "test_*.py"`.
- **Arquivo sob teste:** `identificacao_por_cnpj_6_0.py` (o 5_3 não é testado).
- **Idioma:** nomes e comentários em português, seguindo o código existente.

---

## Estrutura de arquivos

```
tests/
  cadastro_teste.py          CADASTRO_TESTE: dict cnpj_norm -> {codigo, nome} (9 fixos)
  _gerar_fixtures.py         helper de uso único: extrai texto dos PDFs reais -> dados/*.txt
  dados/
    fedcorp_recibo_hifen.txt
    nfse_ff.txt
    nfse_imodata.txt
    nfse_avulsa.txt
    boleto_avulso.txt
    cnpj_checksum_invalido.txt
  test_infra.py              smoke test: app importa, cadastro carrega, fixtures existem
  test_cnpj_valido.py
  test_nome_arquivo.py
  test_saida.py
  test_desempate.py
  test_extracao_cnpj.py
  test_config.py
rodar_testes.bat             duplo clique -> roda tudo
```

Cada `test_*.py` começa com o mesmo bloco de bootstrap de path (4 linhas), para
funcionar tanto pelo `.bat` quanto rodado individualmente. É código de
inicialização que precisa correr **antes** de qualquer import do app, por isso é
repetido e não fatorado.

---

## Task 1: Estrutura — cadastro de teste, fixtures, runner e smoke test

**Files:**
- Create: `tests/cadastro_teste.py`
- Create: `tests/_gerar_fixtures.py`
- Create: `tests/dados/*.txt` (gerados pelo script acima)
- Create: `tests/test_infra.py`
- Create: `rodar_testes.bat`

**Interfaces:**
- Produces: `tests.cadastro_teste.CADASTRO_TESTE` — `dict[str, dict]` mapeando
  CNPJ normalizado (14 dígitos) → `{"codigo": str, "nome": str}`, com as 9
  entradas da tabela abaixo. Todas as tasks seguintes consomem isto.
- Produces: os 6 arquivos `tests/dados/*.txt` (texto UTF-8). Tasks 5 e 6 os leem.

- [ ] **Step 1: Criar o cadastro de teste**

Create `tests/cadastro_teste.py`:

```python
"""
Cadastro fixo usado só pelos testes — NUNCA a planilha real
(cadastro_condominios.xlsx). Congelado de propósito: se os testes lessem a
planilha real, cadastrar um condomínio novo poderia quebrar um teste sem que
nada esteja errado. Os 9 condomínios abaixo são os que apareceram nos bugs
reais corrigidos nas versões 6.1/6.2.

Chave = CNPJ normalizado (14 dígitos). Valor = {"codigo", "nome"}.
"""

CADASTRO_TESTE = {
    "40338774000141": {"codigo": "10695", "nome": "ARAUJO LIMA"},
    "01195716000154": {"codigo": "10004", "nome": "KLOSTERS"},
    "08578541000103": {"codigo": "10002", "nome": "SAN REMO"},
    "05695194000100": {"codigo": "10490", "nome": "CENTRO COM CONDE DE BONFIM RES"},
    "29361458000158": {"codigo": "10491", "nome": "CENTRO COM CONDE DE BONFIM"},
    "07448975000126": {"codigo": "11194", "nome": "MANHATTAN"},
    "10864886000175": {"codigo": "10625", "nome": "ARGENTINA"},
    "29273778000156": {"codigo": "11189", "nome": "VILLE DE BEAUVAIS"},
    "07945453000130": {"codigo": "10590", "nome": "LAGO MAGGIORE"},
}
```

- [ ] **Step 2: Criar o gerador de fixtures**

Create `tests/_gerar_fixtures.py`. Uso único — extrai o texto dos PDFs reais e
grava em `tests/dados/`. Os caminhos-fonte são desta máquina; se for regenerar
noutra, ajuste-os. Os `.txt` gerados é que ficam versionados; este script é só
proveniência.

```python
"""
Gera os fixtures de texto em tests/dados/ a partir dos PDFs reais.
Uso único: `python tests/_gerar_fixtures.py`. Os PDFs-fonte não vão para o
repositório; só os .txt resultantes. Ajuste FONTES se regenerar noutra máquina.
"""
import os
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
from identificacao_por_cnpj_6_0 import extrair_texto_pdf

FONTES = {
    "fedcorp_recibo_hifen.txt": r"C:\Users\Dell\Downloads\BENEFICIO 07-2026 BOLETO\NF-669.pdf",
    "nfse_ff.txt": r"C:\Users\Dell\Documents\NOTAS FF\FF JUNHO 2026\06 Junho\originais\PGR\PGR 10004 Klosters.pdf",
    "nfse_imodata.txt": r"C:\Users\Dell\Downloads\nfses_pdfs_2026-05-04\NFSE_284726_11194.pdf",
    "nfse_avulsa.txt": r"C:\Users\Dell\Documents\EXEMPLO NF.pdf",
    "boleto_avulso.txt": r"C:\Users\Dell\Documents\EXEMPLO BOLETO.pdf",
    "cnpj_checksum_invalido.txt": r"C:\Users\Dell\Downloads\BENEFICIO 07-2026 BOLETO\NF-927.pdf",
}

destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados")
os.makedirs(destino, exist_ok=True)

for nome_txt, caminho_pdf in FONTES.items():
    if not os.path.isfile(caminho_pdf):
        print(f"AVISO: fonte não encontrada, pulando: {caminho_pdf}")
        continue
    texto = extrair_texto_pdf(caminho_pdf)
    with open(os.path.join(destino, nome_txt), "w", encoding="utf-8") as f:
        f.write(texto)
    print(f"gerado: dados/{nome_txt} ({len(texto)} chars)")
```

- [ ] **Step 3: Rodar o gerador e confirmar os 6 fixtures**

Run: `python tests/_gerar_fixtures.py`
Expected: 6 linhas "gerado: dados/...", nenhum "AVISO". Confirme que
`tests/dados/` tem os 6 `.txt`.

- [ ] **Step 4: Criar o runner .bat**

Create `rodar_testes.bat` (na RAIZ do repo):

```bat
@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m unittest discover -s tests -p "test_*.py" -v
echo.
pause
```

- [ ] **Step 5: Escrever o smoke test**

Create `tests/test_infra.py`:

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

import identificacao_por_cnpj_6_0 as app
from cadastro_teste import CADASTRO_TESTE


class TestInfra(unittest.TestCase):
    def test_app_importa(self):
        self.assertTrue(hasattr(app, "buscar_por_nome_arquivo"))

    def test_cadastro_tem_9_entradas(self):
        self.assertEqual(len(CADASTRO_TESTE), 9)
        # todos os CNPJs têm 14 dígitos
        for cnpj in CADASTRO_TESTE:
            self.assertEqual(len(cnpj), 14)

    def test_fixtures_existem_e_tem_texto(self):
        dados = os.path.join(_AQUI, "dados")
        for nome in ("fedcorp_recibo_hifen.txt", "nfse_ff.txt", "nfse_imodata.txt",
                     "nfse_avulsa.txt", "boleto_avulso.txt", "cnpj_checksum_invalido.txt"):
            caminho = os.path.join(dados, nome)
            self.assertTrue(os.path.isfile(caminho), f"faltando fixture: {nome}")
            with open(caminho, encoding="utf-8") as f:
                self.assertGreater(len(f.read().strip()), 100, f"fixture vazio: {nome}")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 6: Rodar e confirmar que passa**

Run: `python -m unittest discover -s tests -p "test_*.py" -v`
Expected: 3 testes OK.

- [ ] **Step 7: Commit**

```bash
git add tests/cadastro_teste.py tests/_gerar_fixtures.py tests/dados/ tests/test_infra.py rodar_testes.bat
git commit -m "test: estrutura de testes (cadastro fixo, fixtures, runner)"
```

---

## Task 2: Testes de validação de CNPJ (`cnpj_valido`)

**Files:**
- Create: `tests/test_cnpj_valido.py`

**Interfaces:**
- Consumes: `app.cnpj_valido(cnpj_normalizado: str) -> bool` (função existente).

Função pura, sem fixtures. Trava o comportamento atual, incluindo o caso real
que evita carimbar código errado (NF-927, `00.001.208-2497-38`).

- [ ] **Step 1: Escrever os testes**

Create `tests/test_cnpj_valido.py`:

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

import identificacao_por_cnpj_6_0 as app
from cadastro_teste import CADASTRO_TESTE


class TestCnpjValido(unittest.TestCase):
    def test_cnpjs_reais_do_cadastro_sao_validos(self):
        for cnpj in CADASTRO_TESTE:
            self.assertTrue(app.cnpj_valido(cnpj), f"deveria ser válido: {cnpj}")

    def test_checksum_invalido_do_nf927(self):
        # CNPJ que a FedCorp gerou errado (00.001.208-2497-38) — dígito não bate
        self.assertFalse(app.cnpj_valido("00001208249738"))

    def test_todos_digitos_iguais_e_invalido(self):
        self.assertFalse(app.cnpj_valido("11111111111111"))

    def test_menos_de_14_digitos_e_invalido(self):
        self.assertFalse(app.cnpj_valido("123"))
        self.assertFalse(app.cnpj_valido(""))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar e confirmar que passa**

Run: `python tests/test_cnpj_valido.py -v`
Expected: 4 testes OK. (Passam de imediato — são de caracterização; travam o
que já funciona.)

- [ ] **Step 3: Commit**

```bash
git add tests/test_cnpj_valido.py
git commit -m "test: validação de dígitos verificadores de CNPJ"
```

---

## Task 3: Testes de identificação pelo nome do arquivo (`buscar_por_nome_arquivo`)

**Files:**
- Create: `tests/test_nome_arquivo.py`

**Interfaces:**
- Consumes: `app.buscar_por_nome_arquivo(nome_arquivo: str, cadastro: dict) ->
  tuple[str | None, str]`. Devolve `(cnpj_norm, descricao)` num match único, ou
  `(None, motivo)` quando não acha ou é ambíguo.

Esta é a suíte mais importante: inclui a trava que impede trocar 10490 por
10491. Todos os valores esperados foram verificados contra o cadastro de teste.

- [ ] **Step 1: Escrever os testes**

Create `tests/test_nome_arquivo.py`:

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

import identificacao_por_cnpj_6_0 as app
from cadastro_teste import CADASTRO_TESTE


def codigo_de(cnpj):
    return CADASTRO_TESTE[cnpj]["codigo"] if cnpj else None


class TestNomeArquivo(unittest.TestCase):
    def test_codigo_exato_no_nome(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 Klosters.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10004")

    def test_codigo_nao_confunde_com_data(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 Klosters 06.2026.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10004")

    def test_fuzzy_pelo_nome(self):
        cnpj, _ = app.buscar_por_nome_arquivo("ARAUJO LIMA QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10695")

    def test_palavra_tipo_documento_nao_atrapalha(self):
        # "QUITADO" derrubava o score abaixo de 0.72 antes do fix
        cnpj, _ = app.buscar_por_nome_arquivo("ARGENTINA QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertEqual(codigo_de(cnpj), "10625")

    def test_ambiguidade_conde_de_bonfim_nao_escolhe(self):
        # TRAVA CRÍTICA: dois condomínios quase idênticos -> não pode escolher
        cnpj, _ = app.buscar_por_nome_arquivo("CONDE DE BONFIM.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_ambiguidade_conde_de_bonfim_com_ruido_nao_escolhe(self):
        cnpj, _ = app.buscar_por_nome_arquivo(
            "CENTRO COM CONDE DE BONFIM QUITADO 05.26.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_dois_codigos_no_nome_e_ambiguo(self):
        cnpj, _ = app.buscar_por_nome_arquivo("PGR 10004 10002 Klosters.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)

    def test_sem_correspondencia(self):
        cnpj, _ = app.buscar_por_nome_arquivo("DOCUMENTO QUALQUER.pdf", CADASTRO_TESTE)
        self.assertIsNone(cnpj)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar e confirmar que passa**

Run: `python tests/test_nome_arquivo.py -v`
Expected: 8 testes OK.

- [ ] **Step 3: Provar que a trava tem dentes (verificação manual, não commitada)**

No `identificacao_por_cnpj_6_0.py`, temporariamente troque
`LIMIAR_DIFERENCA_AMBIGUA = 0.08` por `= 0.0` (desliga a proteção de
ambiguidade). Rode `python tests/test_nome_arquivo.py -v`.
Expected: `test_ambiguidade_conde_de_bonfim_nao_escolhe` (e possivelmente o
"_com_ruido") FALHAM. Isso confirma que o teste pega a regressão.
**Reverta a mudança** (`= 0.08`) antes de continuar e reconfirme que tudo passa.

- [ ] **Step 4: Commit**

```bash
git add tests/test_nome_arquivo.py
git commit -m "test: identificação pelo nome do arquivo (inclui trava de ambiguidade)"
```

---

## Task 4: Testes do nome do arquivo de saída (`nome_saida_com_codigo`)

**Files:**
- Create: `tests/test_saida.py`

**Interfaces:**
- Consumes: `app.nome_saida_com_codigo(nome_original: str, codigo: str) -> str`.

- [ ] **Step 1: Escrever os testes**

Create `tests/test_saida.py`:

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

import identificacao_por_cnpj_6_0 as app


class TestNomeSaida(unittest.TestCase):
    def test_prefixa_codigo(self):
        self.assertEqual(
            app.nome_saida_com_codigo("PGR 10004 Klosters.pdf", "10004"),
            "10004 - PGR 10004 Klosters.pdf",
        )

    def test_sanitiza_caractere_invalido_no_codigo(self):
        # barras e dois-pontos não podem aparecer em nome de arquivo no Windows
        resultado = app.nome_saida_com_codigo("nota.pdf", "10/04")
        for proibido in '\\/:*?"<>|':
            self.assertNotIn(proibido, resultado.split(" - ")[0])

    def test_codigo_vazio_devolve_nome_original(self):
        self.assertEqual(app.nome_saida_com_codigo("nota.pdf", ""), "nota.pdf")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar e confirmar que passa**

Run: `python tests/test_saida.py -v`
Expected: 3 testes OK.

- [ ] **Step 3: Commit**

```bash
git add tests/test_saida.py
git commit -m "test: nome do arquivo de saída com código"
```

---

## Task 5: Extrair `desempatar_por_cadastro` + testes (única alteração no app)

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (extrair função; ajustar o loop)
- Create: `tests/test_desempate.py`

**Interfaces:**
- Produces: `app.desempatar_por_cadastro(candidatos: list[str], cadastro: dict)
  -> list[str]`. Entre vários CNPJs candidatos, se exatamente um estiver no
  cadastro, devolve `[esse_cnpj]`; senão devolve a lista inalterada.
- Consumes (no loop): a função nova, preservando o sufixo de log.

- [ ] **Step 1: Escrever o teste da função nova (ainda não existe)**

Create `tests/test_desempate.py`:

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

import identificacao_por_cnpj_6_0 as app
from cadastro_teste import CADASTRO_TESTE

IMODATA_EMITENTE = "31850191000104"   # não cadastrado
MANHATTAN = "07448975000126"          # cadastrado (11194)


class TestDesempate(unittest.TestCase):
    def test_dois_candidatos_so_um_cadastrado(self):
        resultado = app.desempatar_por_cadastro(
            [IMODATA_EMITENTE, MANHATTAN], CADASTRO_TESTE)
        self.assertEqual(resultado, [MANHATTAN])

    def test_dois_cadastrados_continua_ambiguo(self):
        dois = ["07448975000126", "10864886000175"]  # MANHATTAN e ARGENTINA
        resultado = app.desempatar_por_cadastro(dois, CADASTRO_TESTE)
        self.assertEqual(resultado, dois)

    def test_nenhum_cadastrado_inalterado(self):
        nenhum = ["31850191000104", "12184361000114"]
        resultado = app.desempatar_por_cadastro(nenhum, CADASTRO_TESTE)
        self.assertEqual(resultado, nenhum)

    def test_um_candidato_so_inalterado(self):
        resultado = app.desempatar_por_cadastro([MANHATTAN], CADASTRO_TESTE)
        self.assertEqual(resultado, [MANHATTAN])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar e confirmar que FALHA**

Run: `python tests/test_desempate.py -v`
Expected: FAIL — `AttributeError: module ... has no attribute 'desempatar_por_cadastro'`.

- [ ] **Step 3: Criar a função no app**

Em `identificacao_por_cnpj_6_0.py`, logo após `buscar_por_nome_arquivo`
(procure `    return melhor_cnpj, f"nome do arquivo, score` e insira a nova
função de módulo depois do fim dessa função), adicione:

```python
def desempatar_por_cadastro(candidatos, cadastro):
    """
    Entre vários CNPJs candidatos, prefere o único que já está cadastrado.
    Usado por lotes sem emitente fixo (ex: Notas Diversas), onde os candidatos
    costumam ser o emitente da nota (não cadastrado) + o condomínio tomador
    (cadastrado). Devolve [o_cadastrado] se houver exatamente um cadastrado;
    caso contrário devolve a lista inalterada.
    """
    if len(candidatos) <= 1:
        return candidatos
    cadastrados = [c for c in candidatos if c in cadastro]
    if len(cadastrados) == 1:
        return cadastrados
    return candidatos
```

- [ ] **Step 4: Ligar o loop na função nova**

Em `_processar_em_thread`, substitua o bloco 2b.1 (hoje):

```python
                    # 2b.1) Lote sem emitente fixo (ex: Notas Diversas) — mais de
                    #       um candidato costuma ser o emitente da nota (não
                    #       cadastrado) + o condomínio tomador (cadastrado). Se
                    #       sobrar exatamente um candidato já cadastrado, usa ele.
                    if len(candidatos) > 1 and preferir_cadastrado_em_ambiguo:
                        cadastrados = [c for c in candidatos if c in self.cadastro]
                        if len(cadastrados) == 1:
                            candidatos = cadastrados
                            sufixo_origem += " (desempate: CNPJ cadastrado)"
```

por:

```python
                    # 2b.1) Lote sem emitente fixo (ex: Notas Diversas) — mais de
                    #       um candidato costuma ser o emitente da nota (não
                    #       cadastrado) + o condomínio tomador (cadastrado). Se
                    #       sobrar exatamente um candidato já cadastrado, usa ele.
                    if preferir_cadastrado_em_ambiguo:
                        desempatados = desempatar_por_cadastro(candidatos, self.cadastro)
                        if desempatados != candidatos:
                            candidatos = desempatados
                            sufixo_origem += " (desempate: CNPJ cadastrado)"
```

O comportamento e o texto do log ficam idênticos: `desempatar_por_cadastro` só
muda a lista quando havia >1 candidato e exatamente um cadastrado — exatamente
os casos em que o sufixo era adicionado antes.

- [ ] **Step 5: Rodar o teste do desempate e confirmar que passa**

Run: `python tests/test_desempate.py -v`
Expected: 4 testes OK.

- [ ] **Step 6: Confirmar que nada mais quebrou (py_compile + suíte inteira)**

Run: `python -m py_compile identificacao_por_cnpj_6_0.py`
Expected: sem erro.
Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: todos OK.

- [ ] **Step 7: Commit**

```bash
git add identificacao_por_cnpj_6_0.py tests/test_desempate.py
git commit -m "refactor+test: extrai desempatar_por_cadastro para função testável"
```

---

## Task 6: Testes de extração de CNPJ do conteúdo (`extrair_cnpj_tomador`)

**Files:**
- Create: `tests/test_extracao_cnpj.py`

**Interfaces:**
- Consumes: `app.extrair_cnpj_tomador(texto: str, cnpj_emitente_normalizado:
  str) -> list[str]` e `app.CNPJS_INTERMEDIARIOS` (dict com FedCorp
  `35315360000167` e Imodata `12184361000114`).

Valores esperados verificados contra os fixtures reais. Nota: em notas com dois
CNPJs (Imodata, avulsa, boleto), `extrair_cnpj_tomador` devolve **os dois** — o
desempate (Task 5) é que resolve. Os testes afirmam essa divisão.

- [ ] **Step 1: Escrever os testes**

Create `tests/test_extracao_cnpj.py`:

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

import identificacao_por_cnpj_6_0 as app

FED = "35315360000167"   # FedCorp (emitente/intermediário)
FF = "13736666000154"    # F&F (emitente)
SAN_REMO = "08578541000103"
KLOSTERS = "01195716000154"
MANHATTAN = "07448975000126"
LAGO_MAGGIORE = "07945453000130"
VILLE = "29273778000156"
IMODATA_INTERM = "12184361000114"


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


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar e confirmar que passa**

Run: `python tests/test_extracao_cnpj.py -v`
Expected: 7 testes OK.

- [ ] **Step 3: Commit**

```bash
git add tests/test_extracao_cnpj.py
git commit -m "test: extração de CNPJ do conteúdo (hífen, intermediários, checksum)"
```

---

## Task 7: Testes de configuração e predefinições (`carregar_config`)

**Files:**
- Create: `tests/test_config.py`

**Interfaces:**
- Consumes: `app.carregar_config() -> dict`, `app.salvar_config(config)`,
  `app.pasta_base()` (monkeypatchado para um tempdir), `app.DEFAULTS_CONFIG`.

Estrutura do config: `{"predefinicao_ativa": str, "predefinicoes": {chave:
{...}}, "tema": str}`. Chaves de perfil: `fedcorp`, `ff`, `notas_diversas`.
Os testes trocam `app.pasta_base` por um tempdir para nunca tocar no
`config.json` real.

- [ ] **Step 1: Escrever os testes**

Create `tests/test_config.py`:

```python
import json
import os
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
```

- [ ] **Step 2: Rodar e confirmar que passa**

Run: `python tests/test_config.py -v`
Expected: 5 testes OK.

- [ ] **Step 3: Rodar a suíte inteira pelo runner**

Run: `rodar_testes.bat` (duplo clique) OU
`python -m unittest discover -s tests -p "test_*.py" -v`
Expected: todos os testes OK (~34 no total).

- [ ] **Step 4: Commit**

```bash
git add tests/test_config.py
git commit -m "test: configuração e migração de predefinições"
```

---

## Task 8: Documentação — atualizar CLAUDE.md e .gitignore

**Files:**
- Modify: `CLAUDE.md`
- Modify: `.gitignore`

- [ ] **Step 1: Ignorar caches de teste**

Em `.gitignore`, adicione ao final:

```
.pytest_cache/
tests/__pycache__/
```

- [ ] **Step 2: Documentar os testes no CLAUDE.md**

Em `CLAUDE.md`, na seção "Convenções do código", adicione:

```markdown
- Testes: `python -m unittest discover -s tests -p "test_*.py"` (ou duplo
  clique em `rodar_testes.bat`). Usam `tests/cadastro_teste.py` (9 condomínios
  fixos) e fixtures de texto em `tests/dados/`, nunca a planilha real. Cobrem
  identificação por nome, extração de CNPJ, validação, desempate, nome de saída
  e config/migração. A regra de desempate vive em `desempatar_por_cadastro`
  (extraída do loop justamente para ser testável). Regenerar fixtures:
  `python tests/_gerar_fixtures.py` (precisa dos PDFs-fonte, fora do repo).
```

- [ ] **Step 3: Rodar a suíte uma última vez**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: todos OK.

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md .gitignore
git commit -m "docs: registra a suíte de testes no CLAUDE.md"
```

---

## Verificação final

Ao terminar todas as tasks:
- `rodar_testes.bat` passa em tudo (~34 testes).
- Reverter qualquer uma das três correções da v6.1/v6.2 faz ao menos um teste
  falhar (demonstrado na Task 3 Step 3 para a trava de ambiguidade).
- `processamento.log` continua com o mesmo formato (desempate preservado).
- Nenhuma dependência nova; nenhuma alteração no app além de
  `desempatar_por_cadastro`.
