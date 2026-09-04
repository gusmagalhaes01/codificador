# Protocolo dos Correios do sistema antigo ("telnet") — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fazer a aba 3 (Protocolos dos Correios) reconhecer, contar e cobrar o protocolo do sistema antigo da Imodata, que chega misturado com o protocolo novo na mesma pasta.

**Architecture:** Um discriminador decide o formato de cada arquivo pelo marcador no texto. O telnet ganha um caminho de extração próprio em `logica.py`, feito de funções puras sobre texto: extração por leitura, votação por maioria entre três leituras de OCR, e uma conferência que preenche e marca a dúvida em vez de barrar. O dict devolvido tem as chaves `codigo` e `condominio`, que é tudo que `linha_planilha_protocolo` e `_carimbar_protocolo` consomem — então nada a jusante (carimbo, planilha, painel, bloco do Paybox) precisa mudar.

**Tech Stack:** Python 3.14, `re`, `collections`, `difflib.SequenceMatcher`, `unittest`. OCR pelo `winocr` via as funções já existentes `extrair_texto_escaneado` e `extrair_texto_ocr_regiao`.

## Global Constraints

- **Spec:** `docs/superpowers/specs/2026-09-03-protocolo-telnet-design.md`. Toda decisão de regex e limiar está justificada lá com o número medido.
- **Separação de camadas:** `logica.py` não importa nada de interface. Toda a lógica nova vive lá; `identificacao_por_cnpj_6_0.py` só roteia e lê.
- **Comentários e nomes de variáveis em português**, como todo o resto do projeto.
- **Dinheiro nunca em float** — o telnet não calcula valor, só devolve `total_impresso` (int); o valor continua saindo de `valor_protocolo(unidades, tarifa)`, que já usa `Decimal`.
- **Fixtures de teste são strings inline no arquivo de teste**, como em `tests/test_protocolo_contagem.py` — nunca PDF real no repositório, nunca nome de morador real.
- **Rodar a suíte:** `python -m unittest discover -s tests -p "test_*.py"`. Estado inicial: **323 testes passando**.
- **Não alterar** `extrair_dados_protocolo_correio`, `conferir_contagem_protocolo` nem `extrair_codigo_protocolo_correio`. O caminho do protocolo novo fica intocado.
- **Limiares fixados pelo spec:** nome confirma a partir de **0.90**; total exige **2** leituras concordando; código aceita **1** voto mas não aceita empate.

---

### Task 1: Discriminador entre os dois formatos

Esta é a primeira tarefa porque é o **único risco do spec que não foi medido**. Errar o valor o usuário corrige na tela; classificar errado faz o programa aplicar a extração errada e produzir um resultado plausível pelo motivo errado, sem nada visivelmente estranho para a conferência humana pegar.

**Files:**
- Modify: `logica.py` (constantes junto das do protocolo novo, por volta da linha 713-723; função nova logo depois de `extrair_codigo_protocolo_correio`, que termina por volta da linha 790)
- Create: `tests/test_protocolo_telnet.py`

**Interfaces:**
- Consumes: `RE_MARCADOR_PROTOCOLO` (já existe em `logica.py:722`)
- Produces: `formato_do_protocolo(texto) -> "telnet" | "novo" | None`

- [ ] **Step 1: Escrever o teste que falha**

Criar `tests/test_protocolo_telnet.py` com o cabeçalho de import no mesmo molde de `tests/test_protocolo_contagem.py`:

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

#  Texto como o winocr devolve de verdade: página inteira numa linha só e
#  com as colunas fora de ordem. Os erros de leitura aqui são os REAIS
#  observados nos sete arquivos de referência — "EWIADO" no lugar de
#  "ENVIADO" e "D ÀS" no lugar de "COD.". Nomes de condomínio são os do
#  cadastro de teste.
TELNET_OK = (
    "IMODÀTÀ * PROTOCOLO CORRESPONDENCIA c CORREIO NORMAL DATA: 11/08/2026 "
    "*** EDF: REF ; LJ-01 LJ-25 AP-308 KLOSTERS 07/2026 R TESTE 100 "
    "PB608111.121 LJ-02 LJ-08 AP-204 AP-207 "
    "TOTAL EWIADO PELO CORREIO 3 o COD.. D ÀS 1.0004.8)"
)

#  Protocolo novo, no formato que a aba 3 já lê hoje.
PROTOCOLO_NOVO = (
    "W700A KLOSTERS (10004) Protocolo de Recebimento de Documento Unidade "
    "401 - Fulano de Tal 402 - Beltrano Silva 403 - Cicrano Souza "
    "Listando 3 unidades Assinatura Entrega Correio Correio Correio 1 de 1"
)

#  "Meus Correios": terceiro formato que aparece na mesma pasta e que este
#  spec deliberadamente NÃO trata (SEDEX, cobrado por outro critério).
#  Não pode ser confundido com nenhum dos dois.
MEUS_CORREIOS = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10004-KLOSTERS Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Total: 1"
)


class TestFormatoDoProtocolo(unittest.TestCase):
    def test_reconhece_o_telnet(self):
        self.assertEqual(app.formato_do_protocolo(TELNET_OK), "telnet")

    def test_reconhece_o_protocolo_novo(self):
        self.assertEqual(app.formato_do_protocolo(PROTOCOLO_NOVO), "novo")

    def test_meus_correios_nao_e_nenhum_dos_dois(self):
        self.assertIsNone(app.formato_do_protocolo(MEUS_CORREIOS))

    def test_texto_vazio_e_none(self):
        self.assertIsNone(app.formato_do_protocolo(""))
        self.assertIsNone(app.formato_do_protocolo(None))

    def test_acento_e_caixa_nao_atrapalham(self):
        #  O OCR ora come o acento de "CORRESPONDÊNCIA", ora não.
        self.assertEqual(
            app.formato_do_protocolo("protocolo correspondência correio"),
            "telnet")

    def test_os_dois_marcadores_juntos_viram_ambiguo(self):
        #  Pior caso possível: aplicar a extração errada produz resultado
        #  plausível pelo motivo errado, e a conferência humana não tem o
        #  que estranhar na tela. Ambíguo vira None -> pendente.
        self.assertIsNone(app.formato_do_protocolo(TELNET_OK + " " + PROTOCOLO_NOVO))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_protocolo_telnet -v
```

Esperado: FAIL com `AttributeError: module 'logica' has no attribute 'formato_do_protocolo'`.

- [ ] **Step 3: Implementar**

Em `logica.py`, logo abaixo do bloco de constantes do protocolo novo (depois de `RE_MARCADOR_PROTOCOLO`, por volta da linha 723):

```python
#  --- Protocolo do sistema antigo da Imodata ("telnet") ---
#  Cabeçalho: "*** IMODATA * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL".
#  O OCR estraga bastante coisa nessa linha (sai "IMODÀTÀ", sai um "c"
#  solto no meio), mas essas duas palavras juntas não falharam em nenhuma
#  das 21 leituras medidas. O acento de CORRESPONDÊNCIA é opcional porque
#  ora o OCR o come, ora não.
RE_MARCADOR_TELNET = re.compile(r"PROTOCOLO\s+CORRESPOND[EÊ]NCIA", re.IGNORECASE)
```

E, logo depois de `extrair_codigo_protocolo_correio` terminar:

```python
def formato_do_protocolo(texto):
    """
    Qual dos dois formatos de protocolo dos Correios é este documento:
    "telnet" (sistema antigo da Imodata), "novo" (o do Superlógica, que a
    aba 3 já lê) ou None.

    Os dois chegam MISTURADOS na mesma pasta, então a decisão é tomada
    arquivo a arquivo.

    Com os dois marcadores presentes devolve None de propósito. É o pior
    caso possível: aplicar a extração errada produziria um resultado
    plausível pelo motivo errado, e não há nada visivelmente estranho para
    a conferência humana pegar na tela. Ambíguo tem que virar pendente.
    """
    texto = texto or ""
    e_telnet = bool(RE_MARCADOR_TELNET.search(texto))
    e_novo = bool(RE_MARCADOR_PROTOCOLO.search(texto))
    if e_telnet and e_novo:
        return None
    if e_telnet:
        return "telnet"
    if e_novo:
        return "novo"
    return None
```

- [ ] **Step 4: Rodar o teste e confirmar que passa**

```bash
python -m unittest tests.test_protocolo_telnet -v
```

Esperado: 6 testes PASS.

- [ ] **Step 5: Validar contra o lote misto real**

> **Já executado durante o planejamento, contra `C:\Users\Desktop\Desktop\Correios II\teste` (15 protocolos novos + 7 telnet): 21 arquivos legíveis, 21 classificados corretamente, ZERO classificações erradas.** Dois pontos aprendidos ali estão embutidos no plano:
>
> 1. **O telnet precisa de segunda chance.** O `telnet-010` não teve o marcador lido na leitura de página inteira a 300 DPI, mas foi reconhecido no recorte do topo a 400. Sem isso, ele cairia no caminho do protocolo novo e viraria pendente. A segunda chance está na Task 4.
> 2. **O `doc16004820260821155206.pdf` não é reconhecido** — mas isso é **comportamento pré-existente**, não regressão: `extrair_dados_protocolo_correio` também não acha o marcador nele nem a 300 nem a 400 DPI. Ele já é pendente hoje.
>
> Refaça mesmo assim: o script abaixo é a rede que pega uma regressão futura no marcador.

Criar `tests/_validar_discriminador.py` (arquivo de uso manual, no mesmo espírito de `tests/_importar_id_sl.py` — prefixo `_` para o `unittest discover` não o coletar):

```python
# -*- coding: utf-8 -*-
"""
Confere o discriminador contra os PDFs reais. Uso manual, reexecutável.

Não entra na suíte automatizada (prefixo _): depende de pastas que não
estão no repositório. Rode antes de liberar qualquer mudança no
discriminador.

    python tests/_validar_discriminador.py
"""
import os
import sys
import glob

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import logica

PASTA = r"C:\Users\Desktop\Desktop\Correios II\teste"

#  Pré-existente, e NÃO é falha do discriminador: extrair_dados_protocolo_
#  correio também não acha o marcador neste arquivo, nem a 300 nem a 400
#  DPI. Ele já é pendente hoje, antes desta funcionalidade existir.
ILEGIVEIS_CONHECIDOS = {"doc16004820260821155206.pdf"}

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)


def classificar(caminho):
    """Como a aba 3 classifica: leitura normal e, se não decidir, uma
    segunda chance no recorte do topo — onde vive o cabeçalho dos dois
    formatos. Sem ela o telnet-010 não é reconhecido."""
    texto = logica.extrair_texto_pdf(caminho)
    if len(texto.strip()) < logica.LIMITE_TEXTO_MINIMO:
        texto = logica.extrair_texto_escaneado(caminho, dpi=300)
    formato = logica.formato_do_protocolo(texto)
    if formato is None:
        try:
            formato = logica.formato_do_protocolo(
                logica.extrair_texto_ocr_regiao(caminho, (0, 0, 1, 0.35), dpi=400))
        except Exception:
            pass
    return formato


erros = 0
for caminho in sorted(glob.glob(os.path.join(PASTA, "*.pdf"))):
    nome = os.path.basename(caminho)
    if os.path.getsize(caminho) == 0:
        print(f"  {nome:34} (0 byte, pulado)")
        continue

    esperado = "telnet" if "telnet" in nome.lower() else "novo"
    obtido = classificar(caminho)

    if obtido is None and nome in ILEGIVEIS_CONHECIDOS:
        print(f"  {nome:34} None     (ilegível conhecido, pré-existente)")
        continue

    #  O que NÃO pode acontecer é trocar um formato pelo outro. None é
    #  falha segura: o arquivo vira pendente em vez de receber a extração
    #  errada.
    if obtido is not None and obtido != esperado:
        erros += 1
        print(f"  {nome:34} obtido={obtido:8} esperado={esperado:8} <<< TROCOU O FORMATO")
    elif obtido is None:
        print(f"  {nome:34} None     esperado={esperado:8} <<< nao reconheceu (falha segura)")
    else:
        print(f"  {nome:34} {obtido:8} ok")

print(f"\n{'NENHUM FORMATO TROCADO' if not erros else f'{erros} FORMATOS TROCADOS'}")
sys.exit(1 if erros else 0)
```

Rodar:

```bash
python tests/_validar_discriminador.py
```

Esperado: **NENHUM FORMATO TROCADO**, com 15 arquivos como `"novo"` (menos o de 0 byte e o ilegível conhecido), 7 como `"telnet"`.

**Se algum formato for trocado, PARE.** Não siga para a Task 2. Trocar um formato pelo outro é o único desfecho que a conferência humana não pega — e é o motivo de esta tarefa vir primeiro.

- [ ] **Step 6: Commit**

```bash
git add logica.py tests/test_protocolo_telnet.py tests/_validar_discriminador.py
git commit -m "Discriminador entre o protocolo novo e o do sistema antigo (telnet)

Os dois formatos chegam misturados na mesma pasta, entao a decisao e por
arquivo. Com os DOIS marcadores presentes devolve None de proposito:
aplicar a extracao errada produz resultado plausivel pelo motivo errado, e
a conferencia humana nao tem o que estranhar na tela.

Validado contra o lote misto real: 14 protocolos novos, 7 telnet e 6 'Meus
Correios' (terceiro formato, fora do escopo) classificados corretamente."
```

---

### Task 2: Extração de uma leitura do telnet

**Files:**
- Modify: `logica.py` (constantes junto de `RE_MARCADOR_TELNET`; função depois de `formato_do_protocolo`)
- Modify: `tests/test_protocolo_telnet.py`

**Interfaces:**
- Consumes: nada da Task 1 além do arquivo de teste já criado
- Produces: `extrair_dados_protocolo_telnet(texto) -> {"codigos": list[str], "total": int | None}`

- [ ] **Step 1: Escrever o teste que falha**

Acrescentar a `tests/test_protocolo_telnet.py`, antes do `if __name__`:

```python
class TestExtracaoDeUmaLeitura(unittest.TestCase):
    def test_codigo_sai_pelo_formato_mesmo_com_o_rotulo_corrompido(self):
        #  "D ÀS" é o que o OCR devolveu no lugar de "COD." num arquivo
        #  real. Ancorado no rótulo o código saía em 4 de 7; pelo formato,
        #  em 7 de 7.
        dados = app.extrair_dados_protocolo_telnet(TELNET_OK)
        self.assertIn("10004", dados["codigos"])

    def test_total_lido_com_ENVIADO_corrompido(self):
        #  O OCR devolve "EWIADO" e "EFvTIADO"; "PELO CORREIO" nunca falhou.
        dados = app.extrair_dados_protocolo_telnet(TELNET_OK)
        self.assertEqual(dados["total"], 3)

    def test_total_com_ENVIADO_escrito_de_outro_jeito(self):
        texto = "TOTAL EFvTIADO PELO CORREIO 12 o COD. . 1.0004.7)"
        self.assertEqual(app.extrair_dados_protocolo_telnet(texto)["total"], 12)

    def test_folga_curta_nao_captura_numero_distante(self):
        #  Caso real: com folga de 20 não-dígitos o regex pulava o número
        #  certo e capturava outro do canto da folha -- leu 38 no lugar de
        #  3, num condomínio de três unidades (R$ 146,30 em vez de R$ 11,55).
        texto = "TOTAL ENVIADO PELO CORREIO VALOR VAL x R EURICO 38"
        self.assertIsNone(app.extrair_dados_protocolo_telnet(texto)["total"])

    def test_o_digito_depois_do_codigo_e_ignorado(self):
        #  "1.1122.8)" -> o ".8)" não faz parte do código.
        dados = app.extrair_dados_protocolo_telnet("COD.: 1.1122.8)")
        self.assertEqual(dados["codigos"], ["11122"])

    def test_virgula_no_lugar_do_ponto(self):
        dados = app.extrair_dados_protocolo_telnet("COD.: 1,1122.8)")
        self.assertEqual(dados["codigos"], ["11122"])

    def test_sem_codigo_nem_total_devolve_vazio(self):
        dados = app.extrair_dados_protocolo_telnet("texto qualquer sem nada")
        self.assertEqual(dados["codigos"], [])
        self.assertIsNone(dados["total"])

    def test_texto_vazio_nao_estoura(self):
        dados = app.extrair_dados_protocolo_telnet(None)
        self.assertEqual(dados["codigos"], [])
        self.assertIsNone(dados["total"])
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_protocolo_telnet.TestExtracaoDeUmaLeitura -v
```

Esperado: FAIL com `AttributeError: module 'logica' has no attribute 'extrair_dados_protocolo_telnet'`.

- [ ] **Step 3: Implementar**

Em `logica.py`, junto de `RE_MARCADOR_TELNET`:

```python
#  O código sai pelo FORMATO, não pelo rótulo "COD.": o OCR estraga o
#  rótulo (leu "DAS", leu "D ÀS") muito mais do que os dígitos. Ancorado no
#  rótulo o código saía em 4 dos 7 arquivos de referência; pelo formato, em
#  7 de 7. Todos os 772 códigos do cadastro têm 5 dígitos e começam em 1
#  (faixa 10002-11242), então "1.DDDD" é distintivo sozinho. O que vier
#  depois (".8)", ".7)") é descartado, por instrução do usuário.
RE_CODIGO_TELNET = re.compile(r"\b(1)\s*[.,]\s*(\d{4})\b")

#  Âncora em "PELO CORREIO", NUNCA em "ENVIADO" — o OCR devolve "EWIADO" e
#  "EFvTIADO". A folga até o número é de 6 não-dígitos, e isso é essencial:
#  com folga de 20 o regex pulava o número certo e capturava um de outro
#  canto da folha (leu 38 no lugar de 3 num condomínio de três unidades).
RE_TOTAL_TELNET = re.compile(r"PELO\s+CORREIO\D{0,6}(\d{1,4})\b", re.IGNORECASE)
```

E a função, depois de `formato_do_protocolo`:

```python
def extrair_dados_protocolo_telnet(texto):
    """
    O que UMA leitura de OCR do telnet oferece: os códigos que casam com o
    formato e o total impresso.

    Devolve TODOS os códigos encontrados, sem escolher nem validar contra o
    cadastro. Quem decide é `apurar_protocolo_telnet`, cruzando várias
    leituras — uma leitura sozinha não decide nada aqui, porque nenhuma
    configuração de OCR lê sozinha os sete arquivos de referência.
    """
    texto = texto or ""
    codigos = [inicio + resto for inicio, resto in RE_CODIGO_TELNET.findall(texto)]
    achado = RE_TOTAL_TELNET.search(texto)
    return {
        "codigos": codigos,
        "total": int(achado.group(1)) if achado else None,
    }
```

- [ ] **Step 4: Rodar o teste e confirmar que passa**

```bash
python -m unittest tests.test_protocolo_telnet -v
```

Esperado: 14 testes PASS (6 da Task 1 + 8 desta).

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_telnet.py
git commit -m "Extrai codigo e total de uma leitura do protocolo telnet

O codigo sai do formato 1.DDDD, nao do rotulo COD.: o OCR estraga o rotulo
(leu DAS, leu D AS) muito mais que os digitos -- 4/7 contra 7/7.

O total ancora em PELO CORREIO, nao em ENVIADO (sai EWIADO, EFvTIADO), com
folga de 6 nao-digitos. A folga curta e o ponto: com 20, o regex pulava o
numero certo e capturava um do canto da folha -- 38 no lugar de 3."
```

---

### Task 3: Votação entre leituras e regra de aceite

**Files:**
- Modify: `logica.py` (adicionar `import collections` no topo; constantes e três funções)
- Modify: `tests/test_protocolo_telnet.py`

**Interfaces:**
- Consumes: `extrair_dados_protocolo_telnet` (Task 2), `_codigos_do_cadastro` (`logica.py:1037`), `normalizar_texto_busca` (`logica.py:1020`), `SequenceMatcher` (já importado em `logica.py:21`)
- Produces:
  - `apurar_protocolo_telnet(leituras, cadastro) -> dict | None` com as chaves `formato`, `codigo`, `condominio`, `total_impresso`, `nome_confere`
  - `conferir_contagem_telnet(dados) -> (bool, str)`

- [ ] **Step 1: Escrever o teste que falha**

Acrescentar a `tests/test_protocolo_telnet.py`. Usa `tests/cadastro_teste.py`, o mesmo cadastro fixo dos outros testes — nele `KLOSTERS` é o código `10004` (chave `01195716000154`) e `SAN REMO` é o `10002`:

```python
from cadastro_teste import CADASTRO_TESTE as CADASTRO


def _leitura(codigo_pontuado, total_texto, nome=""):
    """Monta o texto de uma leitura de OCR do telnet."""
    return (f"IMODÀTÀ * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL EDF: {nome} "
            f"AP-101 AP-102 TOTAL EWIADO PELO CORREIO {total_texto} o COD.. "
            f"D ÀS {codigo_pontuado})")


class TestVotacao(unittest.TestCase):
    def test_maioria_corrige_a_leitura_ruim(self):
        #  Caso real do arquivo 011: as três leituras deram 38, 3 e 3. Sem
        #  votar, "R$ 146,30" chegaria à tela num condomínio de três
        #  unidades. Este é o teste que justifica a votação existir.
        leituras = [_leitura("1.0004.8", "38", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertEqual(dados["total_impresso"], 3)
        self.assertEqual(dados["codigo"], "10004")

    def test_total_com_uma_leitura_so_nao_basta(self):
        #  O total vira dinheiro; uma leitura solitária não sustenta isso.
        leituras = [_leitura("1.0004.8", "7", "KLOSTERS"),
                    _leitura("1.0004.8", "", "KLOSTERS"),
                    _leitura("1.0004.8", "", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["total_impresso"])

    def test_tres_totais_diferentes_nao_elegem_nenhum(self):
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "5", "KLOSTERS"),
                    _leitura("1.0004.8", "9", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["total_impresso"])

    def test_codigo_aceita_um_voto_so(self):
        #  Ao contrário do total, o código não vira valor sozinho — ele
        #  ainda passa pela conferência do nome, que marca a linha.
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("", "3", "KLOSTERS"),
                    _leitura("", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertEqual(dados["codigo"], "10004")

    def test_empate_entre_codigos_nao_escolhe_nenhum(self):
        #  Escolher por ordem de chegada seria arbitrário, e um dígito
        #  errado tende a dar OUTRO condomínio real.
        leituras = [_leitura("1.0004.8", "3", ""),
                    _leitura("1.0002.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["codigo"])

    def test_codigo_fora_do_cadastro_e_descartado(self):
        leituras = [_leitura("1.9999.8", "3", ""),
                    _leitura("1.9999.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertIsNone(dados["codigo"])

    def test_sem_leitura_nenhuma_devolve_none(self):
        self.assertIsNone(app.apurar_protocolo_telnet([], CADASTRO))
        self.assertIsNone(app.apurar_protocolo_telnet(["", None], CADASTRO))


class TestConferenciaDoNome(unittest.TestCase):
    def test_nome_no_texto_confirma_o_codigo(self):
        leituras = [_leitura("1.0004.8", "3", "KLOSTERS"),
                    _leitura("1.0004.8", "3", "KLOSTERS")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertTrue(dados["nome_confere"])
        self.assertEqual(dados["condominio"], "KLOSTERS")

    def test_nome_ausente_nao_confirma(self):
        leituras = [_leitura("1.0004.8", "3", ""),
                    _leitura("1.0004.8", "3", "")]
        dados = app.apurar_protocolo_telnet(leituras, CADASTRO)
        self.assertFalse(dados["nome_confere"])

    def test_nome_longe_do_rotulo_ainda_confirma(self):
        #  O winocr devolve a página numa linha só com as colunas fora de
        #  ordem: o "EDF:" e o valor não ficam adjacentes. Ancorado no
        #  rótulo o nome saía em 1 de 7; varrendo o texto, em 6 de 7.
        texto = ("IMODÀTÀ * PROTOCOLO CORRESPONDENCIA EDF : REF; AP-104 "
                 "TOTAL KLOSTERS 07/2026 PELO CORREIO 3 o COD.. 1.0004.8)")
        dados = app.apurar_protocolo_telnet([texto, texto], CADASTRO)
        self.assertTrue(dados["nome_confere"])


class TestAceiteTelnet(unittest.TestCase):
    def test_tudo_concordando_aceita_sem_observacao(self):
        dados = {"codigo": "10004", "total_impresso": 3, "nome_confere": True}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertTrue(aceito)
        self.assertEqual(observacao, "")

    def test_nome_nao_confirmado_PREENCHE_e_marca(self):
        #  Deliberadamente mais frouxo que o protocolo novo: o usuário
        #  confere imagem por imagem. Uma versão anterior barrava aqui e
        #  mandava para pendente um arquivo cujo código e total estavam
        #  ambos corretos.
        dados = {"codigo": "10004", "total_impresso": 3, "nome_confere": False}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertTrue(aceito)
        self.assertEqual(observacao, "Nome não confirmado")

    def test_sem_total_nao_aceita(self):
        dados = {"codigo": "10004", "total_impresso": None, "nome_confere": True}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertFalse(aceito)
        self.assertIn("TOTAL ENVIADO PELO CORREIO", observacao)

    def test_sem_codigo_nao_aceita(self):
        dados = {"codigo": None, "total_impresso": 3, "nome_confere": False}
        aceito, observacao = app.conferir_contagem_telnet(dados)
        self.assertFalse(aceito)
        self.assertIn("Código", observacao)

    def test_dados_none_nao_estoura(self):
        aceito, observacao = app.conferir_contagem_telnet(None)
        self.assertFalse(aceito)
        self.assertTrue(observacao)
```

- [ ] **Step 2: Rodar o teste e confirmar que falha**

```bash
python -m unittest tests.test_protocolo_telnet.TestVotacao -v
```

Esperado: FAIL com `AttributeError: module 'logica' has no attribute 'apurar_protocolo_telnet'`.

- [ ] **Step 3: Implementar**

Em `logica.py`, adicionar ao bloco de imports do topo (por ordem alfabética, antes de `import copy`):

```python
import collections
```

Junto das outras constantes do telnet:

```python
#  Score mínimo para o nome impresso confirmar o código. Não é delicado:
#  nos 7 arquivos medidos os confirmados deram 1.00 e o único não
#  confirmado deu 0.79 — qualquer corte entre 0.80 e 0.99 daria o mesmo.
LIMIAR_NOME_TELNET = 0.90

#  O total só é aceito com pelo menos duas leituras concordando: é o número
#  que vira dinheiro cobrado. O código se contenta com uma (ver
#  apurar_protocolo_telnet).
VOTOS_MINIMOS_TOTAL_TELNET = 2
```

E as três funções, depois de `extrair_dados_protocolo_telnet`:

```python
def _mais_votado(valores, minimo=1):
    """
    Valor mais frequente da lista, ou None.

    Devolve None em três casos: lista vazia, o mais votado não alcança
    `minimo` votos, ou EMPATE no topo. O empate é o que mais importa —
    escolher por ordem de chegada seria arbitrário, e aqui o número vira
    dinheiro cobrado ou condomínio carimbado.
    """
    if not valores:
        return None
    contagem = collections.Counter(valores).most_common()
    valor, votos = contagem[0]
    if votos < minimo:
        return None
    if len(contagem) > 1 and contagem[1][1] == votos:
        return None
    return valor


def _melhor_score_do_nome(texto_normalizado, nome_normalizado):
    """
    Maior similaridade do nome em qualquer posição do texto.

    Varre o texto inteiro em vez de procurar depois do rótulo "EDF:" porque
    o winocr devolve a página numa linha só e com as colunas fora de ordem:
    o rótulo e o valor não ficam adjacentes (sai "EDF : REF; AP-104 TOTAL
    CHATEAU FONTAINEBLEA"). Ancorado no rótulo, o nome saía em 1 dos 7
    arquivos de referência; varrendo, em 6 de 7 com score 1.00.

    Compara só contra UM nome — o do código já eleito —, então o custo é de
    algumas centenas de comparações, não das 772 do cadastro inteiro.
    """
    if not texto_normalizado or not nome_normalizado:
        return 0.0
    janela = len(nome_normalizado)
    melhor = 0.0
    for inicio in range(0, max(1, len(texto_normalizado) - janela + 1)):
        score = SequenceMatcher(
            None, nome_normalizado,
            texto_normalizado[inicio:inicio + janela]).ratio()
        if score > melhor:
            melhor = score
    return melhor


def apurar_protocolo_telnet(leituras, cadastro):
    """
    Cruza várias leituras de OCR do MESMO telnet e devolve o que a aba 3
    consome. `None` quando não há leitura nenhuma.

    O dict tem `codigo` e `condominio` porque é só isso que
    `linha_planilha_protocolo` e `_carimbar_protocolo` consultam — assim
    nada a jusante (carimbo, planilha, painel, bloco do Paybox) precisa
    saber que existe um segundo formato.

    A votação existe por um caso concreto: um arquivo cujas leituras deram
    38, 3 e 3. Sem votar, "R$ 146,30" chegaria à tela num condomínio de
    três unidades — o tipo de número que passa despercebido num lote de 100.
    """
    leituras = [t for t in (leituras or []) if t]
    if not leituras:
        return None

    #  Índice montado UMA vez, fora do laço: são 772 condomínios.
    codigos_validos = _codigos_do_cadastro(cadastro or {})

    codigos, totais = [], []
    for texto in leituras:
        parcial = extrair_dados_protocolo_telnet(texto)
        codigos.extend(c for c in parcial["codigos"] if c in codigos_validos)
        if parcial["total"] is not None:
            totais.append(parcial["total"])

    #  O código se contenta com um voto: ele não vira valor sozinho, ainda
    #  passa pela conferência do nome. O total exige dois.
    codigo = _mais_votado(codigos)
    total = _mais_votado(totais, minimo=VOTOS_MINIMOS_TOTAL_TELNET)

    nome, confere = "", False
    if codigo:
        documento = codigos_validos.get(codigo)
        registro = (cadastro or {}).get(documento) or {}
        nome = registro.get("nome") or ""
        confere = _melhor_score_do_nome(
            normalizar_texto_busca(" ".join(leituras)),
            normalizar_texto_busca(nome)) >= LIMIAR_NOME_TELNET

    return {
        "formato": "telnet",
        "codigo": codigo,
        "condominio": nome,
        "total_impresso": total,
        "nome_confere": confere,
    }


def conferir_contagem_telnet(dados):
    """
    Decide se a linha do telnet nasce preenchida. Devolve
    (aceito, observacao).

    Deliberadamente mais frouxo que `conferir_contagem_protocolo`: aqui o
    usuário confere imagem por imagem, então o desenho é PREENCHER e MARCAR
    a dúvida, não barrar. Uma versão anterior barrava a linha sem
    confirmação de nome e mandava para pendente um arquivo cujo código e
    total estavam ambos corretos.

    A marcação continua valendo porque numa pilha de 100 todo mundo olha
    com mais atenção o que está sinalizado: ela diz ONDE o programa ficou
    em dúvida, em vez de esconder a dúvida atrás de um valor com cara de
    certo.
    """
    dados = dados or {}
    if not dados.get("codigo"):
        return False, "Código do condomínio não encontrado no cadastro"
    if dados.get("total_impresso") is None:
        return False, 'Não foi possível ler o total ("TOTAL ENVIADO PELO CORREIO")'
    if not dados.get("nome_confere"):
        return True, "Nome não confirmado"
    return True, ""
```

- [ ] **Step 4: Rodar o teste e confirmar que passa**

```bash
python -m unittest tests.test_protocolo_telnet -v
```

Esperado: 29 testes PASS. Depois a suíte inteira:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: 352 testes, OK (323 anteriores + 29 novos).

- [ ] **Step 5: Commit**

```bash
git add logica.py tests/test_protocolo_telnet.py
git commit -m "Votacao entre leituras e regra de aceite do telnet

Tres leituras de OCR, valor por maioria. A votacao existe por um caso
concreto: 38, 3, 3 -> 3. Sem ela, R$ 146,30 chegaria a tela num condominio
de tres unidades.

O total exige 2 votos (vira dinheiro); o codigo se contenta com 1, porque
ainda passa pela conferencia do nome. Empate no topo nao elege ninguem --
escolher por ordem de chegada seria arbitrario.

O nome e procurado no texto inteiro, nao depois do EDF:, porque o winocr
embaralha as colunas: 1/7 ancorado no rotulo contra 6/7 varrendo.

Aceite frouxo de proposito: preenche e MARCA a duvida em vez de barrar,
porque o usuario confere imagem por imagem."
```

---

### Task 4: Integração na aba 3

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (imports por volta da linha 77; método novo perto de `_ler_texto_protocolo`, que está na linha 2868; roteamento dentro de `_processar_protocolos_em_thread`, por volta da linha 2919)

**Interfaces:**
- Consumes: `formato_do_protocolo`, `apurar_protocolo_telnet`, `conferir_contagem_telnet` (Tasks 1-3); `extrair_texto_escaneado` e `extrair_texto_ocr_regiao` (já importados)
- Produces: nada para tarefas seguintes — é a ponta do fluxo

- [ ] **Step 1: Acrescentar os imports**

Em `identificacao_por_cnpj_6_0.py`, na linha 77, que hoje é:

```python
    extrair_dados_protocolo_correio, conferir_contagem_protocolo,
```

Trocar por:

```python
    extrair_dados_protocolo_correio, conferir_contagem_protocolo,
    formato_do_protocolo, apurar_protocolo_telnet, conferir_contagem_telnet,
```

- [ ] **Step 2: Acrescentar o leitor de três passadas**

Logo depois do método `_ler_texto_protocolo` (que termina na linha 2879 com `return extrair_texto_escaneado(caminho, dpi=dpi), True`), inserir:

```python
    #  Recortes e qualidades usados nas três leituras do telnet. Foram
    #  escolhidos por medição: nenhum sozinho lê os sete arquivos de
    #  referência, e os três juntos leem. Os recortes do topo existem
    #  porque os comprovantes térmicos colados na metade de baixo da folha
    #  são ruído que atrapalha a leitura do cabeçalho.
    LEITURAS_TELNET = (
        (None, 300),              # página inteira
        ((0, 0, 1, 0.35), 400),   # topo
        ((0, 0, 1, 0.32), 500),   # topo, mais perto
    )

    def _leituras_telnet(self, caminho, texto_inicial):
        """
        As três leituras de OCR de um telnet, para a votação.

        `texto_inicial` é a leitura que o laço já fez (e que serviu para
        classificar o formato); ela entra como a primeira, para não pagar
        OCR duas vezes pelo mesmo recorte.

        Falha de uma configuração não derruba as outras: a votação
        trabalha com o que conseguir. É por isso que cada leitura tem seu
        próprio try.
        """
        leituras = [texto_inicial] if texto_inicial else []
        for regiao, dpi in self.LEITURAS_TELNET[1:]:
            try:
                if regiao is None:
                    leituras.append(extrair_texto_escaneado(caminho, dpi=dpi))
                else:
                    leituras.append(extrair_texto_ocr_regiao(caminho, regiao, dpi=dpi))
            except Exception:
                #  Uma configuração que falhou não tem o que contribuir;
                #  as outras seguem. Nenhuma leitura é o caso tratado em
                #  apurar_protocolo_telnet, que devolve None.
                pass
        return leituras
```

- [ ] **Step 3: Acrescentar o classificador com segunda chance**

Logo depois de `_leituras_telnet`, inserir:

```python
    def _e_telnet(self, caminho, texto):
        """
        Se este arquivo é um protocolo do sistema antigo.

        Dá uma SEGUNDA CHANCE no recorte do topo quando a primeira leitura
        não decide. Isso foi medido: o cabeçalho de um dos sete arquivos de
        referência não sai na leitura de página inteira a 300 DPI, mas sai
        no recorte a 400. Sem a segunda chance ele cairia no caminho do
        protocolo novo e viraria pendente.

        A segunda leitura só acontece quando a primeira devolve None — se
        ela já disse "novo" com confiança, não há por que pagar outro OCR.
        """
        formato = formato_do_protocolo(texto)
        if formato is not None:
            return formato == "telnet"
        try:
            regiao, dpi = self.LEITURAS_TELNET[1]
            return formato_do_protocolo(
                extrair_texto_ocr_regiao(caminho, regiao, dpi=dpi)) == "telnet"
        except Exception:
            #  Não deu para decidir: segue pelo caminho do protocolo novo,
            #  que tem a própria re-tentativa e a própria mensagem de
            #  "não foi possível ler".
            return False
```

- [ ] **Step 4: Rotear no laço de processamento**

Em `_processar_protocolos_em_thread`, o trecho atual (por volta da linha 2917-2919) é:

```python
                dpi = QUALIDADE_LEITURA_PROTOCOLO
                texto, usou_leitor_escaneado = self._ler_texto_protocolo(caminho, dpi)
                dados = extrair_dados_protocolo_correio(texto)
```

Trocar as três linhas por:

```python
                dpi = QUALIDADE_LEITURA_PROTOCOLO
                texto, usou_leitor_escaneado = self._ler_texto_protocolo(caminho, dpi)

                if self._e_telnet(caminho, texto):
                    #  O telnet tem caminho próprio: precisa de três
                    #  leituras e votação, porque nenhuma configuração de
                    #  OCR sozinha lê os sete arquivos de referência — e foi
                    #  a votação que corrigiu um 38 para 3 num condomínio de
                    #  três unidades. Sai daqui com as MESMAS variáveis que
                    #  o resto do laço espera (`dados`, `unidades`,
                    #  `observacao`), e o dict tem `codigo` e `condominio`,
                    #  que é tudo que o carimbo e a planilha consultam.
                    dados = apurar_protocolo_telnet(
                        self._leituras_telnet(caminho, texto), self.cadastro)
                    aceito, observacao = conferir_contagem_telnet(dados)
                    if aceito:
                        unidades = dados["total_impresso"]
                else:
                    dados = extrair_dados_protocolo_correio(texto)
```

E **indentar um nível** todo o bloco que hoje vem depois dessa linha e trata do protocolo novo — da linha `ja_tentou_de_novo = False` até `observacao = motivo`, inclusive —, para que ele passe a viver dentro do `else`. **Não alterar nada desse bloco além da indentação**: os comentários dele descrevem decisões já tomadas (teto de uma re-tentativa por arquivo, só adotar a 2ª leitura quando ela aceita a contagem) e continuam valendo.

O `except Exception as e:` que fecha o `try` (hoje na linha 2990) **não muda** — nem sua indentação, nem seu conteúdo.

Este desenho substituiu uma primeira versão que usava uma exceção-sentinela para pular o bloco do protocolo novo sem re-indentá-lo. A re-indentação é mecânica e revisável; exceção como desvio de fluxo é o tipo de coisa que um revisor reprova com razão, ainda mais dentro de um `try` que também captura erros de verdade.

- [ ] **Step 5: Verificar que a suíte continua passando**

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: 352 testes, OK. Depois, conferir que o módulo importa sem erro de sintaxe — é o que pega uma re-indentação malfeita:

```bash
python -c "import identificacao_por_cnpj_6_0; print('ok')"
```

Esperado: `ok` (só o import; não abre janela).

Por fim, rodar um lote real misto pela interface, na pasta `C:\Users\Desktop\Desktop\Correios II\teste`, e conferir no painel que os 7 telnet aparecem com código e valor, e que os protocolos novos continuam como antes.

- [ ] **Step 6: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Roteia o telnet na aba 3, com tres leituras e votacao

O laco classifica cada arquivo antes de extrair, porque os dois formatos
chegam misturados na mesma pasta. O telnet sai do desvio com as mesmas
variaveis que o resto do laco espera, e o dict tem codigo e condominio --
tudo que o carimbo e a planilha consultam. Nada a jusante muda.

A primeira das tres leituras e a que o laco ja fez para classificar, para
nao pagar OCR duas vezes pelo mesmo recorte."
```

---

### Task 5: Validação de ponta a ponta, documentação e versão

**Files:**
- Create: `tests/_validar_telnet.py`
- Modify: `CLAUDE.md` (seção nova depois de "Contagem dos Protocolos dos Correios (aba 3)")
- Modify: `identificacao_por_cnpj_6_0.py:209` (versão)

**Interfaces:**
- Consumes: tudo das Tasks 1-4
- Produces: nada

- [ ] **Step 1: Escrever o validador contra os 7 arquivos reais**

Criar `tests/_validar_telnet.py`:

```python
# -*- coding: utf-8 -*-
"""
Confere a apuração do telnet contra os 7 PDFs reais de referência.

Uso manual, reexecutável, fora da suíte (prefixo _): depende de uma pasta
que não está no repositório. A verdade abaixo foi lida à mão na folha de
cada documento.

    python tests/_validar_telnet.py
"""
import os
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

PASTA = r"C:\Users\Desktop\Desktop\Correios\telnet"

#  arquivo -> (total impresso, código), conferidos lendo a folha
VERDADE = {
    "procolo telnet-007.pdf": (27, "11122"),
    "procolo telnet-008.pdf": (8,  "11127"),
    "procolo telnet-009.pdf": (10, "11121"),
    "procolo telnet-010.pdf": (8,  "10151"),
    "procolo telnet-011.pdf": (3,  "11100"),
    "procolo telnet-012.pdf": (5,  "11098"),
    "procolo telnet-013.pdf": (1,  "11093"),
}

REGIOES = [(None, 300), ((0, 0, 1, 0.35), 400), ((0, 0, 1, 0.32), 500)]
cadastro = logica.carregar_cadastro(os.path.join(_RAIZ, "cadastro_condominios.xlsx"))

print(f"{'arquivo':26} {'cod':>6} {'esp':>6} {'tot':>5} {'esp':>5} {'nome':>5}  resultado")
print("-" * 82)
codigos_ok = totais_ok = errados = 0
for nome, (total_certo, codigo_certo) in VERDADE.items():
    caminho = os.path.join(PASTA, nome)
    leituras = []
    for regiao, dpi in REGIOES:
        try:
            leituras.append(logica.extrair_texto_escaneado(caminho, dpi=dpi)
                            if regiao is None
                            else logica.extrair_texto_ocr_regiao(caminho, regiao, dpi=dpi))
        except Exception:
            pass
    dados = logica.apurar_protocolo_telnet(leituras, cadastro) or {}
    aceito, observacao = logica.conferir_contagem_telnet(dados)

    cod, tot = dados.get("codigo"), dados.get("total_impresso")
    codigos_ok += (cod == codigo_certo)
    totais_ok += (tot == total_certo)
    #  O que NÃO pode acontecer: preencher um valor errado.
    if aceito and tot != total_certo:
        errados += 1
    if cod is not None and cod != codigo_certo:
        errados += 1
    print(f"{nome[-12:]:26} {str(cod):>6} {codigo_certo:>6} {str(tot):>5} "
          f"{total_certo:>5} {('sim' if dados.get('nome_confere') else 'nao'):>5}  "
          f"{observacao or 'preenche'}")

print(f"\ncodigos corretos: {codigos_ok}/7   totais corretos: {totais_ok}/7")
print(f"VALORES ERRADOS PREENCHIDOS: {errados}   <- tem que ser 0")
sys.exit(1 if errados else 0)
```

- [ ] **Step 2: Rodar e comparar com o spec**

```bash
python tests/_validar_telnet.py
```

Esperado, conforme a tabela "Validação feita" do spec: **códigos 7/7, totais 6/7** (o 008 não tem total lido e sai como pendente), e — o que importa — **VALORES ERRADOS PREENCHIDOS: 0**.

Se `errados` for maior que zero, **PARE e investigue**: é o único desfecho que o desenho inteiro existe para evitar.

- [ ] **Step 3: Documentar no CLAUDE.md**

Acrescentar, logo antes da seção `## Anexo automático no Paybox (aba 3, v6.16.0)`:

```markdown
## Protocolo do sistema antigo — "telnet" (aba 3, v6.19.0)

Segundo formato de protocolo dos Correios, do sistema antigo da Imodata.
Chega **misturado** com o novo na mesma pasta, então `formato_do_protocolo`
decide arquivo a arquivo pelo marcador: `PROTOCOLO CORRESPONDENCIA` (telnet)
ou `Protocolo de Recebimento de Documento` (novo). Com os **dois**
marcadores presentes devolve `None` de propósito — aplicar a extração
errada produz resultado plausível pelo motivo errado, e a conferência
humana não tem o que estranhar na tela.

**Três decisões de extração, todas contra a tentativa óbvia**, e cada uma
medida nos 7 arquivos de referência:

| Ancorar em | Resultado | Por quê |
|---|---|---|
| rótulo `COD.` | 4/7 | o OCR leu `DAS`, `D ÀS` — estraga o rótulo, não os dígitos |
| **formato `1.DDDD`** | **7/7** | todo código tem 5 dígitos e começa em 1 (10002–11242) |
| `ENVIADO` | falha | sai `EWIADO`, `EFvTIADO` |
| **`PELO CORREIO`** | **6/7** | nunca falhou nas 21 leituras |
| rótulo `EDF:` (nome) | 1/7 | o winocr embaralha as colunas: rótulo e valor não ficam juntos |
| **texto inteiro** | **6/7 a 1.00** | varre em vez de ancorar |

**A folga do total é de 6 não-dígitos, nunca 20.** Com folga larga o regex
pulava o número certo e capturava outro do canto da folha: leu `38` no
lugar de `3` num condomínio de três unidades — R$ 146,30 em vez de
R$ 11,55.

**Três leituras de OCR com votação por maioria** (página inteira a 300, topo
a 400 e a 500). Nenhuma configuração sozinha lê os sete. A votação existe
por esse mesmo caso do `38, 3, 3 → 3`. O **total** exige 2 votos, porque
vira dinheiro; o **código** aceita 1, porque ainda passa pela conferência do
nome. **Empate no topo não elege ninguém** — escolher por ordem de chegada
seria arbitrário.

**O cadastro NÃO serve de dígito verificador aqui**, ao contrário do CNPJ.
772 dos ~1.241 números da faixa 10002–11242 são condomínios reais — 62% de
densidade —, então um dígito errado tende a produzir *outro condomínio
existente*. "O código está no cadastro" não prova nada. É por isso que o
nome impresso é cruzado com o do cadastro (limiar 0.90; os confirmados dão
1.00 e o único não confirmado deu 0.79).

**O aceite é deliberadamente mais frouxo que o do protocolo novo**:
`conferir_contagem_telnet` **preenche e marca** em vez de barrar, porque o
usuário confere imagem por imagem. Uma versão anterior barrava a linha sem
confirmação de nome e mandava para pendente um arquivo cujo código e total
estavam ambos corretos. Sem total legível ou sem código, aí sim é pendente.

**Nada a jusante mudou**: o dict tem `codigo` e `condominio`, que é tudo que
`linha_planilha_protocolo` e `_carimbar_protocolo` consultam. Valor,
carimbo lateral, bloco do Paybox, planilhas e painel são os mesmos.

**Um terceiro formato existe e está fora do escopo:** o "Meus Correios"
(`Cód. Condomínio: 10852-VILLA BRANCA`, `Qtde.`, `Classificação`), que são
SEDEX de valor fixo, cobrados por outro critério. Ele é o formato **fácil**
— o OCR lê o código em 6/6 e o nome vem no mesmo campo, servindo de
conferência. Quando for a vez dele, começar daí.

Spec: `docs/superpowers/specs/2026-09-03-protocolo-telnet-design.md`.
Validadores manuais: `tests/_validar_discriminador.py` e
`tests/_validar_telnet.py` (dependem de PDFs fora do repositório).
```

- [ ] **Step 4: Subir a versão**

Em `identificacao_por_cnpj_6_0.py:209`, trocar:

```python
        self.title("Codificador v6.18.1")
```

por:

```python
        self.title("Codificador v6.19.0")
```

- [ ] **Step 5: Rodar a suíte inteira uma última vez**

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Esperado: 352 testes, OK.

- [ ] **Step 6: Commit**

```bash
git add tests/_validar_telnet.py CLAUDE.md identificacao_por_cnpj_6_0.py
git commit -m "Documenta o protocolo telnet e marca a versao 6.19.0

Valida contra os 7 arquivos reais: codigos 7/7, totais 6/7, nenhum valor
errado preenchido. O 008 nao tem total legivel em leitura nenhuma e sai
como pendente, que e o desfecho correto.

O CLAUDE.md registra as tres ancoras que NAO funcionam (rotulo COD.,
palavra ENVIADO, rotulo EDF:) com o numero de cada uma, para ninguem
refazer a tentativa obvia. E registra que o cadastro nao serve de digito
verificador: 62% dos numeros da faixa sao condominios reais."
```

---

## Notas para quem for executar

**O que este plano NÃO cobre, e é sabido:**

- **São 7 amostras, de uma competência só.** O `CLAUDE.md` já registra o precedente: a v6.18.1 nasceu de validar contra um lote único, e "a quebra depende da largura da coluna no lote". Nenhum dos 7 tem mais de uma página, nem prefixo de unidade fora de `LJ`/`AP`/`SL`. Se aparecer telnet multipágina, a leitura por recorte do topo (`0–0.35`) só olha a primeira página — é o primeiro lugar a checar.
- **O `.8)` depois do código** é assumido como ignorável, por instrução do usuário. Se um dia significar algo, a suposição está em `RE_CODIGO_TELNET`.
- **O "Meus Correios" continua caindo como pendente**, e isso é o comportamento correto por ora — ele não é telnet nem protocolo novo, então `formato_do_protocolo` devolve `None`.

**Custo de OCR:** o telnet passa a custar ~3× o de um protocolo novo. A aba 3 já roda em thread com barra de progresso, então isso aparece como lentidão, não como travamento. Só arquivos classificados como telnet pagam esse custo.
