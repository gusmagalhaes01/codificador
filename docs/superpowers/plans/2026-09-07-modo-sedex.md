# Modo SEDEX e leitor do "Meus Correios" — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fazer a aba 3 ler o protocolo do formato "Meus Correios" (sistema Agile) e ganhar um modo de lote que desliga a contagem de unidades, deixando o valor para ser digitado à mão na grade do painel.

**Architecture:** Duas peças ortogonais. O **formato** (novo / telnet / meus correios) decide *como ler* o condomínio e é detectado no documento; o **modo** (normal / SEDEX) decide *se conta* unidades e é um alternador que o usuário liga, porque nada na folha distingue um lote de cartas simples de um de SEDEX. O formato "Meus Correios" nunca conta unidades, em modo nenhum — ele tem `Qtde. 1`, um envio por documento. Nada a jusante muda: o dict devolvido tem `codigo` e `condominio`, que é tudo que o carimbo e as planilhas consultam.

**Tech Stack:** Python 3.14, `re`, `winocr` via `extrair_texto_ocr_regiao`, CustomTkinter (aba 3), `unittest`.

## Global Constraints

- Comentários e nomes de variáveis em **português**.
- `logica.py` **não importa nada de interface** (nada de tkinter/customtkinter).
- Fixtures de teste são **strings inline** no arquivo de teste — nunca PDF no repositório, nunca nome real de morador.
- **Não alterar** `extrair_dados_protocolo_correio`, `conferir_contagem_protocolo`, `extrair_codigo_protocolo_correio`, `extrair_dados_protocolo_telnet`, `apurar_protocolo_telnet` nem `conferir_contagem_telnet`. Os caminhos do protocolo novo e do telnet ficam intocados.
- Cores da interface nunca hardcoded: sempre `self.tema_atual[chave]`. Todo widget CustomTkinter com `corner_radius=0`.
- O cobalto (`acento`) é reservado ao botão primário e ao que **resolve** uma pendência.
- Vocabulário da interface é para leigos: nunca "OCR" nem "DPI" em texto visível.
- Suíte hoje: **358 testes**. Rodar com `python -m unittest discover -s tests -p "test_*.py"`.
- Trabalhar em `C:\Users\Desktop\Pictures\projeto\.claude\worktrees\telnet` (branch `protocolo-telnet`).
- Validadores manuais levam prefixo `_` para o `unittest discover` não os coletar.

---

### Task 1: Discriminador reconhece o formato "Meus Correios"

**Files:**
- Modify: `logica.py:732` (junto de `RE_MARCADOR_TELNET`)
- Modify: `logica.py:804-826` (`classificar_formato_protocolo`)
- Test: `tests/test_protocolo_sedex.py` (criar)

**Interfaces:**
- Consumes: `RE_MARCADOR_PROTOCOLO` (`logica.py:723`), `RE_MARCADOR_TELNET` (`logica.py:732`)
- Produces: `RE_MARCADOR_MEUS_CORREIOS`; `classificar_formato_protocolo(texto)` passa a devolver `"meus_correios"` além de `"telnet"`, `"novo"`, `"ambiguo"` e `None`.

- [ ] **Step 1: Write the failing tests**

Criar `tests/test_protocolo_sedex.py`:

```python
# -*- coding: utf-8 -*-
"""Formato "Meus Correios" (sistema Agile) e modo SEDEX."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

#  Texto como o winocr devolve: página numa linha só, colunas fora de ordem.
#  Arranjo em TABELA (o mais comum nos 6 arquivos de referência).
MEUS_CORREIOS_TABELA = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10852-VILLA BRANCA Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Histórico ENVIO CONTRACHEQUES FUNCIONÁRIOS "
    "Contar. 1 Total: 1"
)

#  Arranjo VERTICAL (mesmo sistema, largura diferente).
MEUS_CORREIOS_VERTICAL = (
    "Correios [1 2704] ID 12704 Data 18/08/2026 Status Gerado "
    "Cód. Condomínio 1 0193-MONACO Qtde. 01 Classificação "
    "590 - EBCT - CORREIO REG / AR Histórico ENVI RESCISAO DE CONTRATO"
)

TELNET = (
    "*** IMODATA * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL DATA: 11/08/2026 "
    "EDF: MENESCAL LJ-01 AP-204 TOTAL ENVIADO PELO CORREIO..: 27 COD.: 1.1122.8)"
)

PROTOCOLO_NOVO = (
    "W700A VILLARS (10005) Protocolo de Recebimento de Documento "
    "Listando 3 unidades 702 - Fulano de Tal Correio"
)


class TestDiscriminadorMeusCorreios(unittest.TestCase):

    def test_reconhece_o_arranjo_em_tabela(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(MEUS_CORREIOS_TABELA),
            "meus_correios")

    def test_reconhece_o_arranjo_vertical(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(MEUS_CORREIOS_VERTICAL),
            "meus_correios")

    def test_nao_confunde_com_telnet(self):
        self.assertEqual(logica.classificar_formato_protocolo(TELNET), "telnet")

    def test_nao_confunde_com_o_protocolo_novo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(PROTOCOLO_NOVO), "novo")

    def test_cod_do_telnet_nao_dispara_o_marcador(self):
        #  O telnet tem "COD.: 1.1122.8)". O marcador do Meus Correios exige
        #  "COND" logo depois de "Cód", então não pode casar aqui — se casasse,
        #  todo telnet viraria ambíguo e o lote inteiro cairia em pendente.
        self.assertNotEqual(
            logica.classificar_formato_protocolo(TELNET), "ambiguo")

    def test_misturado_com_telnet_vira_ambiguo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(TELNET + " " + MEUS_CORREIOS_TABELA),
            "ambiguo")

    def test_misturado_com_o_novo_vira_ambiguo(self):
        self.assertEqual(
            logica.classificar_formato_protocolo(
                PROTOCOLO_NOVO + " " + MEUS_CORREIOS_TABELA),
            "ambiguo")

    def test_texto_sem_marcador_nenhum(self):
        self.assertIsNone(logica.classificar_formato_protocolo("boleto qualquer"))

    def test_formato_do_protocolo_funde_meus_correios_em_none(self):
        #  O wrapper antigo tem contrato de dois formatos; quem precisa do
        #  terceiro usa classificar_formato_protocolo diretamente.
        self.assertIsNone(logica.formato_do_protocolo(MEUS_CORREIOS_TABELA))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_protocolo_sedex -v`
Expected: FAIL — `classificar_formato_protocolo` devolve `None` para os textos do Meus Correios.

- [ ] **Step 3: Add the marker constant**

Em `logica.py`, logo depois de `RE_MARCADOR_TELNET` (linha 732):

```python
#  Marcador do formato "Meus Correios" (sistema Agile). Dois arranjos, mesmo
#  sistema: tabela ("Meus Correios") e vertical ("Correios [12704]"). O
#  terceiro marcador, "Cód. Condomínio", é o mais confiável dos três porque
#  aparece nos dois arranjos.
#
#  O "COND" depois de "Cód" é LOAD-BEARING: o telnet traz "COD.: 1.1122.8)"
#  na linha do total, e um marcador que casasse só "COD." tornaria todo
#  telnet ambíguo — o lote inteiro cairia em pendente.
RE_MARCADOR_MEUS_CORREIOS = re.compile(
    r"MEUS\s+CORREIOS"
    r"|C[OÓ]D[.\s]{0,3}COND[OÔ]M[IÍ]NIO"
    r"|CORREIOS\s*\[\s*\d[\d\s]{2,8}\]",
    re.IGNORECASE)
```

- [ ] **Step 4: Extend the discriminator**

Substituir o corpo de `classificar_formato_protocolo` (`logica.py:815-826`) por:

```python
    texto = texto or ""
    achados = []
    if RE_MARCADOR_TELNET.search(texto):
        achados.append("telnet")
    if RE_MARCADOR_PROTOCOLO.search(texto):
        achados.append("novo")
    if RE_MARCADOR_MEUS_CORREIOS.search(texto):
        achados.append("meus_correios")
    if len(achados) > 1:
        return "ambiguo"
    return achados[0] if achados else None
```

E acrescentar ao docstring da função, depois da primeira linha:

```
    Devolve "telnet", "novo", "meus_correios", "ambiguo" ou None (nenhum
    marcador). Mais de um marcador no mesmo texto é sempre "ambiguo",
    qualquer que seja a combinação.
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m unittest tests.test_protocolo_sedex -v`
Expected: PASS — 9 testes.

- [ ] **Step 6: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **367 testes, OK** (358 + 9). Nenhum teste do telnet ou do protocolo novo pode quebrar.

- [ ] **Step 7: Commit**

```bash
git add logica.py tests/test_protocolo_sedex.py
git commit -m "Discriminador reconhece o formato Meus Correios (Agile)"
```

---

### Task 2: Leitor do "Meus Correios"

**Files:**
- Modify: `logica.py` (constantes junto de `RE_MARCADOR_MEUS_CORREIOS`; função depois de `conferir_contagem_telnet`, que termina na linha ~985)
- Test: `tests/test_protocolo_sedex.py` (acrescentar classes)

**Interfaces:**
- Consumes: `_codigos_do_cadastro(cadastro)` (`logica.py:1276`) → dict código→cnpj_norm; `normalizar_texto_busca(texto)` (`logica.py:1259`). Os dois já existem — reimplementar qualquer um deles é defeito.
- Produces: `extrair_dados_meus_correios(texto, cadastro)` → `dict` com as chaves `formato` (sempre `"meus_correios"`), `codigo` (str ou `None`), `condominio` (str, nome do cadastro ou o do documento), `servico` (str ou `""`), `nome_confere` (bool)

- [ ] **Step 1: Write the failing tests**

Acrescentar a `tests/test_protocolo_sedex.py`, antes do `if __name__`:

```python
from cadastro_teste import CADASTRO_TESTE as CADASTRO


#  Um código que existe no cadastro de teste, com o nome que ele tem lá.
#  KLOSTERS é 10004 em tests/cadastro_teste.py.
MC_KLOSTERS = (
    "Correios Meus Correios Página 1 de 1 ID 12700 Data 13/08/2026 Status "
    "Gerado Cód. Condomínio 10004-KLOSTERS Qtde. 01 Classificação "
    "824 - EBCT - SEDEX Histórico ENVIO DE DOCUMENTOS Contar. 1 Total: 1"
)


class TestLeitorMeusCorreios(unittest.TestCase):

    def test_extrai_codigo_e_resolve_o_condominio(self):
        d = logica.extrair_dados_meus_correios(MC_KLOSTERS, CADASTRO)
        self.assertEqual(d["formato"], "meus_correios")
        self.assertEqual(d["codigo"], "10004")
        self.assertEqual(d["condominio"], "KLOSTERS")
        self.assertTrue(d["nome_confere"])

    def test_espaco_no_meio_dos_digitos(self):
        #  O OCR devolve "1 0004-KLOSTERS" com espaço no meio do número.
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "1 0004-KLOSTERS")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")

    def test_espaco_depois_do_traco(self):
        #  Medido: "10520- SENADOR LEITE OITICICA".
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "10004- KLOSTERS")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")
        self.assertTrue(d["nome_confere"])

    def test_codigo_fora_do_cadastro(self):
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "99999-INEXISTENTE")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertIsNone(d["codigo"])
        self.assertFalse(d["nome_confere"])

    def test_nome_divergente_nao_confere_mas_o_codigo_vale(self):
        #  Código bom, nome ilegível: o código manda (é o que o cadastro
        #  resolve), e nome_confere=False sinaliza a dúvida para a tela.
        texto = MC_KLOSTERS.replace("10004-KLOSTERS", "10004-XXXXXXX")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["codigo"], "10004")
        self.assertFalse(d["nome_confere"])

    def test_servico_sedex(self):
        d = logica.extrair_dados_meus_correios(MC_KLOSTERS, CADASTRO)
        self.assertEqual(d["servico"], "824 - EBCT - SEDEX")

    def test_servico_registrado_corta_no_historico(self):
        texto = MC_KLOSTERS.replace(
            "824 - EBCT - SEDEX Histórico",
            "590 - EBCT- CORREIO REG / AR Histórico")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "590 - EBCT- CORREIO REG / AR")

    def test_servico_corta_em_ID_quando_o_ocr_embaralha(self):
        #  Medido no telnet-003: depois da Classificação vem "ID 12699
        #  Contar: 1 Data ...", não "Histórico".
        texto = MC_KLOSTERS.replace(
            "824 - EBCT - SEDEX Histórico ENVIO DE DOCUMENTOS",
            "590 - EBCT- CORREIO REG / AR ID 12699 Contar: 1 Data 13/08/2026")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "590 - EBCT- CORREIO REG / AR")

    def test_sem_classificacao_o_servico_fica_vazio(self):
        #  Best-effort: o serviço é informação de apoio, nunca bloqueia nada.
        texto = MC_KLOSTERS.replace(
            "Classificação 824 - EBCT - SEDEX ", "")
        d = logica.extrair_dados_meus_correios(texto, CADASTRO)
        self.assertEqual(d["servico"], "")
        self.assertEqual(d["codigo"], "10004")

    def test_texto_vazio_nao_estoura(self):
        d = logica.extrair_dados_meus_correios("", CADASTRO)
        self.assertIsNone(d["codigo"])
        self.assertEqual(d["servico"], "")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m unittest tests.test_protocolo_sedex -v`
Expected: FAIL — `module 'logica' has no attribute 'extrair_dados_meus_correios'`.

- [ ] **Step 3: Add the extraction constants**

Em `logica.py`, logo depois de `RE_MARCADOR_MEUS_CORREIOS`:

```python
#  "Cód. Condomínio: 10852-VILLA BRANCA" — código E nome no mesmo campo, o
#  que dá conferência de graça: o código resolve pelo cadastro e o nome ao
#  lado confirma. Nenhum outro formato de protocolo oferece isso.
#
#  O `(?:\d\s*){5}` tolera espaço no meio dos dígitos: o OCR devolve
#  "1 0193-MONACO". O `\s*` depois do traço cobre "10520- SENADOR LEITE".
RE_COND_MEUS_CORREIOS = re.compile(
    r"COND[OÔ]M[IÍ]NIO\s*[:\-]?\s*((?:\d\s*){5})\s*[-–—]\s*([A-Z0-9 .'/]{2,40})",
    re.IGNORECASE)

#  A Classificação é o serviço postal (824 SEDEX, 590 CORREIO REG / AR). Vai
#  para a observação porque é o que diz à pessoa QUANTO digitar de valor.
RE_SERVICO_MEUS_CORREIOS = re.compile(
    r"CLASSIFICA[CÇ][AÃ]O\s*[:\-]?\s*(\d{3}\s*-[^\n]{2,60})", re.IGNORECASE)

#  Rótulos que podem vir logo DEPOIS do valor da Classificação, e onde ele
#  precisa ser cortado. Não é sempre "Histórico": o winocr embaralha as
#  colunas, e num dos seis arquivos de referência vem "ID 12699 Contar: 1
#  Data ...". Sem esse corte a observação carregaria meia página junto.
SEGUINTES_MEUS_CORREIOS = ("HISTORICO", "HISTÓRICO", "ID ", "CONTAR",
                           "DATA", "QTDE", "TOTAL", "FUNCION", "STATUS")
```

- [ ] **Step 4: Write the extraction function**

Em `logica.py`, depois de `conferir_contagem_telnet`:

```python
def _cortar_servico_meus_correios(bruto):
    """Corta o valor da Classificação no primeiro rótulo que vier depois.

    O winocr devolve a página numa linha só, então o campo seguinte fica
    colado no valor. Corta no rótulo mais à esquerda entre os conhecidos —
    não no primeiro da lista, que daria resultado diferente conforme a
    ordem em que os rótulos aparecerem no papel.
    """
    texto = (bruto or "").strip()
    alto = texto.upper()
    corte = len(texto)
    for rotulo in SEGUINTES_MEUS_CORREIOS:
        pos = alto.find(rotulo)
        if 0 < pos < corte:
            corte = pos
    return texto[:corte].strip(" -–—\t")


def extrair_dados_meus_correios(texto, cadastro):
    """
    Lê um protocolo do formato "Meus Correios" (sistema Agile).

    Devolve sempre um dict, nunca None:
      formato ....... "meus_correios"
      codigo ........ str de 5 dígitos que EXISTE no cadastro, ou None
      condominio .... nome do cadastro; na falta dele, o nome do documento
      servico ....... "824 - EBCT - SEDEX", ou "" quando não deu para ler
      nome_confere .. o nome impresso bate com o do cadastro

    Este formato NÃO tem unidades para contar — traz `Qtde. 1`, um envio por
    documento. Por isso não devolve contagem nem valor: o valor vem impresso
    num comprovante escaneado, que o OCR não lê com segurança (medido na
    investigação do telnet: `TOTAL: 27  103,95` não saiu em DPI nenhum), e é
    digitado à mão na grade do painel.

    `codigo` só é preenchido quando o código lido EXISTE no cadastro. Código
    que não existe é o mesmo que código não lido, para quem chama: não há o
    que carimbar nem como montar o lançamento, e a linha vira pendente.
    """
    texto = texto or ""
    achado = RE_COND_MEUS_CORREIOS.search(texto)
    nome_documento = ""
    codigo = None
    if achado:
        codigo = re.sub(r"\D", "", achado.group(1))[:5]
        nome_documento = achado.group(2).strip()

    registro = None
    if codigo:
        cnpj = _codigos_do_cadastro(cadastro or {}).get(codigo)
        registro = (cadastro or {}).get(cnpj) if cnpj else None
    if registro is None:
        codigo = None

    nome_cadastro = (registro or {}).get("nome", "")
    #  Compara só o começo: o nome do documento vem seguido do próximo campo
    #  ("KLOSTERS Qtde. 01"), porque o OCR não separa as colunas.
    confere = False
    if nome_cadastro:
        alvo = normalizar_texto_busca(nome_cadastro)
        lido = normalizar_texto_busca(nome_documento)
        confere = bool(alvo) and lido.startswith(alvo[:max(4, len(alvo) - 2)])

    servico = RE_SERVICO_MEUS_CORREIOS.search(texto)
    return {
        "formato": "meus_correios",
        "codigo": codigo,
        "condominio": nome_cadastro or nome_documento,
        "servico": _cortar_servico_meus_correios(servico.group(1)) if servico else "",
        "nome_confere": confere,
    }
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m unittest tests.test_protocolo_sedex -v`
Expected: PASS — 19 testes na classe nova mais as 9 da Task 1.

- [ ] **Step 6: Run the whole suite**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **377 testes, OK** (367 + 10).

- [ ] **Step 7: Commit**

```bash
git add logica.py tests/test_protocolo_sedex.py
git commit -m "Le o codigo e o servico do protocolo Meus Correios"
```

---

### Task 3: Alternador "Modo SEDEX" na aba 3

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py` (`__init__`, junto de `self.separar_em_lotes` na linha ~260)
- Modify: `identificacao_por_cnpj_6_0.py:850-908` (`_montar_controle_lotes`) — acrescentar um método irmão
- Modify: `identificacao_por_cnpj_6_0.py:2641` (montagem da aba 3)
- Modify: `identificacao_por_cnpj_6_0.py:2839` (`_acao_processar_protocolos`, pular a tarifa)

**Interfaces:**
- Consumes: `self.tema_atual`, `familia_fonte()`, o padrão `registrar(widget, {chave: token})` usado por `_montar_controle_lotes`
- Produces: `self.modo_sedex` (`tk.BooleanVar`); `_montar_controle_sedex(parent, registrar)` → frame

- [ ] **Step 1: Declare the variable**

Em `__init__`, logo abaixo de `self.separar_em_lotes` (linha ~261):

```python
        #  Modo SEDEX: desliga a contagem de unidades no lote inteiro. NÃO
        #  vem do config e NÃO é salvo lá, ao contrário de `separar_em_lotes`
        #  — aquele é preferência do sistema de destino, este é propriedade
        #  do lote que está na mesa agora. Nasce desligado toda sessão.
        self.modo_sedex = tk.BooleanVar(value=False)
```

- [ ] **Step 2: Build the control**

Acrescentar depois de `_salvar_preferencia_lotes` (linha ~918):

```python
    def _montar_controle_sedex(self, parent, registrar):
        """
        Alternador "Modo SEDEX", para ficar logo abaixo do controle de lotes
        na aba 3.

        Só existe na aba 3: é a única que conta unidades e calcula valor.

        Sem `command` de persistência de propósito — este alternador não é
        gravado no config.json. Deixar ligado de uma sessão para a outra faria
        um lote de cartas simples sair inteiro sem contagem.
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        linha = registrar(
            ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )

        chk = registrar(
            ctk.CTkCheckBox(
                linha,
                text="Modo SEDEX — não contar unidades, valor digitado à mão",
                variable=self.modo_sedex, corner_radius=0,
                fg_color=tema["acento"], hover_color=tema["acento"],
                border_color=tema["borda_forte"], text_color=tema["texto"],
                font=(fonte, 13),
            ),
            {"fg_color": "acento", "hover_color": "acento",
             "border_color": "borda_forte", "text_color": "texto"},
        )
        chk.pack(side="left")
        return linha
```

- [ ] **Step 3: Mount it in tab 3**

O controle de lotes da aba 3 está em `row=7` (linha 2641-2642):

```python
        self._montar_controle_lotes(corpo, registrar).grid(
            row=7, column=0, sticky="w", pady=(0, 8))
```

Acrescentar **logo depois** dele:

```python
        self._montar_controle_sedex(corpo, registrar).grid(
            row=8, column=0, sticky="w", pady=(0, 8))
```

E somar 1 ao `row` dos três widgets que vinham depois, **nesta ordem** (de baixo para cima, para os números não colidirem no meio da edição):

| Widget | De | Para |
|---|---|---|
| `self.label_status_protocolos.grid(...)` | `row=10` | `row=11` |
| `self.barra_protocolos.grid(...)` | `row=9` | `row=10` |
| `explicacao.grid(...)` | `row=8` | `row=9` |

Nenhum outro widget da aba 3 usa `row` maior que 10 — conferido.

- [ ] **Step 4: Skip the tariff question in SEDEX mode**

Em `_acao_processar_protocolos`, substituir (linha ~2839):

```python
        tarifa = self._pedir_tarifa()
        if tarifa is None:
            return
```

por:

```python
        #  Em modo SEDEX não há o que multiplicar: a tarifa não é perguntada
        #  e segue None por todo o fluxo. Isso é seguro de ponta a ponta —
        #  `linha_planilha_protocolo` já devolve Unidades/Tarifa/Valor vazios
        #  quando `unidades` é None, `valor_do_carimbo` já devolve None e
        #  carimba só a identificação lateral, e `validar_edicao_protocolo` já
        #  recusa editar a coluna Unidades explicando que o lote não tem
        #  tarifa. Nenhuma guarda nova é necessária.
        tarifa = None
        if not self.modo_sedex.get():
            tarifa = self._pedir_tarifa()
            if tarifa is None:
                return
```

- [ ] **Step 5: Pass the mode to the worker**

Na chamada da thread (linha ~2864), acrescentar o argumento:

```python
        thread = threading.Thread(
            target=self._processar_protocolos_em_thread,
            args=(pasta, pasta_saida, destino, tarifa, vencimento, chave,
                  self.modo_sedex.get()),
            daemon=True)
```

E na assinatura (linha ~2970):

```python
    def _processar_protocolos_em_thread(self, pasta, pasta_saida, destino, tarifa,
                                        vencimento, chave, modo_sedex=False):
```

- [ ] **Step 6: Verify it imports and the suite still passes**

Run: `python -c "import identificacao_por_cnpj_6_0; print('ok')"`
Expected: `ok`

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **377 testes, OK** (esta tarefa não acrescenta teste — é interface).

- [ ] **Step 7: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Alternador Modo SEDEX na aba 3, sem persistir entre sessoes"
```

---

### Task 4: Roteamento do "Meus Correios" e do modo SEDEX no laço

**Files:**
- Modify: `identificacao_por_cnpj_6_0.py:3010-3040` (o `if formato == ...` dentro de `_processar_protocolos_em_thread`)
- Modify: `identificacao_por_cnpj_6_0.py` (import de `logica`, no topo)

**Interfaces:**
- Consumes: `classificar_formato_protocolo` (Task 1), `extrair_dados_meus_correios(texto, cadastro)` (Task 2), `self._classificar_arquivo_protocolo(caminho, texto)` → `(formato, leitura_extra)`
- Produces: nada novo — o laço continua saindo com `dados`, `unidades`, `observacao`

- [ ] **Step 1: Add the import**

Na lista de `from logica import (...)` no topo do arquivo, acrescentar `extrair_dados_meus_correios,` junto das outras funções de protocolo.

- [ ] **Step 1b: Fix the classifier's docstring**

`_classificar_arquivo_protocolo` (linha ~2934) **não precisa de mudança de código** — ele repassa o que `classificar_formato_protocolo` devolver, então `"meus_correios"` já flui sozinho. Mas o docstring dele lista só três formatos e passaria a mentir. Trocar a primeira frase por:

```
        Qual dos formatos de protocolo dos Correios este arquivo é:
        "telnet", "novo", "meus_correios", "ambiguo" (mais de um marcador
        no mesmo texto — ver `classificar_formato_protocolo` em logica.py)
        ou None (nenhum marcador, nem na segunda chance).
```

E, mais abaixo no mesmo docstring, trocar `"telnet", "novo" e "ambiguo" já são decisões com confiança` por `"telnet", "novo", "meus_correios" e "ambiguo" já são decisões com confiança`.

- [ ] **Step 2: Route the new format**

No `if/elif/else` do laço (linha ~3012), acrescentar um ramo **antes** do `else:` final:

```python
                elif formato == "meus_correios":
                    #  Formato do sistema Agile: um envio por documento, sem
                    #  lista de unidades para contar (traz `Qtde. 1`). Sai
                    #  daqui com `unidades` em None de propósito — o valor é
                    #  digitado na grade, porque ele vem num comprovante
                    #  escaneado que o OCR não lê com segurança.
                    dados = extrair_dados_meus_correios(texto, self.cadastro)
                    #  A Classificação (824 SEDEX / 590 CORREIO REG / AR) é o
                    #  que diz à pessoa QUANTO digitar. Vai primeiro na
                    #  observação: numa linha pendente, o texto que explica o
                    #  que fazer tem de sobreviver ao corte da coluna.
                    if dados["servico"]:
                        observacao = dados["servico"]
                    if not dados["codigo"]:
                        dados = None
```

- [ ] **Step 3: Make SEDEX mode skip counting**

No ramo `if formato == "telnet":`, substituir:

```python
                    aceito, observacao = conferir_contagem_telnet(dados)
                    if aceito:
                        unidades = dados["total_impresso"]
```

por:

```python
                    if modo_sedex:
                        #  O modo desliga só a contagem: o código e o nome
                        #  continuam sendo lidos e carimbados. `unidades`
                        #  fica None e o valor é digitado na grade.
                        pass
                    else:
                        aceito, observacao = conferir_contagem_telnet(dados)
                        if aceito:
                            unidades = dados["total_impresso"]
```

E no ramo `else:` (protocolo novo), envolver a checagem de contagem existente — o bloco que começa em `ja_tentou_de_novo = False` e termina em `observacao = motivo` — num `if not modo_sedex:`, **re-indentando um nível e sem alterar uma vírgula do conteúdo**. O `dados = extrair_dados_protocolo_correio(texto)` fica FORA desse `if`: em modo SEDEX o documento ainda precisa ser identificado, só não contado.

- [ ] **Step 4: Verify the re-indentation changed nothing else**

Run: `git diff -b -- identificacao_por_cnpj_6_0.py`
Expected: nenhuma linha do bloco antigo aparece como removida ou modificada — só as adições novas. Se aparecer qualquer outra diferença, desfazer e refazer.

- [ ] **Step 5: Verify it imports and the suite still passes**

Run: `python -c "import identificacao_por_cnpj_6_0; print('ok')"`
Expected: `ok`

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **377 testes, OK**.

- [ ] **Step 6: Commit**

```bash
git add identificacao_por_cnpj_6_0.py
git commit -m "Roteia o Meus Correios e faz o modo SEDEX pular a contagem"
```

---

### Task 5: Validação contra PDFs reais, CLAUDE.md e versão

**Files:**
- Create: `tests/_validar_meus_correios.py`
- Modify: `CLAUDE.md` (seção nova antes de `## Anexo automático no Paybox (aba 3, v6.16.0)`)
- Modify: `identificacao_por_cnpj_6_0.py:210` (`self.title`)

**Interfaces:**
- Consumes: `classificar_formato_protocolo`, `extrair_dados_meus_correios`, `carregar_cadastro`

- [ ] **Step 1: Write the manual validator**

Criar `tests/_validar_meus_correios.py`:

```python
# -*- coding: utf-8 -*-
"""
Valida o leitor do "Meus Correios" contra os PDFs reais.

Fora da suíte (prefixo `_`): depende de arquivos que não estão no
repositório. Reexecutável.

Rode:  python tests/_validar_meus_correios.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

PASTA = r"C:\Users\Desktop\Desktop\Correios\telnet"
#  Os 6 "Meus Correios" que convivem com os 7 telnet nessa pasta.
ESPERADO = {
    "procolo telnet-002.pdf": "10852",
    "procolo telnet-003.pdf": "10520",
    "procolo telnet-004.pdf": "11124",
    "procolo telnet-005.pdf": "10193",
    "procolo telnet-006.pdf": "10255",
    "procolo telnet.pdf": "11124",
}

cadastro = logica.carregar_cadastro("cadastro_condominios.xlsx")

print(f"{'arquivo':26} {'formato':14} {'cod':>6} {'esp':>6} {'nome':5} servico")
print("-" * 96)
certos = trocados = 0
for nome, esperado in sorted(ESPERADO.items()):
    caminho = os.path.join(PASTA, nome)
    texto = logica.extrair_texto_ocr_regiao(caminho, (0, 0, 1, 0.45), dpi=400)
    formato = logica.classificar_formato_protocolo(texto)
    if formato != "meus_correios":
        trocados += 1
        print(f"{nome[-24:]:26} {str(formato):14} FORMATO TROCADO")
        continue
    d = logica.extrair_dados_meus_correios(texto, cadastro)
    ok = d["codigo"] == esperado
    certos += ok
    print(f"{nome[-24:]:26} {formato:14} {str(d['codigo']):>6} {esperado:>6} "
          f"{('sim' if d['nome_confere'] else 'nao'):5} {d['servico']}")

print(f"\ncodigos corretos: {certos}/{len(ESPERADO)}")
print(f"FORMATOS TROCADOS: {trocados}   <- tem que ser 0")
```

- [ ] **Step 2: Run the validator**

Run: `python tests/_validar_meus_correios.py`
Expected: **códigos corretos 6/6** e **FORMATOS TROCADOS: 0**.

Se divergir, **parar e relatar** — não ajustar o validador para casar com o número esperado.

- [ ] **Step 3: Re-run the telnet validators (regression)**

Run: `python tests/_validar_discriminador.py`
Expected: `NENHUM FORMATO TROCADO` — o marcador novo não pode ter tornado nenhum telnet ambíguo.

Run: `python tests/_validar_telnet.py`
Expected: códigos 7/7, totais 6/7, `VALORES ERRADOS PREENCHIDOS: 0`.

- [ ] **Step 4: Document it in CLAUDE.md**

Inserir **antes** de `## Anexo automático no Paybox (aba 3, v6.16.0)`:

```markdown
## Modo SEDEX e o protocolo "Meus Correios" (aba 3, v6.20.0)

Terceiro formato de protocolo, gerado pelo sistema **Agile**, e um modo de
lote que desliga a contagem de unidades. São coisas independentes:

```
FORMATO (novo / telnet / meus correios) → COMO ler o condomínio → o documento diz
MODO    (normal / SEDEX)                → SE conta unidades     → o documento NÃO diz
```

O mesmo envio SEDEX chega nos três formatos, e nada na folha distingue um
protocolo de cartas simples de um de SEDEX — por isso o modo é um alternador
que o usuário liga, e não algo deduzido do papel.

| Formato | Modo normal | Modo SEDEX |
|---|---|---|
| novo (Superlógica) | conta unidades | não conta, valor à mão |
| telnet | conta unidades | não conta, valor à mão |
| **meus correios** | **não conta (intrínseco)** | não conta |

**"Meus Correios" nunca conta unidades, em modo nenhum** — esse formato traz
`Qtde. 1`, um envio por documento. Isso é propriedade do formato, não do
modo, e não é esquecimento.

**O valor é digitado, e isso foi medido, não presumido.** Ele vem impresso
num comprovante escaneado, e a investigação do telnet mostrou que o valor
desses comprovantes (`TOTAL: 27  103,95`) **não sai em DPI nenhum** (200, 300,
400). Vale aqui o mesmo princípio já registrado na aba 2: valor e data não têm
dígito verificador, então uma leitura errada entra na planilha sem ninguém
perceber. Digitando na **grade do painel**, `_recarimbar_linha` regrava o PDF
com valor e bloco do Paybox — o anexo automático volta a funcionar, o que não
aconteceria preenchendo no Excel depois.

**A leitura é a mais fácil dos três formatos:** `Cód. Condomínio:
10852-VILLA BRANCA` traz código e nome no mesmo campo, então o código resolve
pelo cadastro e o nome ao lado confirma. Medido: **6/6 códigos** nos arquivos
de referência.

**Dois detalhes de regex que são load-bearing:**

- O marcador exige `COND` logo depois de `Cód` (`C[OÓ]D[.\s]{0,3}COND...`).
  O telnet traz `COD.: 1.1122.8)` na linha do total; um marcador que casasse
  só `COD.` tornaria **todo telnet ambíguo** e o lote inteiro cairia em
  pendente.
- O valor da Classificação é cortado no rótulo seguinte, e **não é sempre
  "Histórico"**: o winocr embaralha as colunas, e num dos seis arquivos vem
  `ID 12699 Contar: 1 Data ...`. Sem o corte, a observação carregaria meia
  página junto.

**A Classificação vai para a observação** (`824 - EBCT - SEDEX`), como está no
documento, sem traduzir — é o texto que a pessoa cruza com a tabela de preços
dos Correios para saber quanto digitar. É best-effort: não sendo legível, a
observação sai sem ela e nada é bloqueado.

**Nos 6 arquivos de referência, 5 são `590 - EBCT - CORREIO REG / AR` e só 1
é `824 - EBCT - SEDEX`**, apesar de o lote ser chamado de "os sedex". São
serviços diferentes, com preços possivelmente diferentes — é por isso que o
valor é por linha e não um só para o lote.

**O alternador não persiste entre sessões**, ao contrário do de lotes: aquele
é preferência do sistema de destino, este é propriedade do lote que está na
mesa. Deixá-lo ligado faria um lote de cartas simples sair inteiro sem
contagem.

**`tarifa=None` é seguro de ponta a ponta em modo SEDEX**, e isso não exigiu
guarda nova: `linha_planilha_protocolo` já devolve Unidades/Tarifa/Valor
vazios quando `unidades` é None, `valor_do_carimbo` já devolve None e carimba
só a identificação lateral, e `validar_edicao_protocolo` já recusa editar a
coluna Unidades com a frase "Este lote não tem tarifa...".

**Não validado:** SEDEX nos formatos telnet e Superlógica. O usuário informou
que existem, mas não tinha exemplos na época. O desenho cobre os três; a
medição cobriu um. É o mesmo ponto cego que produziu a v6.18.1 — lá, validar
contra um lote de uma competência só escondeu uma variação de layout que
quebrava a leitura inteira.

Spec: `docs/superpowers/specs/2026-09-07-modo-sedex-design.md`.
Validador manual: `tests/_validar_meus_correios.py`.
```

- [ ] **Step 5: Bump the version**

Em `identificacao_por_cnpj_6_0.py:210`:

```python
        self.title("Codificador v6.20.0")
```

- [ ] **Step 6: Run the whole suite one last time**

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: **377 testes, OK**.

- [ ] **Step 7: Commit**

```bash
git add CLAUDE.md identificacao_por_cnpj_6_0.py tests/_validar_meus_correios.py
git commit -m "Documenta o modo SEDEX e o Meus Correios, marca a versao 6.20.0"
```

---

## Notas para quem executa

- **O passo mais delicado do plano é a re-indentação na Task 4, Step 3.** É o mesmo tipo de mudança que a branch do telnet fez em `_processar_protocolos_em_thread`, e lá funcionou porque foi conferida com `git diff -b`. Faça o mesmo: se `-b` mostrar qualquer linha do bloco antigo como alterada, o conteúdo mudou junto com a indentação.
- **Um teste manual pela interface fica para o controlador**, não para o implementador: rodar um lote misto com o alternador ligado e desligado, e confirmar que em modo normal nada mudou.
- O caminho do protocolo novo e o do telnet têm que sair desta branch **idênticos** em modo normal. Os dois validadores manuais existem para provar isso.
