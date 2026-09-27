# Associação automática das notas da F&F no Paybox — Plano de Implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Um favorito do navegador (bookmarklet) que percorre a fila do Paybox do Superlógica e associa cada NFS-e da F&F à sua despesa, com simulação, relatório CSV e parada no inesperado.

**Architecture:** Todo o código fica numa única função, `criarPayboxFF(ambiente)`, em `ferramentas/paybox_ff/associar_ff.js`. É assim que ela cabe inteira num favorito: `instalar.html` monta o link a partir de `criarPayboxFF.toString()`. O `ambiente` injeta rede, espera, confirmação e download. Nos testes, a rede é trocada por respostas gravadas, então todo o fluxo roda sem Superlógica. Os testes rodam no Edge sem janela, chamado por um script Python, porque não há Node na máquina.

**Tech Stack:** JavaScript de navegador (ES2017: `async/await`, sem módulos, sem bibliotecas), HTML, Python 3 (só o executor de testes), Microsoft Edge headless.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-09-27-associacao-paybox-ff-design.md`.
- **Fora do Codificador:** nada em `logica.py`, `identificacao_por_cnpj_6_0.py`, `codificador.spec`, `gerar_exe.bat` nem `requirements*.txt`. Tudo em `ferramentas/paybox_ff/`. A suíte Python do Codificador (`python -m unittest discover -s tests -p "test_*.py"`, **463 testes**) não pode mudar.
- **`associar_ff.js` contém só a função `criarPayboxFF`** e nada fora dela: nenhuma variável global e nenhum código de topo. Ela precisa funcionar depois de convertida em texto por `toString()`, então **não pode referenciar nada de fora de si mesma** além do que vem em `ambiente` e dos globais do navegador (`JSON`, `Promise`, `URLSearchParams`, `Blob`, `Date`, `Math`, `Array`, `Object`, `String`, `Number`, `Error`).
- **Nunca ler, guardar, registrar nem enviar** senha, cookie, token ou o campo `session` das respostas. Mensagens de erro nunca incluem o corpo cru de uma resposta.
- **Nenhuma chamada a outro domínio:** toda requisição vai para `ambiente.origem + "/condor/atual/..."`.
- CNPJ da F&F: `"13736666000154"`. Pausa entre associações: `1000` ms. Tolerância de valor: `0.005`.
- Comentários, nomes e textos visíveis em **português**. Textos visíveis sem jargão ("fila", "nota", "despesa"; nunca "envelope", "endpoint", "JSON").
- Visual do painel no estilo Swiss do Codificador: cantos retos (`border-radius: 0`) e paleta do `TEMA_CLARO` (fundo `#FAFAF9`, superfície `#F5F5F4`, borda `#E7E5E4`, texto `#1C1917`, texto secundário `#57534E`). O cobalto `#003B8E` (texto `#FFFFFF`) só aparece no botão **Iniciar**.
- Nenhum HAR, `index.json` ou dado real de sessão entra no repositório. As fixtures são as dos testes abaixo: sintéticas, com o CNPJ de teste `01195716000154` (KLOSTERS, das fixtures do Codificador).
- Rodar os testes: `python ferramentas/paybox_ff/rodar_testes.py` (a saída termina em `RESUMO: N ok, 0 falhas` e o código de saída é 0).

---

### Task 1: Esqueleto, funções puras e a bancada de testes

**Files:**
- Create: `ferramentas/paybox_ff/associar_ff.js`
- Create: `ferramentas/paybox_ff/testes.html`
- Create: `ferramentas/paybox_ff/testes.js`
- Create: `ferramentas/paybox_ff/rodar_testes.py`

**Interfaces:**
- Produces:
  - `criarPayboxFF(ambiente) -> { logica: {...}, processar, abrir, execucaoAtual }`. Nesta task só existe `logica`.
  - `logica.CNPJ_FF` = `"13736666000154"`.
  - `logica.soDigitos(texto) -> string`.
  - `logica.normalizarNumeroDocumento(texto) -> string`: só dígitos, sem zeros à esquerda; `"0"` continua `"0"`.
  - `logica.dataBrParaUs("27/09/2026") -> "09/27/2026"`, ou `null` se o texto for inválido.
  - `logica.dataUsSemHora("08/24/2026 00:00:00") -> "08/24/2026"`.
  - `logica.paraNumero("84.65") -> 84.65`, ou `null`.
  - `logica.valorComPonto(84.65) -> "84.65"` e `logica.valorComVirgula(84.65) -> "84,65"`.
  - `logica.valoresIguais(a, b) -> boolean`: tolerância de 0.005, e `false` se algum for `null`.
  - Na bancada (`testes.js`): `teste(nome, fn)`, `igual(obtido, esperado, msg)`, `verdade(valor, msg)`, `lanca(fn, trecho)`, `ambienteFalso(transporte)` e `transporteFalso(rotas)`. As tasks seguintes os usam.

- [ ] **Step 1: Criar o executor dos testes**

Crie `ferramentas/paybox_ff/rodar_testes.py`:

```python
# -*- coding: utf-8 -*-
"""Roda testes.html no Edge sem janela e diz se passou.

Uso: python ferramentas/paybox_ff/rodar_testes.py

Não há Node nesta máquina; o Edge vem com o Windows e roda o mesmo
JavaScript que o favorito vai rodar dentro do Superlógica.
"""
import html
import os
import re
import subprocess
import sys
import tempfile
import urllib.request

AQUI = os.path.dirname(os.path.abspath(__file__))
NAVEGADORES = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
)


def navegador():
    for caminho in NAVEGADORES:
        if os.path.exists(caminho):
            return caminho
    sys.exit("Nenhum Edge ou Chrome encontrado para rodar os testes.")


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    pagina = "file:" + urllib.request.pathname2url(os.path.join(AQUI, "testes.html"))
    #  Perfil temporário: sem ele, um Edge já aberto pelo usuário trava o
    #  perfil padrão e o modo sem janela sai sem rodar nada.
    with tempfile.TemporaryDirectory() as perfil:
        saida = subprocess.run(
            [navegador(), "--headless=new", "--disable-gpu",
             f"--user-data-dir={perfil}", "--virtual-time-budget=20000",
             "--dump-dom", pagina],
            capture_output=True, text=True, encoding="utf-8", timeout=180,
        ).stdout
    achado = re.search(r'<pre id="resultado">(.*?)</pre>', saida, re.S)
    if not achado:
        print(saida[-2000:])
        sys.exit("Não achei o resultado dos testes na página.")
    texto = html.unescape(achado.group(1))
    print(texto)
    #  "0 ok" também é falha: significa que o arquivo de casos quebrou antes
    #  de registrar os testes (ex.: associar_ff.js com erro de sintaxe).
    resumo = re.search(r"RESUMO: (\d+) ok, (\d+) falhas", texto)
    passou = resumo and int(resumo.group(1)) > 0 and resumo.group(2) == "0"
    sys.exit(0 if passou else 1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Criar a página dos testes**

Crie `ferramentas/paybox_ff/testes.html`:

```html
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Testes — associar notas da F&amp;F no Paybox</title>
</head>
<body>
<pre id="resultado">pendente</pre>
<script src="associar_ff.js"></script>
<script src="testes.js"></script>
<script>
  //  Roda os casos em sequência (alguns são async) e escreve o resultado no
  //  <pre>, que é de onde rodar_testes.py o lê.
  (async function () {
    var linhas = [], ok = 0, falhas = 0;
    for (var i = 0; i < CASOS.length; i++) {
      var nome = CASOS[i][0], fn = CASOS[i][1];
      try { await fn(); ok++; linhas.push("ok      " + nome); }
      catch (e) { falhas++; linhas.push("FALHOU  " + nome + " — " + (e && e.message)); }
    }
    linhas.push("RESUMO: " + ok + " ok, " + falhas + " falhas");
    document.getElementById("resultado").textContent = linhas.join("\n");
  })();
</script>
</body>
</html>
```

- [ ] **Step 3: Escrever a bancada e os testes das funções puras**

Crie `ferramentas/paybox_ff/testes.js`:

```js
// Casos de teste do associar_ff.js. Rodados por testes.html.
// Nada aqui é dado real: CNPJ de teste do Codificador (KLOSTERS) e números
// inventados no formato das respostas do Superlógica.

var CASOS = [];
function teste(nome, fn) { CASOS.push([nome, fn]); }

function igual(obtido, esperado, msg) {
  var a = JSON.stringify(obtido), b = JSON.stringify(esperado);
  if (a !== b) throw new Error((msg ? msg + ": " : "") + "esperado " + b + ", veio " + a);
}
function verdade(valor, msg) { if (!valor) throw new Error(msg || "esperado verdadeiro"); }
async function lanca(fn, trecho) {
  try { await fn(); }
  catch (e) {
    if (trecho && String(e && e.message).indexOf(trecho) < 0) {
      throw new Error("erro com mensagem inesperada: " + (e && e.message));
    }
    return e;
  }
  throw new Error("esperava um erro e não veio nenhum");
}

// Transporte falso: `rotas(caminho, pares)` devolve o objeto de resposta
// (vira texto JSON, como viria da rede) ou uma string crua (ex.: HTML de
// login). Guarda todas as chamadas em `.chamadas` para os testes conferirem.
function transporteFalso(rotas) {
  var chamadas = [];
  var t = function (caminho, pares) {
    chamadas.push({ caminho: caminho, pares: pares });
    try {
      var r = rotas(caminho, pares);
      return Promise.resolve(typeof r === "string" ? r : JSON.stringify(r));
    } catch (e) { return Promise.reject(e); }
  };
  t.chamadas = chamadas;
  return t;
}

function ambienteFalso(transporte, extra) {
  var amb = {
    janela: window,
    origem: "https://admin1.superlogica.net",
    hostname: "admin1.superlogica.net",
    transporte: transporte || transporteFalso(function () { throw new Error("sem rota"); }),
    esperar: function () { return Promise.resolve(); },
    confirmar: function () { return true; },
    avisar: function () {},
    baixar: function () {}
  };
  for (var k in (extra || {})) amb[k] = extra[k];
  return amb;
}

var L = criarPayboxFF(ambienteFalso()).logica;

// ---- funções puras ----

teste("CNPJ da F&F é o fixado no spec", function () {
  igual(L.CNPJ_FF, "13736666000154");
});

teste("soDigitos tira pontuação e tolera nulo", function () {
  igual(L.soDigitos("13.736.666/0001-54"), "13736666000154");
  igual(L.soDigitos(null), "");
  igual(L.soDigitos(12048), "12048");
});

teste("número do documento compara sem zeros à esquerda", function () {
  igual(L.normalizarNumeroDocumento("000012048"), "12048");
  igual(L.normalizarNumeroDocumento(" 12.048 "), "12048");
  igual(L.normalizarNumeroDocumento("0"), "0");
  igual(L.normalizarNumeroDocumento(""), "");
});

teste("data DD/MM/AAAA vira MM/DD/AAAA, como o Superlógica espera", function () {
  igual(L.dataBrParaUs("27/09/2026"), "09/27/2026");
  igual(L.dataBrParaUs(" 01/10/2026 "), "10/01/2026");
});

teste("data inválida é recusada, não virada em outra", function () {
  igual(L.dataBrParaUs("31/02/2026"), null);
  igual(L.dataBrParaUs("2026-09-27"), null);
  igual(L.dataBrParaUs("27/9/2026"), null);
  igual(L.dataBrParaUs(""), null);
  igual(L.dataBrParaUs(null), null);
});

teste("data da despesa perde a hora", function () {
  igual(L.dataUsSemHora("08/24/2026 00:00:00"), "08/24/2026");
  igual(L.dataUsSemHora("08/24/2026"), "08/24/2026");
  igual(L.dataUsSemHora(""), "");
});

teste("valores do Superlógica vêm com ponto", function () {
  igual(L.paraNumero("84.65"), 84.65);
  igual(L.paraNumero(" 78.61 "), 78.61);
  igual(L.paraNumero(""), null);
  igual(L.paraNumero(null), null);
  igual(L.paraNumero("abc"), null);
});

teste("valor com ponto e com vírgula, sempre 2 casas", function () {
  igual(L.valorComPonto(84.65), "84.65");
  igual(L.valorComPonto(80), "80.00");
  igual(L.valorComVirgula(84.65), "84,65");
  igual(L.valorComVirgula(1234.5), "1234,50");
});

teste("valores iguais dentro da tolerância; nulo nunca é igual", function () {
  verdade(L.valoresIguais(84.65, 84.65));
  verdade(L.valoresIguais(84.65, 84.654));
  verdade(!L.valoresIguais(84.65, 84.66));
  verdade(!L.valoresIguais(null, 84.65));
  verdade(!L.valoresIguais(84.65, null));
});
```

- [ ] **Step 4: Rodar e ver falhar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: código de saída 1, com `RESUMO: 0 ok, 0 falhas`. `associar_ff.js` ainda não existe, então `testes.js` para em `var L = criarPayboxFF(...)` sem registrar nenhum caso. O executor trata "0 ok" como falha justamente para isto não passar em silêncio.

- [ ] **Step 5: Escrever o esqueleto e as funções puras**

Crie `ferramentas/paybox_ff/associar_ff.js`:

```js
// Associa as NFS-e da F&F às despesas no Paybox do Superlógica.
//
// Tudo vive dentro de criarPayboxFF de propósito: instalar.html transforma
// esta função em texto (toString) para montar o favorito, então ela não pode
// depender de nada que esteja fora dela. `ambiente` traz o que muda entre o
// navegador de verdade e os testes (rede, espera, confirmação, download).
//
// Nunca lê, guarda nem envia senha, cookie, token ou o campo "session" das
// respostas: as chamadas usam a sessão que a própria página já tem.
// Spec: docs/superpowers/specs/2026-09-27-associacao-paybox-ff-design.md
function criarPayboxFF(ambiente) {
  "use strict";

  var CNPJ_FF = "13736666000154";
  var PAUSA_ENTRE_ASSOCIACOES_MS = 1000;
  var TOLERANCIA_VALOR = 0.005;

  // ---- funções puras ----

  function soDigitos(texto) {
    return String(texto == null ? "" : texto).replace(/\D/g, "");
  }

  //  O Superlógica guarda o número do documento como texto; "012048" e
  //  "12048" são a mesma nota.
  function normalizarNumeroDocumento(texto) {
    return soDigitos(texto).replace(/^0+(?=\d)/, "");
  }

  //  Quem digita é brasileiro; a busca de despesas do Superlógica espera
  //  MM/DD/AAAA. Data impossível (31/02) é recusada em vez de "rolar" para
  //  o mês seguinte, que é o que o Date do JavaScript faria.
  function dataBrParaUs(texto) {
    var m = /^\s*(\d{2})\/(\d{2})\/(\d{4})\s*$/.exec(texto == null ? "" : String(texto));
    if (!m) return null;
    var dia = +m[1], mes = +m[2], ano = +m[3];
    var d = new Date(ano, mes - 1, dia);
    if (d.getFullYear() !== ano || d.getMonth() !== mes - 1 || d.getDate() !== dia) return null;
    return m[2] + "/" + m[1] + "/" + m[3];
  }

  function dataUsSemHora(texto) {
    var m = /^(\d{2}\/\d{2}\/\d{4})/.exec(texto == null ? "" : String(texto));
    return m ? m[1] : "";
  }

  //  Valores do Superlógica vêm como texto com ponto decimal ("84.65").
  function paraNumero(texto) {
    if (texto == null) return null;
    var limpo = String(texto).trim();
    if (!limpo) return null;
    var n = Number(limpo);
    return isFinite(n) ? n : null;
  }

  function valorComPonto(n) { return n.toFixed(2); }
  function valorComVirgula(n) { return n.toFixed(2).replace(".", ","); }

  function valoresIguais(a, b) {
    return a != null && b != null && Math.abs(a - b) < TOLERANCIA_VALOR;
  }

  return {
    logica: {
      CNPJ_FF: CNPJ_FF,
      soDigitos: soDigitos,
      normalizarNumeroDocumento: normalizarNumeroDocumento,
      dataBrParaUs: dataBrParaUs,
      dataUsSemHora: dataUsSemHora,
      paraNumero: paraNumero,
      valorComPonto: valorComPonto,
      valorComVirgula: valorComVirgula,
      valoresIguais: valoresIguais
    }
  };
}
```

`PAUSA_ENTRE_ASSOCIACOES_MS` ainda não é usada: ela é da Task 4 e fica declarada aqui para as constantes do spec ficarem juntas no topo.

- [ ] **Step 6: Rodar e ver passar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: 9 linhas `ok`, `RESUMO: 9 ok, 0 falhas`, código de saída 0.

- [ ] **Step 7: Commit**

```bash
git add ferramentas/paybox_ff/associar_ff.js ferramentas/paybox_ff/testes.html ferramentas/paybox_ff/testes.js ferramentas/paybox_ff/rodar_testes.py
git commit -m "Paybox F&F: esqueleto, funcoes puras e bancada de testes no Edge

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Leitura das respostas do Superlógica

**Files:**
- Modify: `ferramentas/paybox_ff/associar_ff.js`: novas funções logo acima do `return {` final, e novos nomes no objeto `logica`.
- Modify: `ferramentas/paybox_ff/testes.js`: novos casos no fim do arquivo.

**Interfaces:**
- Consumes: `soDigitos`, `paraNumero` e `CNPJ_FF` (Task 1).
- Produces (todos em `logica`):
  - `ErroInesperado`: classe de erro. Todo erro que deve **parar** o processo é dela, e a mensagem é para leigos.
  - `conferirResposta(texto) -> objeto`. Lança `ErroInesperado` quando o corpo é HTML (mensagem com `"Sessão expirou"`), quando não é JSON (`"Resposta que não é JSON"`) e quando `status` é diferente de `"200"` (mensagem com `"status <x>"`).
  - `notasDaFila(json) -> [{ envelopeId, arquivo, tipo, emitente, numero }]`. Aceita `data` como lista e `data.envelopes` como lista, com itens que trazem `metadata` direto ou dentro de `sources[]`. `{ sources: [] }` vira zero notas. Qualquer outra forma lança `ErroInesperado`.
  - `totalDePaginas(json) -> number | null`.
  - `eNotaDaFF(nota) -> boolean`: `tipo === "nfse"` e `emitente === CNPJ_FF`.
  - `dadosDoEnvelope(json)` devolve `{ envelopeId, notas: [{ tipo, numero, valor, valorBruto }], idCondominio, nomeCondominio, emitente, idFornecedor, classificacao: { tributaria, idTributaria, servico, idServico } }`, ou lança `ErroInesperado`.
  - `despesasDaResposta(json) -> [despesa]`: os objetos crus do Superlógica, depois de conferidos os campos que o resto usa. Lança `ErroInesperado` se faltar algum.
  - Fixtures em `testes.js`: `FIX.fila(itens, totalPaginas)`, `FIX.nota(envelopeId, numero, extra)`, `FIX.envelope(id, numero, valor, extra)` e `FIX.despesa(extra)`. As Tasks 3 e 4 usam.

- [ ] **Step 1: Escrever as fixtures e os testes**

Acrescente ao fim de `ferramentas/paybox_ff/testes.js`:

```js
// ---- fixtures (formato das respostas capturadas; dados inventados) ----

var FIX = {
  // Lista da fila no formato de index.json: data é uma lista de arquivos.
  fila: function (itens, totalPaginas) {
    var data = totalPaginas == null ? itens : { totaldepaginas: String(totalPaginas), envelopes: itens };
    return { status: "200", msg: "", data: data, session: "NAO-USAR" };
  },
  nota: function (envelopeId, numero, extra) {
    var m = { filename: "PGR 10004 Klosters.pdf", "document-type": "nfse",
              "document-recipient": "13736666000154", "invoice-number": numero };
    for (var k in (extra || {})) m[k] = extra[k];
    return { envelope_id: envelopeId, file_id: "arq-" + envelopeId, metadata: m };
  },
  envelope: function (id, numero, valor, extra) {
    var d = {
      id: id,
      invoices: [{ amount: valor, amount_before_taxes: valor, type: "nfse", number: numero }],
      customer: { name: "CONDOMINIO DO EDIFICIO KLOSTERS", document_number: "01195716000154",
                  id_condominio_cond: "44", st_fantasia_cond: "KLOSTERS" },
      issuer: { document_number: "13736666000154", id_contato_con: "750",
                st_classificacao_tributaria: "99 Pessoas Jurídicas em Geral", id_classificacao_tributaria: "19",
                st_classificacao_servico_prestado: "100000006 Preparação de dados para processamento",
                id_classificacao_servico_prestado: "6" },
      st_complemento_apro: "ASSESSORIA OU CONSULTORIA DE QUALQUER NATUREZA . FATURA NO . 1"
    };
    for (var k in (extra || {})) d[k] = extra[k];
    return { status: "200", msg: "", data: d, session: "NAO-USAR" };
  },
  despesa: function (extra) {
    var d = { id_despesa_des: "56112", id_parcela_pdes: "57606", st_documento_des: "12048",
              st_cpf_con: "13736666000154", id_contato_con: "750", vl_valor_pdes: "84.65",
              dt_despesa_des: "08/24/2026 00:00:00", fl_modelotrabalho_des: "2", id_tipo_doc: "1",
              id_forma_pag: "8", st_complemento_pdes: "Exames Médicos", arquivos: [],
              st_fantasia_cond: "KLOSTERS", st_nome_con: "F  F Consultoria e Serviços Médicos" };
    for (var k in (extra || {})) d[k] = extra[k];
    return d;
  }
};

// ---- leitura das respostas ----

teste("resposta 200 é aceita e devolvida como objeto", function () {
  igual(L.conferirResposta('{"status":"200","msg":"","data":[]}').data, []);
});

teste("HTML no lugar de JSON é sessão expirada", async function () {
  var e = await lanca(function () { L.conferirResposta("  <!DOCTYPE html><html>login</html>"); }, "Sessão expirou");
  verdade(e instanceof L.ErroInesperado, "tem que ser ErroInesperado");
});

teste("corpo que não é JSON para, sem repetir o corpo na mensagem", async function () {
  var e = await lanca(function () { L.conferirResposta("segredo-qualquer"); }, "não é JSON");
  verdade(String(e.message).indexOf("segredo") < 0, "mensagem não pode ecoar o corpo");
});

teste("status diferente de 200 para, com o status e a msg", async function () {
  await lanca(function () { L.conferirResposta('{"status":"500","msg":"Falhou"}'); }, "status 500");
  await lanca(function () { L.conferirResposta('{"status":"500","msg":"Falhou"}'); }, "Falhou");
});

teste("fila no formato lista (index.json)", function () {
  var notas = L.notasDaFila(FIX.fila([FIX.nota("env-1", "11882")]));
  igual(notas, [{ envelopeId: "env-1", arquivo: "PGR 10004 Klosters.pdf", tipo: "nfse",
                  emitente: "13736666000154", numero: "11882" }]);
});

teste("fila no formato data.envelopes com sources", function () {
  var item = { envelope_id: "env-2", sources: [{ metadata: { filename: "a.pdf", "document-type": "nfse",
               "document-recipient": "13.736.666/0001-54", "invoice-number": "7" } }] };
  var json = FIX.fila([item], 3);
  igual(L.notasDaFila(json)[0].envelopeId, "env-2");
  igual(L.notasDaFila(json)[0].emitente, "13736666000154");
  igual(L.totalDePaginas(json), 3);
});

teste("página vazia { sources: [] } vira zero notas", function () {
  igual(L.notasDaFila(FIX.fila([{ sources: [] }], 1)), []);
});

teste("fila sem total de páginas devolve null", function () {
  igual(L.totalDePaginas(FIX.fila([])), null);
});

teste("fila num formato desconhecido para", async function () {
  await lanca(function () { L.notasDaFila({ status: "200", data: "x" }); }, "formato");
  await lanca(function () { L.notasDaFila(FIX.fila([{ outra: 1 }])); }, "formato");
  await lanca(function () { L.notasDaFila(FIX.fila([{ metadata: {} }])); }, "sem identificação");
});

teste("só NFS-e emitida pela F&F entra", function () {
  verdade(L.eNotaDaFF({ tipo: "nfse", emitente: "13736666000154" }));
  verdade(!L.eNotaDaFF({ tipo: "outro", emitente: "13736666000154" }));
  verdade(!L.eNotaDaFF({ tipo: "nfse", emitente: "11222333000181" }));
});

teste("dados do envelope: nota, condomínio, fornecedor e classificação", function () {
  var e = L.dadosDoEnvelope(FIX.envelope("env-1", "12048", "84.65"));
  igual(e.envelopeId, "env-1");
  igual(e.notas, [{ tipo: "nfse", numero: "12048", valor: 84.65, valorBruto: 84.65 }]);
  igual(e.idCondominio, "44");
  igual(e.nomeCondominio, "KLOSTERS");
  igual(e.emitente, "13736666000154");
  igual(e.idFornecedor, "750");
  igual(e.classificacao, { tributaria: "99 Pessoas Jurídicas em Geral", idTributaria: "19",
                           servico: "100000006 Preparação de dados para processamento", idServico: "6" });
});

teste("envelope sem condomínio identificado devolve idCondominio vazio", function () {
  var json = FIX.envelope("env-1", "1", "10.00");
  json.data.customer.id_condominio_cond = "";
  igual(L.dadosDoEnvelope(json).idCondominio, "");
});

teste("envelope num formato desconhecido para", async function () {
  await lanca(function () { L.dadosDoEnvelope({ status: "200", data: { id: "x" } }); }, "formato");
});

teste("despesas: lista conferida campo a campo", function () {
  var lista = L.despesasDaResposta({ status: "200", data: [FIX.despesa()] });
  igual(lista.length, 1);
  igual(lista[0].id_parcela_pdes, "57606");
});

teste("despesa sem um campo usado pelo vínculo para", async function () {
  var d = FIX.despesa(); delete d.id_parcela_pdes;
  await lanca(function () { L.despesasDaResposta({ status: "200", data: [d] }); }, "id_parcela_pdes");
  await lanca(function () { L.despesasDaResposta({ status: "200", data: {} }); }, "formato");
});
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: código 1. Os 9 testes da Task 1 passam, e os novos falham com `L.conferirResposta is not a function` e parecidos.

- [ ] **Step 3: Implementar a leitura**

Em `associar_ff.js`, logo acima do `return {` final, acrescente:

```js
  // ---- leitura das respostas ----

  //  Tudo que deve PARAR o processo é ErroInesperado. Se o Superlógica mudar
  //  o formato de uma resposta, o certo é parar, nunca "pular" e seguir.
  class ErroInesperado extends Error {
    constructor(mensagem) { super(mensagem); this.name = "ErroInesperado"; }
  }

  //  A mensagem nunca repete o corpo da resposta: ele pode trazer dados da
  //  sessão.
  function conferirResposta(texto) {
    var corpo = String(texto == null ? "" : texto).trim();
    if (corpo.charAt(0) === "<") {
      throw new ErroInesperado("Sessão expirou — entre de novo no Superlógica e rode outra vez.");
    }
    var json;
    try { json = JSON.parse(corpo); }
    catch (e) { throw new ErroInesperado("O Superlógica mandou uma resposta que não é JSON."); }
    if (!json || typeof json !== "object" || String(json.status) !== "200") {
      var status = json && typeof json === "object" ? json.status : "?";
      var msg = json && typeof json === "object" && json.msg ? " (" + json.msg + ")" : "";
      throw new ErroInesperado("O Superlógica respondeu status " + status + msg + ".");
    }
    return json;
  }

  //  A lista da fila chega em dois formatos: `data` como lista de arquivos
  //  (é o index.json exportado) ou `data.envelopes`, com os arquivos em
  //  `sources`. A lista só serve para achar os envelopes; quem manda é o
  //  getenvelope.
  function notasDaFila(json) {
    var dados = json && json.data, lista;
    if (Array.isArray(dados)) lista = dados;
    else if (dados && Array.isArray(dados.envelopes)) lista = dados.envelopes;
    else throw new ErroInesperado("A lista da fila veio num formato desconhecido.");
    var notas = [];
    lista.forEach(function (item) {
      if (item && item.metadata) { notas.push(notaDoItem(item, item)); return; }
      if (item && Array.isArray(item.sources)) {
        item.sources.forEach(function (fonte) {
          if (!fonte || !fonte.metadata) throw new ErroInesperado("Um item da fila veio num formato desconhecido.");
          notas.push(notaDoItem(fonte, item));
        });
        return;
      }
      throw new ErroInesperado("Um item da fila veio num formato desconhecido.");
    });
    return notas;
  }

  function notaDoItem(fonte, envelope) {
    var m = fonte.metadata;
    var id = fonte.envelope_id || envelope.envelope_id || envelope.id;
    if (!id) throw new ErroInesperado("Um arquivo da fila veio sem identificação.");
    return {
      envelopeId: String(id),
      arquivo: m.filename || "",
      tipo: m["document-type"] || "",
      emitente: soDigitos(m["document-recipient"]),
      numero: m["invoice-number"] || ""
    };
  }

  function totalDePaginas(json) {
    var dados = json && json.data;
    if (!dados || Array.isArray(dados) || dados.totaldepaginas == null) return null;
    var n = parseInt(dados.totaldepaginas, 10);
    return isFinite(n) ? n : null;
  }

  function eNotaDaFF(nota) {
    return nota.tipo === "nfse" && nota.emitente === CNPJ_FF;
  }

  //  O getenvelope é a fonte de verdade da nota: número, valor, condomínio
  //  e fornecedor. A classificação da Reinf sai do cadastro do fornecedor
  //  (bloco issuer), exatamente como a tela de associação manual faz.
  function dadosDoEnvelope(json) {
    var d = json && json.data;
    if (!d || !Array.isArray(d.invoices) || !d.customer || !d.issuer || !d.id) {
      throw new ErroInesperado("Os dados de uma nota da fila vieram num formato desconhecido.");
    }
    return {
      envelopeId: String(d.id),
      notas: d.invoices.map(function (n) {
        return { tipo: n.type || "", numero: n.number || "",
                 valor: paraNumero(n.amount), valorBruto: paraNumero(n.amount_before_taxes) };
      }),
      idCondominio: d.customer.id_condominio_cond ? String(d.customer.id_condominio_cond) : "",
      nomeCondominio: d.customer.st_fantasia_cond || d.customer.name || "",
      emitente: soDigitos(d.issuer.document_number),
      idFornecedor: d.issuer.id_contato_con ? String(d.issuer.id_contato_con) : "",
      classificacao: {
        tributaria: d.issuer.st_classificacao_tributaria || "",
        idTributaria: d.issuer.id_classificacao_tributaria || "",
        servico: d.issuer.st_classificacao_servico_prestado || "",
        idServico: d.issuer.id_classificacao_servico_prestado || ""
      }
    };
  }

  //  Os campos que o casamento e o vínculo usam. Faltando qualquer um, o
  //  formato mudou e o certo é parar.
  var CAMPOS_DESPESA = ["id_despesa_des", "id_parcela_pdes", "st_documento_des", "st_cpf_con",
                        "id_contato_con", "vl_valor_pdes", "dt_despesa_des", "fl_modelotrabalho_des",
                        "id_tipo_doc", "id_forma_pag", "arquivos"];

  function despesasDaResposta(json) {
    if (!json || !Array.isArray(json.data)) {
      throw new ErroInesperado("A busca de despesas veio num formato desconhecido.");
    }
    json.data.forEach(function (d) {
      CAMPOS_DESPESA.forEach(function (campo) {
        if (!d || !(campo in d)) {
          throw new ErroInesperado("A busca de despesas veio sem o campo " + campo + ".");
        }
      });
    });
    return json.data;
  }
```

E acrescente ao objeto `logica`, depois de `valoresIguais: valoresIguais` (não esqueça a vírgula):

```js
      ErroInesperado: ErroInesperado,
      conferirResposta: conferirResposta,
      notasDaFila: notasDaFila,
      totalDePaginas: totalDePaginas,
      eNotaDaFF: eNotaDaFF,
      dadosDoEnvelope: dadosDoEnvelope,
      despesasDaResposta: despesasDaResposta
```

- [ ] **Step 4: Rodar e ver passar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 24 ok, 0 falhas`, código 0.

- [ ] **Step 5: Commit**

```bash
git add ferramentas/paybox_ff/associar_ff.js ferramentas/paybox_ff/testes.js
git commit -m "Paybox F&F: leitura e conferencia das respostas do Superlogica

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Casamento, corpo do vínculo e relatório

**Files:**
- Modify: `ferramentas/paybox_ff/associar_ff.js`: novas funções logo acima do `return {` final, e novos nomes em `logica`.
- Modify: `ferramentas/paybox_ff/testes.js`: novos casos no fim.

**Interfaces:**
- Consumes: `soDigitos`, `normalizarNumeroDocumento`, `paraNumero`, `valoresIguais`, `valorComPonto`, `valorComVirgula`, `dataUsSemHora` e `CNPJ_FF` (Task 1); `dadosDoEnvelope` e `FIX` (Task 2).
- Produces (em `logica`):
  - `candidatasDoEnvelope(envelope, despesas) -> [despesa]`: despesas da F&F com o mesmo número de documento, sem repetir parcela (`id_parcela_pdes`).
  - `escolherDespesa(envelope, despesas)` devolve `{ acao: "associar", despesa }` ou `{ acao: "fora", motivo }`. As mensagens são exatamente as da tabela do spec.
  - `montarCorpoVinculo(envelope, despesa) -> [[campo, valor], ...]`: todos os valores são strings, na ordem da captura.
  - `gerarCsv(linhas) -> string`, com BOM, separador `;` e quebra `\r\n`. `linhas` é uma lista de `{ arquivo, nota, condominio, valor, resultado, despesa, motivo }`.
  - `resumoPorMotivo(linhas) -> { associadas, associaria, fora, motivos: { motivo: n } }`.

- [ ] **Step 1: Escrever os testes**

Acrescente ao fim de `ferramentas/paybox_ff/testes.js`:

```js
// ---- casamento ----

function envelopePadrao(extra) {
  return L.dadosDoEnvelope(FIX.envelope("env-1", "12048", "84.65", extra));
}

teste("uma despesa da F&F com o mesmo documento, sem anexo e mesmo valor: associa", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa()]);
  igual(r.acao, "associar");
  igual(r.despesa.id_despesa_des, "56112");
});

teste("documento comparado sem zeros à esquerda", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa({ st_documento_des: "0012048" })]);
  igual(r.acao, "associar");
});

teste("despesa de outro fornecedor ou outro documento não é candidata", function () {
  var r = L.escolherDespesa(envelopePadrao(), [
    FIX.despesa({ st_cpf_con: "11222333000181" }),
    FIX.despesa({ st_documento_des: "13062", id_parcela_pdes: "58385" })
  ]);
  igual(r, { acao: "fora", motivo: "Despesa não encontrada no período" });
});

teste("nenhuma despesa: fora", function () {
  igual(L.escolherDespesa(envelopePadrao(), []), { acao: "fora", motivo: "Despesa não encontrada no período" });
});

teste("duas parcelas com o mesmo documento: fora", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa(), FIX.despesa({ id_parcela_pdes: "99999" })]);
  igual(r, { acao: "fora", motivo: "Mais de uma despesa com o documento 12048" });
});

teste("a mesma parcela vinda de duas buscas conta uma vez só", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa(), FIX.despesa()]);
  igual(r.acao, "associar");
});

teste("despesa que já tem anexo: fora", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa({ arquivos: [{ id_arquivo_arq: "1" }] })]);
  igual(r, { acao: "fora", motivo: "Despesa já tem anexo" });
});

teste("valor diferente: fora, com os dois valores", function () {
  var r = L.escolherDespesa(envelopePadrao(), [FIX.despesa({ vl_valor_pdes: "80.00" })]);
  igual(r, { acao: "fora", motivo: "Valor da despesa (80,00) difere da nota (84,65)" });
});

teste("valor igual ao bruto da nota também vale", function () {
  var json = FIX.envelope("env-1", "12048", "80.00");
  json.data.invoices[0].amount_before_taxes = "84.65";
  var r = L.escolherDespesa(L.dadosDoEnvelope(json), [FIX.despesa()]);
  igual(r.acao, "associar");
});

teste("envelope sem condomínio: fora", function () {
  var json = FIX.envelope("env-1", "12048", "84.65");
  json.data.customer.id_condominio_cond = "";
  igual(L.escolherDespesa(L.dadosDoEnvelope(json), [FIX.despesa()]),
        { acao: "fora", motivo: "Paybox não identificou o condomínio" });
});

teste("envelope com duas notas: fora", function () {
  var json = FIX.envelope("env-1", "12048", "84.65");
  json.data.invoices.push({ amount: "1.00", amount_before_taxes: "1.00", type: "nfse", number: "1" });
  igual(L.escolherDespesa(L.dadosDoEnvelope(json), [FIX.despesa()]),
        { acao: "fora", motivo: "Envelope com 2 notas" });
});

teste("nota que não é NFS-e da F&F: fora", function () {
  var json = FIX.envelope("env-1", "12048", "84.65");
  json.data.issuer.document_number = "11222333000181";
  igual(L.escolherDespesa(L.dadosDoEnvelope(json), [FIX.despesa()]),
        { acao: "fora", motivo: "Não é NFS-e da F&F" });
});

// ---- corpo do vínculo ----

teste("corpo do vínculo: campos na ordem da tela, copiados da despesa", function () {
  var corpo = L.montarCorpoVinculo(envelopePadrao(), FIX.despesa());
  igual(corpo, [
    ["ID_CONDOMINIO_COND", "44"],
    ["ID_ENVELOPE_ENV", "env-1"],
    ["ID_DESPESA_DES", "56112"],
    ["ID_PARCELA_PDES", "57606"],
    ["ID_CONTATO_CON", "750"],
    ["ID_ENVELOPEMENSAGEIRO_PDES", "env-1"],
    ["VL_VALOR_PDES", "84.65"],
    ["ST_DOCUMENTO_DES", "12048"],
    ["DT_DESPESA_DES", "08/24/2026"],
    ["FL_MODELOTRABALHO_DES", "2"],
    ["ST_COMPLEMENTO_APRO", "Exames Médicos"],
    ["ST_VALORPRIMEIRAPARCELA", "84.65"],
    ["ST_CPF_CNPJ_FORNECEDOR_ENV", "13736666000154"],
    ["VALIDAR_CPF_CNPJ", "1"],
    ["ST_CLASSIFICACAO_TRIBUTARIA", "99 Pessoas Jurídicas em Geral"],
    ["ID_CLASSIFICACAOTRIBUTARIA_DES", "19"],
    ["ST_CLASSIFICACAO_SERVICO_PRESTADO", "100000006 Preparação de dados para processamento"],
    ["ID_CLASSSERVICOPRESTADO_DES", "6"],
    ["ID_TIPO_DOC", "1"],
    ["VL_DOCUMENTO_DES", "84.65"],
    ["ID_FORMA_PAG", "8"],
    ["ST_FANTASIA_COND", "KLOSTERS"],
    ["ST_NOME_CON", "F  F Consultoria e Serviços Médicos"],
    ["VL_VALOR_ENV", "84,65"],
    ["FL_MARCAR_PARA_LANCAMENTO", "0"],
    ["despesa", "on"],
    ["salvar", "Vincular"]
  ]);
});

teste("complemento vazio da despesa vai vazio, nunca a descrição da nota", function () {
  var corpo = L.montarCorpoVinculo(envelopePadrao(), FIX.despesa({ st_complemento_pdes: "" }));
  var complemento = corpo.filter(function (p) { return p[0] === "ST_COMPLEMENTO_APRO"; })[0][1];
  igual(complemento, "");
  verdade(JSON.stringify(corpo).indexOf("ASSESSORIA") < 0, "descrição da nota não pode ir no corpo");
});

teste("todo valor do corpo é texto", function () {
  var corpo = L.montarCorpoVinculo(envelopePadrao(), FIX.despesa({ st_complemento_pdes: null }));
  corpo.forEach(function (p) { verdade(typeof p[1] === "string", p[0] + " não é texto"); });
});

// ---- relatório ----

var LINHAS = [
  { arquivo: "PGR 10004 Klosters.pdf", nota: "12048", condominio: "KLOSTERS", valor: 84.65,
    resultado: "associada", despesa: "56112", motivo: "" },
  { arquivo: "PCMSO; com \"aspas\".pdf", nota: "7", condominio: "SAN REMO", valor: 10,
    resultado: "fora", despesa: "", motivo: "Despesa já tem anexo" },
  { arquivo: "c.pdf", nota: "8", condominio: "SAN REMO", valor: 10,
    resultado: "fora", despesa: "", motivo: "Despesa já tem anexo" }
];

teste("CSV com BOM, ; e vírgula decimal", function () {
  var csv = L.gerarCsv(LINHAS);
  verdade(csv.charAt(0) === "\uFEFF", "falta o BOM");
  var linhas = csv.slice(1).split("\r\n");
  igual(linhas[0], "Arquivo;Nota;Condomínio;Valor;Resultado;Despesa;Motivo");
  igual(linhas[1], "PGR 10004 Klosters.pdf;12048;KLOSTERS;84,65;associada;56112;");
  igual(linhas[2], "\"PCMSO; com \"\"aspas\"\".pdf\";7;SAN REMO;10,00;fora;;Despesa já tem anexo");
  igual(linhas[linhas.length - 1], "");
});

teste("resumo conta por resultado e por motivo", function () {
  igual(L.resumoPorMotivo(LINHAS), { associadas: 1, associaria: 0, fora: 2,
                                     motivos: { "Despesa já tem anexo": 2 } });
});
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: código 1. Os 24 anteriores passam, e os novos falham com `L.escolherDespesa is not a function` e parecidos.

- [ ] **Step 3: Implementar**

Em `associar_ff.js`, logo acima do `return {` final, acrescente:

```js
  // ---- casamento ----

  //  O número do documento é a chave; o valor é só conferência. A mesma
  //  parcela pode vir de duas buscas (valor líquido e bruto) e conta uma vez.
  function candidatasDoEnvelope(envelope, despesas) {
    var numero = normalizarNumeroDocumento(envelope.notas.length ? envelope.notas[0].numero : "");
    var vistas = {}, candidatas = [];
    despesas.forEach(function (d) {
      if (soDigitos(d.st_cpf_con) !== CNPJ_FF) return;
      if (normalizarNumeroDocumento(d.st_documento_des) !== numero) return;
      if (vistas[d.id_parcela_pdes]) return;
      vistas[d.id_parcela_pdes] = true;
      candidatas.push(d);
    });
    return candidatas;
  }

  function escolherDespesa(envelope, despesas) {
    if (envelope.notas.length !== 1) {
      return { acao: "fora", motivo: "Envelope com " + envelope.notas.length + " notas" };
    }
    var nota = envelope.notas[0];
    if (nota.tipo !== "nfse" || envelope.emitente !== CNPJ_FF) {
      return { acao: "fora", motivo: "Não é NFS-e da F&F" };
    }
    if (!envelope.idCondominio) {
      return { acao: "fora", motivo: "Paybox não identificou o condomínio" };
    }
    var candidatas = candidatasDoEnvelope(envelope, despesas);
    if (candidatas.length === 0) return { acao: "fora", motivo: "Despesa não encontrada no período" };
    if (candidatas.length > 1) {
      return { acao: "fora", motivo: "Mais de uma despesa com o documento " + normalizarNumeroDocumento(nota.numero) };
    }
    var despesa = candidatas[0];
    if (despesa.arquivos && despesa.arquivos.length > 0) return { acao: "fora", motivo: "Despesa já tem anexo" };
    var valorDespesa = paraNumero(despesa.vl_valor_pdes);
    if (!valoresIguais(valorDespesa, nota.valor) && !valoresIguais(valorDespesa, nota.valorBruto)) {
      return { acao: "fora", motivo: "Valor da despesa (" + textoValor(valorDespesa) +
               ") difere da nota (" + textoValor(nota.valor) + ")" };
    }
    return { acao: "associar", despesa: despesa };
  }

  function textoValor(n) { return n == null ? "?" : valorComVirgula(n); }

  // ---- corpo do vínculo ----

  //  Mesmos campos e mesma ordem que a tela do Paybox manda ao clicar em
  //  Vincular (capturado em 2026-09-27). O que é da despesa sai da despesa,
  //  nunca da nota: a tela manda a descrição da nota como complemento e
  //  forma de pagamento 0, e repetir isso 1.800 vezes seria arriscar
  //  reescrever os lançamentos. A classificação da Reinf sai do cadastro do
  //  fornecedor, como na tela.
  function montarCorpoVinculo(envelope, despesa) {
    var valor = paraNumero(despesa.vl_valor_pdes);
    var pares = [
      ["ID_CONDOMINIO_COND", envelope.idCondominio],
      ["ID_ENVELOPE_ENV", envelope.envelopeId],
      ["ID_DESPESA_DES", despesa.id_despesa_des],
      ["ID_PARCELA_PDES", despesa.id_parcela_pdes],
      ["ID_CONTATO_CON", despesa.id_contato_con],
      ["ID_ENVELOPEMENSAGEIRO_PDES", envelope.envelopeId],
      ["VL_VALOR_PDES", valorComPonto(valor)],
      ["ST_DOCUMENTO_DES", despesa.st_documento_des],
      ["DT_DESPESA_DES", dataUsSemHora(despesa.dt_despesa_des)],
      ["FL_MODELOTRABALHO_DES", despesa.fl_modelotrabalho_des],
      ["ST_COMPLEMENTO_APRO", despesa.st_complemento_pdes],
      ["ST_VALORPRIMEIRAPARCELA", valorComPonto(valor)],
      ["ST_CPF_CNPJ_FORNECEDOR_ENV", CNPJ_FF],
      ["VALIDAR_CPF_CNPJ", "1"],
      ["ST_CLASSIFICACAO_TRIBUTARIA", envelope.classificacao.tributaria],
      ["ID_CLASSIFICACAOTRIBUTARIA_DES", envelope.classificacao.idTributaria],
      ["ST_CLASSIFICACAO_SERVICO_PRESTADO", envelope.classificacao.servico],
      ["ID_CLASSSERVICOPRESTADO_DES", envelope.classificacao.idServico],
      ["ID_TIPO_DOC", despesa.id_tipo_doc],
      ["VL_DOCUMENTO_DES", valorComPonto(valor)],
      ["ID_FORMA_PAG", despesa.id_forma_pag],
      ["ST_FANTASIA_COND", despesa.st_fantasia_cond || envelope.nomeCondominio],
      ["ST_NOME_CON", despesa.st_nome_con],
      ["VL_VALOR_ENV", valorComVirgula(valor)],
      ["FL_MARCAR_PARA_LANCAMENTO", "0"],
      ["despesa", "on"],
      ["salvar", "Vincular"]
    ];
    return pares.map(function (p) { return [p[0], p[1] == null ? "" : String(p[1])]; });
  }

  // ---- relatório ----

  var COLUNAS_RELATORIO = [
    ["arquivo", "Arquivo"], ["nota", "Nota"], ["condominio", "Condomínio"], ["valor", "Valor"],
    ["resultado", "Resultado"], ["despesa", "Despesa"], ["motivo", "Motivo"]
  ];

  //  ";" e BOM: é assim que o Excel em português abre o CSV com acento e
  //  colunas certas.
  function gerarCsv(linhas) {
    function celula(valor) {
      var s = valor == null ? "" : String(valor);
      return /[";\r\n]/.test(s) ? "\"" + s.replace(/"/g, "\"\"") + "\"" : s;
    }
    var saida = [COLUNAS_RELATORIO.map(function (c) { return c[1]; }).join(";")];
    linhas.forEach(function (l) {
      saida.push(COLUNAS_RELATORIO.map(function (c) {
        var v = l[c[0]];
        return celula(c[0] === "valor" && typeof v === "number" ? valorComVirgula(v) : v);
      }).join(";"));
    });
    return "\uFEFF" + saida.join("\r\n") + "\r\n";
  }

  function resumoPorMotivo(linhas) {
    var r = { associadas: 0, associaria: 0, fora: 0, motivos: {} };
    linhas.forEach(function (l) {
      if (l.resultado === "associada") r.associadas++;
      else if (l.resultado === "associaria") r.associaria++;
      else {
        r.fora++;
        r.motivos[l.motivo] = (r.motivos[l.motivo] || 0) + 1;
      }
    });
    return r;
  }
```

E acrescente ao objeto `logica`:

```js
      candidatasDoEnvelope: candidatasDoEnvelope,
      escolherDespesa: escolherDespesa,
      montarCorpoVinculo: montarCorpoVinculo,
      gerarCsv: gerarCsv,
      resumoPorMotivo: resumoPorMotivo
```

- [ ] **Step 4: Rodar e ver passar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 41 ok, 0 falhas`, código 0.

- [ ] **Step 5: Commit**

```bash
git add ferramentas/paybox_ff/associar_ff.js ferramentas/paybox_ff/testes.js
git commit -m "Paybox F&F: casamento nota-despesa, corpo do vinculo e relatorio CSV

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Orquestração (`processar`) e transporte

**Files:**
- Modify: `ferramentas/paybox_ff/associar_ff.js`: novas funções acima do `return {` final, `processar` no objeto devolvido e a leitura de `ambiente` no topo.
- Modify: `ferramentas/paybox_ff/testes.js`: novos casos no fim.

**Interfaces:**
- Consumes: tudo das Tasks 1–3; `ambienteFalso` e `transporteFalso` da bancada.
- Produces:
  - O objeto devolvido ganha `processar(opcoes, aoProgredir) -> Promise<{ linhas, parouPor }>`.
    - `opcoes = { modo: "simular" | "associar", deVencimento: "DD/MM/AAAA", ateVencimento: "DD/MM/AAAA", limite: number | null, sinalParar: { parar: boolean } }`.
    - `parouPor` é `null` quando o processo termina normalmente; senão é a mensagem de por que parou.
    - Período inválido **rejeita** a promessa com `Error("Período de vencimento inválido — use DD/MM/AAAA.")` antes de qualquer chamada.
  - `aoProgredir(estado)` é chamado com `{ fase: "fila", pagina, totalPaginas }` ou `{ fase: "notas", atual, total, associadas, fora }`.
  - Cada `linha` é `{ arquivo, nota, condominio, valor, resultado: "associada" | "associaria" | "fora", despesa, motivo }`.
  - Contrato do `ambiente.transporte(caminho, pares) -> Promise<string>`: `caminho` é relativo a `/condor/atual/` (ex.: `"ocr/getenvelope?idEnvelope=x"`), `pares` é uma lista `[[campo, valor], ...]`, e a promessa devolve o corpo cru da resposta.

- [ ] **Step 1: Escrever os testes**

Acrescente ao fim de `ferramentas/paybox_ff/testes.js`:

```js
// ---- orquestração ----

function parametrosDaLeitura(pares) {
  return JSON.parse(pares[0][1]).params[0];
}
function paginaDoCaminho(caminho) {
  var m = /pagina=(\d+)/.exec(caminho); return m ? +m[1] : null;
}

// Monta as rotas de um Superlógica de mentira.
//   paginasFila: lista de listas de itens (uma por página), servida como
//     index.json (data lista) e com página vazia depois da última.
//   envelopes: { envelopeId: FIX.envelope(...) }
//   despesasPorValor: { "84.65": [despesa...] }  (a pesquisa é pelo valor)
function superlogicaFalso(paginasFila, envelopes, despesasPorValor, extras) {
  extras = extras || {};
  return transporteFalso(function (caminho, pares) {
    if (caminho.indexOf("ocr/getenvelopes") === 0) {
      if (extras.fila) return extras.fila(caminho);
      var p = paginaDoCaminho(caminho);
      return FIX.fila(paginasFila[p - 1] || []);
    }
    if (caminho.indexOf("ocr/getenvelope?") === 0) {
      var id = decodeURIComponent(/idEnvelope=([^&]+)/.exec(caminho)[1]);
      if (extras.envelope) { var r = extras.envelope(id); if (r) return r; }
      return envelopes[id];
    }
    if (caminho === "despesas/index") {
      var params = parametrosDaLeitura(pares);
      var todas = despesasPorValor[params.pesquisa] || [];
      var inicio = (params.pagina - 1) * +params.itensPorPagina;
      return { status: "200", msg: "", data: todas.slice(inicio, inicio + +params.itensPorPagina) };
    }
    if (caminho === "ocr/vincularenvelopeadespesa") return { status: "200", msg: "", data: [] };
    throw new Error("rota inesperada: " + caminho);
  });
}

function opcoes(extra) {
  var o = { modo: "simular", deVencimento: "01/10/2026", ateVencimento: "31/10/2026",
            limite: null, sinalParar: { parar: false } };
  for (var k in (extra || {})) o[k] = extra[k];
  return o;
}

// Dois PDFs da F&F (um casa, outro sem despesa) e um arquivo de outro emitente.
function cenarioBasico(extras) {
  return superlogicaFalso(
    [[FIX.nota("env-1", "12048"),
      FIX.nota("env-9", "1", { "document-recipient": "11222333000181" }),
      FIX.nota("env-2", "12049")]],
    { "env-1": FIX.envelope("env-1", "12048", "84.65"),
      "env-2": FIX.envelope("env-2", "12049", "78.61") },
    { "84.65": [FIX.despesa()] },
    extras);
}

function chamadasA(t, prefixo) {
  return t.chamadas.filter(function (c) { return c.caminho.indexOf(prefixo) === 0; });
}

teste("simulação: associaria o que casa, relata o resto, ignora outro emitente", async function () {
  var t = cenarioBasico();
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.parouPor, null);
  igual(r.linhas.map(function (l) { return [l.arquivo, l.nota, l.resultado, l.despesa, l.motivo]; }), [
    ["PGR 10004 Klosters.pdf", "12048", "associaria", "56112", ""],
    ["PGR 10004 Klosters.pdf", "12049", "fora", "", "Despesa não encontrada no período"]
  ]);
  igual(chamadasA(t, "ocr/vincularenvelopeadespesa").length, 0, "simulação nunca vincula");
  igual(chamadasA(t, "ocr/getenvelope?idEnvelope=env-9").length, 0, "outro emitente nem é aberto");
});

teste("associar: vincula com o corpo montado da despesa e espera entre associações", async function () {
  var t = cenarioBasico(), esperas = [];
  var amb = ambienteFalso(t, { esperar: function (ms) { esperas.push(ms); return Promise.resolve(); } });
  var r = await criarPayboxFF(amb).processar(opcoes({ modo: "associar" }));
  var vinculos = chamadasA(t, "ocr/vincularenvelopeadespesa");
  igual(vinculos.length, 1);
  igual(vinculos[0].pares, L.montarCorpoVinculo(L.dadosDoEnvelope(FIX.envelope("env-1", "12048", "84.65")), FIX.despesa()));
  igual(r.linhas[0].resultado, "associada");
  igual(esperas, [1000]);
});

teste("busca de despesas: condomínio, fornecedor, valor e período em MM/DD/AAAA", async function () {
  var t = cenarioBasico();
  await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  var busca = chamadasA(t, "despesas/index")[0];
  var params = parametrosDaLeitura(busca.pares);
  igual(params, { comStatus: "todas", idCondominio: "44", FAVORECIDOS: ["750"], pesquisa: "84.65",
                  tipoFiltroData: "periodo", dtInicio: "10/01/2026", dtFim: "10/31/2026",
                  itensPorPagina: "50", pagina: 1 });
  igual(JSON.parse(busca.pares[0][1]).url, "https://admin1.superlogica.net/condor/atual/despesas/index");
});

teste("sem candidata pelo valor líquido, tenta o bruto", async function () {
  var env = FIX.envelope("env-1", "12048", "80.00");
  env.data.invoices[0].amount_before_taxes = "84.65";
  var t = superlogicaFalso([[FIX.nota("env-1", "12048")]], { "env-1": env }, { "84.65": [FIX.despesa()] });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas[0].resultado, "associaria");
  igual(chamadasA(t, "despesas/index").map(function (c) { return parametrosDaLeitura(c.pares).pesquisa; }),
        ["80.00", "84.65"]);
});

teste("busca de despesas lê todas as páginas", async function () {
  var outras = [];
  for (var i = 0; i < 50; i++) outras.push(FIX.despesa({ st_documento_des: "9" + i, id_parcela_pdes: "p" + i }));
  var t = superlogicaFalso([[FIX.nota("env-1", "12048")]],
                           { "env-1": FIX.envelope("env-1", "12048", "84.65") },
                           { "84.65": outras.concat([FIX.despesa()]) });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas[0].resultado, "associaria");
  igual(chamadasA(t, "despesas/index").length, 2);
});

teste("fila em várias páginas, até vir uma vazia", async function () {
  var t = superlogicaFalso([[FIX.nota("env-1", "12048")], [FIX.nota("env-2", "12049")]],
                           { "env-1": FIX.envelope("env-1", "12048", "84.65"),
                             "env-2": FIX.envelope("env-2", "12049", "78.61") }, {});
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas.length, 2);
  igual(chamadasA(t, "ocr/getenvelopes").length, 3);
});

teste("fila que ignora o número da página não vira laço infinito", async function () {
  var t = superlogicaFalso([], { "env-1": FIX.envelope("env-1", "12048", "84.65") }, {},
    { fila: function () { return FIX.fila([FIX.nota("env-1", "12048")]); } });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas.length, 1);
  igual(chamadasA(t, "ocr/getenvelopes").length, 2);
});

teste("fila com total de páginas para no total", async function () {
  var t = superlogicaFalso([], { "env-1": FIX.envelope("env-1", "12048", "84.65") }, {},
    { fila: function (caminho) {
        return paginaDoCaminho(caminho) === 1 ? FIX.fila([{ envelope_id: "env-1",
          sources: [FIX.nota("env-1", "12048")] }], 1) : FIX.fila([], 1);
      } });
  await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(chamadasA(t, "ocr/getenvelopes").length, 1);
});

teste("limite: para depois de N associações, dizendo por quê", async function () {
  var t = superlogicaFalso([[FIX.nota("env-1", "12048"), FIX.nota("env-2", "12049")]],
    { "env-1": FIX.envelope("env-1", "12048", "84.65"), "env-2": FIX.envelope("env-2", "12049", "84.65") },
    { "84.65": [FIX.despesa(), FIX.despesa({ st_documento_des: "12049", id_parcela_pdes: "2", id_despesa_des: "2" })] });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes({ modo: "associar", limite: 1 }));
  igual(chamadasA(t, "ocr/vincularenvelopeadespesa").length, 1);
  igual(r.linhas.length, 1);
  verdade(/Limite de 1/.test(r.parouPor), "parouPor: " + r.parouPor);
});

teste("resposta inesperada para tudo e guarda o que já foi feito", async function () {
  var t = cenarioBasico({ envelope: function (id) {
    return id === "env-2" ? { status: "500", msg: "Erro interno" } : null;
  } });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas.length, 1);
  verdade(/status 500/.test(r.parouPor), "parouPor: " + r.parouPor);
});

teste("sessão expirada na fila para com a mensagem de entrar de novo", async function () {
  var t = cenarioBasico({ fila: function () { return "<html>login</html>"; } });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas, []);
  verdade(/Sessão expirou/.test(r.parouPor), "parouPor: " + r.parouPor);
});

teste("falha de rede também para, em vez de estourar", async function () {
  var t = transporteFalso(function () { throw new Error("HTTP 502 em ocr/getenvelopes"); });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  verdade(/HTTP 502/.test(r.parouPor), "parouPor: " + r.parouPor);
});

teste("botão Parar interrompe entre uma nota e outra", async function () {
  var t = cenarioBasico(), o = opcoes();
  var r = await criarPayboxFF(ambienteFalso(t)).processar(o, function (estado) {
    if (estado.fase === "notas" && estado.atual === 1) o.sinalParar.parar = true;
  });
  igual(r.linhas.length, 1);
  igual(r.parouPor, "Interrompido por você.");
});

teste("período inválido recusa antes de qualquer chamada", async function () {
  var t = cenarioBasico();
  await lanca(function () { return criarPayboxFF(ambienteFalso(t)).processar(opcoes({ ateVencimento: "31/02/2026" })); },
              "Período de vencimento inválido");
  igual(t.chamadas.length, 0);
});

teste("progresso informa a fila e cada nota", async function () {
  var estados = [];
  await criarPayboxFF(ambienteFalso(cenarioBasico())).processar(opcoes(), function (e) { estados.push(e); });
  igual(estados[0], { fase: "fila", pagina: 1, totalPaginas: null });
  igual(estados[estados.length - 1], { fase: "notas", atual: 2, total: 2, associadas: 1, fora: 1 });
});

teste("nota repetida em duas páginas é processada uma vez", async function () {
  var t = superlogicaFalso([[FIX.nota("env-1", "12048")], [FIX.nota("env-1", "12048"), FIX.nota("env-2", "12049")]],
    { "env-1": FIX.envelope("env-1", "12048", "84.65"), "env-2": FIX.envelope("env-2", "12049", "78.61") },
    { "84.65": [FIX.despesa()] });
  var r = await criarPayboxFF(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas.length, 2);
});
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: código 1. Os 41 anteriores passam, e os novos falham com `criarPayboxFF(...).processar is not a function`.

- [ ] **Step 3: Ler o `ambiente` no topo**

Em `associar_ff.js`, logo depois da linha `var TOLERANCIA_VALOR = 0.005;`, acrescente:

```js
  var ITENS_POR_PAGINA_DESPESAS = 50;
  var LIMITE_PAGINAS_FILA = 1000;

  //  O que muda entre o navegador de verdade e os testes. No favorito,
  //  instalar.html chama criarPayboxFF({ janela: window }) e o resto usa
  //  os padrões abaixo.
  ambiente = ambiente || {};
  var janela = ambiente.janela;
  var origem = ambiente.origem || janela.location.origin;
  var hostname = ambiente.hostname || janela.location.hostname;
  var transporte = ambiente.transporte || transportePadrao;
  var esperar = ambiente.esperar || function (ms) {
    return new Promise(function (fim) { janela.setTimeout(fim, ms); });
  };
```

- [ ] **Step 4: Implementar o transporte e `processar`**

Logo acima do `return {` final, acrescente:

```js
  // ---- transporte e orquestração ----

  //  Mesmo formato que a tela usa: POST de formulário, com cabeçalho
  //  X-Requested-With, na mesma origem da página. A sessão é a da própria
  //  página (credentials same-origin); nada aqui lê cookie.
  async function transportePadrao(caminho, pares) {
    var resposta = await janela.fetch(origem + "/condor/atual/" + caminho, {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "X-Requested-With": "XMLHttpRequest" },
      body: new URLSearchParams(pares).toString()
    });
    if (resposta.status !== 200) {
      throw new ErroInesperado("O Superlógica respondeu HTTP " + resposta.status + " em " + caminho.split("?")[0] + ".");
    }
    return resposta.text();
  }

  //  Leituras levam um único campo "json" com os parâmetros e a URL
  //  completa, como a tela manda.
  async function ler(caminho, params) {
    var pares = [["json", JSON.stringify({ params: [params || {}], url: origem + "/condor/atual/" + caminho })]];
    return conferirResposta(await transporte(caminho, pares));
  }

  async function lerFila(aoProgredir) {
    var notas = [], vistos = {};
    for (var pagina = 1; ; pagina++) {
      if (pagina > LIMITE_PAGINAS_FILA) throw new ErroInesperado("A fila passou de " + LIMITE_PAGINAS_FILA + " páginas.");
      var json = await ler("ocr/getenvelopes?idCondominio=-2&pagina=" + pagina, {});
      var total = totalDePaginas(json);
      aoProgredir({ fase: "fila", pagina: pagina, totalPaginas: total });
      var daPagina = notasDaFila(json), novas = 0;
      daPagina.forEach(function (n) {
        if (vistos[n.envelopeId]) return;
        vistos[n.envelopeId] = true;
        novas++;
        notas.push(n);
      });
      //  Página vazia, total alcançado, ou uma página só com o que já foi
      //  visto (servidor ignorando o número da página): fim da fila.
      if (daPagina.length === 0 || novas === 0 || (total != null && pagina >= total)) break;
    }
    return notas;
  }

  async function buscarDespesas(envelope, valor, dtInicio, dtFim) {
    var todas = [];
    for (var pagina = 1; ; pagina++) {
      var json = await ler("despesas/index", {
        comStatus: "todas", idCondominio: envelope.idCondominio, FAVORECIDOS: [envelope.idFornecedor],
        pesquisa: valorComPonto(valor), tipoFiltroData: "periodo", dtInicio: dtInicio, dtFim: dtFim,
        itensPorPagina: String(ITENS_POR_PAGINA_DESPESAS), pagina: pagina
      });
      var lista = despesasDaResposta(json);
      todas = todas.concat(lista);
      if (lista.length < ITENS_POR_PAGINA_DESPESAS) return todas;
    }
  }

  //  Primeiro pelo valor da nota; sem candidata e com valor bruto
  //  diferente, tenta pelo bruto.
  async function despesasDoEnvelope(envelope, dtInicio, dtFim) {
    if (!envelope.idCondominio || envelope.notas.length !== 1) return [];
    var nota = envelope.notas[0];
    var despesas = nota.valor == null ? [] : await buscarDespesas(envelope, nota.valor, dtInicio, dtFim);
    if (candidatasDoEnvelope(envelope, despesas).length === 0 &&
        nota.valorBruto != null && !valoresIguais(nota.valorBruto, nota.valor)) {
      despesas = despesas.concat(await buscarDespesas(envelope, nota.valorBruto, dtInicio, dtFim));
    }
    return despesas;
  }

  async function processar(opcoes, aoProgredir) {
    aoProgredir = aoProgredir || function () {};
    var dtInicio = dataBrParaUs(opcoes.deVencimento), dtFim = dataBrParaUs(opcoes.ateVencimento);
    if (!dtInicio || !dtFim) throw new Error("Período de vencimento inválido — use DD/MM/AAAA.");
    var associar = opcoes.modo === "associar";
    var sinal = opcoes.sinalParar || { parar: false };
    var linhas = [], associadas = 0, fora = 0;
    try {
      var notas = (await lerFila(aoProgredir)).filter(eNotaDaFF);
      for (var i = 0; i < notas.length; i++) {
        if (sinal.parar) return { linhas: linhas, parouPor: "Interrompido por você." };
        if (opcoes.limite && associadas >= opcoes.limite) {
          return { linhas: linhas, parouPor: "Limite de " + opcoes.limite + " associações atingido." };
        }
        var nota = notas[i];
        var envelope = dadosDoEnvelope(await ler("ocr/getenvelope?idEnvelope=" + encodeURIComponent(nota.envelopeId), {}));
        var escolha = escolherDespesa(envelope, await despesasDoEnvelope(envelope, dtInicio, dtFim));
        var linha = {
          arquivo: nota.arquivo,
          nota: envelope.notas.length === 1 ? envelope.notas[0].numero : nota.numero,
          condominio: envelope.nomeCondominio,
          valor: envelope.notas.length === 1 ? envelope.notas[0].valor : null,
          resultado: "fora", despesa: "", motivo: ""
        };
        if (escolha.acao === "associar") {
          linha.despesa = escolha.despesa.id_despesa_des;
          if (associar) {
            conferirResposta(await transporte("ocr/vincularenvelopeadespesa", montarCorpoVinculo(envelope, escolha.despesa)));
            linha.resultado = "associada";
            await esperar(PAUSA_ENTRE_ASSOCIACOES_MS);
          } else {
            linha.resultado = "associaria";
          }
          associadas++;
        } else {
          linha.motivo = escolha.motivo;
          fora++;
        }
        linhas.push(linha);
        aoProgredir({ fase: "notas", atual: i + 1, total: notas.length, associadas: associadas, fora: fora });
      }
      return { linhas: linhas, parouPor: null };
    } catch (e) {
      //  Qualquer erro para o processo e devolve o que já foi feito: o
      //  relatório parcial diz exatamente até onde foi.
      return { linhas: linhas, parouPor: (e && e.message) || "Erro desconhecido." };
    }
  }
```

E no `return {` final, acrescente `processar` ao lado de `logica`:

```js
  return {
    processar: processar,
    logica: {
      ...
```

Mantenha o conteúdo existente de `logica` onde está, sem mudança.

- [ ] **Step 5: Rodar e ver passar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 57 ok, 0 falhas`, código 0.

- [ ] **Step 6: Commit**

```bash
git add ferramentas/paybox_ff/associar_ff.js ferramentas/paybox_ff/testes.js
git commit -m "Paybox F&F: percorre a fila, busca a despesa e associa, com simulacao e parada

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Painel na página (`abrir`)

**Files:**
- Modify: `ferramentas/paybox_ff/associar_ff.js`: leitura de `confirmar`/`avisar`/`baixar` no topo, funções do painel acima do `return {`, e `abrir` e `execucaoAtual` no objeto devolvido.
- Modify: `ferramentas/paybox_ff/testes.js`: novos casos no fim.

**Interfaces:**
- Consumes: `processar`, `gerarCsv` e `resumoPorMotivo` (Tasks 3–4).
- Produces:
  - `api.abrir() -> HTMLElement | null`: devolve o painel. Fora de `*.superlogica.net` chama `ambiente.avisar(texto)` e devolve `null`. Um segundo `abrir()` devolve o painel já aberto.
  - `api.execucaoAtual`: a promessa da execução disparada pelo botão Iniciar, que os testes aguardam.
  - Ids no DOM: `paybox-ff-painel`, `paybox-ff-de`, `paybox-ff-ate`, `paybox-ff-limite`, `paybox-ff-modo-simular`, `paybox-ff-modo-associar`, `paybox-ff-iniciar`, `paybox-ff-parar`, `paybox-ff-fechar`, `paybox-ff-status`, `paybox-ff-resumo` e `paybox-ff-baixar`.
  - `ambiente.baixar(nomeArquivo, texto)` recebe um nome terminado em `.csv`.

- [ ] **Step 1: Escrever os testes**

Acrescente ao fim de `ferramentas/paybox_ff/testes.js`:

```js
// ---- painel ----

function limparPainel() {
  var p = document.getElementById("paybox-ff-painel");
  if (p) p.remove();
}
function campo(id) { return document.getElementById("paybox-ff-" + id); }

teste("fora do Superlógica só avisa, sem painel", function () {
  limparPainel();
  var avisos = [];
  var api = criarPayboxFF(ambienteFalso(null, { hostname: "exemplo.com.br",
                                                avisar: function (t) { avisos.push(t); } }));
  igual(api.abrir(), null);
  igual(document.getElementById("paybox-ff-painel"), null);
  verdade(/Superlógica/.test(avisos[0]), "aviso: " + avisos[0]);
});

teste("clicar no favorito duas vezes não abre dois painéis", function () {
  limparPainel();
  var api = criarPayboxFF(ambienteFalso(cenarioBasico()));
  var p1 = api.abrir(), p2 = criarPayboxFF(ambienteFalso(cenarioBasico())).abrir();
  verdade(p1 === p2, "tem que devolver o mesmo painel");
  igual(document.querySelectorAll("#paybox-ff-painel").length, 1);
  limparPainel();
});

teste("painel: cantos retos e cobalto só no Iniciar", function () {
  limparPainel();
  criarPayboxFF(ambienteFalso(cenarioBasico())).abrir();
  var painel = document.getElementById("paybox-ff-painel");
  var cobaltos = Array.prototype.filter.call(painel.querySelectorAll("*"), function (el) {
    return getComputedStyle(el).backgroundColor === "rgb(0, 59, 142)";
  });
  igual(cobaltos.map(function (el) { return el.id; }), ["paybox-ff-iniciar"]);
  Array.prototype.forEach.call([painel].concat(Array.prototype.slice.call(painel.querySelectorAll("*"))), function (el) {
    igual(getComputedStyle(el).borderTopLeftRadius, "0px", "raio em " + (el.id || el.tagName));
  });
  verdade(campo("modo-simular").checked, "simular vem marcado");
  limparPainel();
});

teste("simular pelo painel mostra o resumo e libera o relatório", async function () {
  limparPainel();
  var baixados = [];
  var api = criarPayboxFF(ambienteFalso(cenarioBasico(), {
    baixar: function (nome, texto) { baixados.push([nome, texto]); } }));
  api.abrir();
  campo("de").value = "01/10/2026"; campo("ate").value = "31/10/2026";
  campo("iniciar").click();
  await api.execucaoAtual;
  var resumo = campo("resumo").textContent;
  verdade(/Associaria: 1/.test(resumo), "resumo: " + resumo);
  verdade(/Fora: 1/.test(resumo), "resumo: " + resumo);
  verdade(/Despesa não encontrada no período: 1/.test(resumo), "resumo: " + resumo);
  verdade(!campo("baixar").disabled, "baixar tem que ficar liberado");
  campo("baixar").click();
  igual(baixados.length, 1);
  verdade(/^paybox-ff-simulacao-\d{4}-\d{2}-\d{2}-\d{4}\.csv$/.test(baixados[0][0]), "nome: " + baixados[0][0]);
  verdade(baixados[0][1].charAt(0) === "\uFEFF", "CSV sem BOM");
  limparPainel();
});

teste("associar pede confirmação; recusar não chama nada", async function () {
  limparPainel();
  var t = cenarioBasico();
  var api = criarPayboxFF(ambienteFalso(t, { confirmar: function () { return false; } }));
  api.abrir();
  campo("de").value = "01/10/2026"; campo("ate").value = "31/10/2026";
  campo("modo-associar").checked = true;
  campo("iniciar").click();
  await api.execucaoAtual;
  igual(t.chamadas.length, 0);
  limparPainel();
});

teste("período inválido aparece no painel, sem chamar nada", async function () {
  limparPainel();
  var t = cenarioBasico();
  var api = criarPayboxFF(ambienteFalso(t));
  api.abrir();
  campo("de").value = "2026-10-01"; campo("ate").value = "31/10/2026";
  campo("iniciar").click();
  await api.execucaoAtual;
  verdade(/Período de vencimento inválido/.test(campo("status").textContent), campo("status").textContent);
  igual(t.chamadas.length, 0);
  limparPainel();
});

teste("limite que não é número positivo é recusado", async function () {
  limparPainel();
  var t = cenarioBasico();
  var api = criarPayboxFF(ambienteFalso(t));
  api.abrir();
  campo("de").value = "01/10/2026"; campo("ate").value = "31/10/2026";
  campo("limite").value = "abc";
  campo("iniciar").click();
  await api.execucaoAtual;
  verdade(/Limite/.test(campo("status").textContent), campo("status").textContent);
  igual(t.chamadas.length, 0);
  limparPainel();
});

teste("parada mostra o motivo no resumo", async function () {
  limparPainel();
  var api = criarPayboxFF(ambienteFalso(cenarioBasico({ fila: function () { return "<html></html>"; } })));
  api.abrir();
  campo("de").value = "01/10/2026"; campo("ate").value = "31/10/2026";
  campo("iniciar").click();
  await api.execucaoAtual;
  verdade(/Parou: Sessão expirou/.test(campo("resumo").textContent), campo("resumo").textContent);
  limparPainel();
});

teste("fechar remove o painel", function () {
  limparPainel();
  criarPayboxFF(ambienteFalso(cenarioBasico())).abrir();
  campo("fechar").click();
  igual(document.getElementById("paybox-ff-painel"), null);
});
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: código 1. Os 57 anteriores passam, e os novos falham com `api.abrir is not a function`.

- [ ] **Step 3: Ler o resto do `ambiente` no topo**

Em `associar_ff.js`, logo depois do bloco `var esperar = ...;` da Task 4, acrescente:

```js
  var confirmar = ambiente.confirmar || function (texto) { return janela.confirm(texto); };
  var avisar = ambiente.avisar || function (texto) { janela.alert(texto); };
  var baixar = ambiente.baixar || baixarPadrao;
```

- [ ] **Step 4: Implementar o painel**

Logo acima do `return {` final, acrescente:

```js
  // ---- painel ----

  var ID_PAINEL = "paybox-ff-painel";
  //  Paleta do TEMA_CLARO do Codificador. Cobalto só no botão Iniciar.
  var COR = { fundo: "#FAFAF9", superficie: "#F5F5F4", borda: "#E7E5E4", texto: "#1C1917",
              secundario: "#57534E", acento: "#003B8E", sobreAcento: "#FFFFFF" };
  var FONTE = "13px 'IBM Plex Sans','Segoe UI',system-ui,sans-serif";

  function baixarPadrao(nome, texto) {
    var doc = janela.document;
    var url = janela.URL.createObjectURL(new Blob([texto], { type: "text/csv;charset=utf-8" }));
    var a = doc.createElement("a");
    a.href = url; a.download = nome;
    doc.body.appendChild(a); a.click(); a.remove();
    janela.setTimeout(function () { janela.URL.revokeObjectURL(url); }, 1000);
  }

  function nomeRelatorio(modo) {
    var d = new Date();
    function dois(n) { return (n < 10 ? "0" : "") + n; }
    return "paybox-ff-" + (modo === "associar" ? "associacao" : "simulacao") + "-" +
           d.getFullYear() + "-" + dois(d.getMonth() + 1) + "-" + dois(d.getDate()) + "-" +
           dois(d.getHours()) + dois(d.getMinutes()) + ".csv";
  }

  function textoResumo(r, modo) {
    var s = resumoPorMotivo(r.linhas), partes = [];
    partes.push(modo === "associar" ? "Associadas: " + s.associadas : "Associaria: " + s.associaria);
    partes.push("Fora: " + s.fora);
    Object.keys(s.motivos).forEach(function (m) { partes.push("  " + m + ": " + s.motivos[m]); });
    if (r.parouPor) partes.push("Parou: " + r.parouPor);
    return partes.join("\n");
  }

  function textoProgresso(estado, inicio) {
    if (estado.fase === "fila") {
      return "Lendo a fila… página " + estado.pagina + (estado.totalPaginas ? " de " + estado.totalPaginas : "");
    }
    var decorrido = (Date.now() - inicio) / 1000;
    var faltam = Math.ceil((estado.total - estado.atual) * (decorrido / estado.atual) / 60);
    return "Nota " + estado.atual + " de " + estado.total + " — associadas " + estado.associadas +
           " · fora " + estado.fora + (estado.atual < estado.total ? " — faltam ~" + faltam + " min" : "");
  }

  function montarPainel(doc) {
    function el(tag, estilo, texto) {
      var e = doc.createElement(tag);
      e.style.cssText = "border-radius:0;box-sizing:border-box;font:" + FONTE + ";color:" + COR.texto + ";" + (estilo || "");
      if (texto != null) e.textContent = texto;
      return e;
    }
    function botao(id, texto, primario) {
      var b = el("button", primario
        ? "background:" + COR.acento + ";color:" + COR.sobreAcento + ";border:1px solid " + COR.acento + ";padding:6px 14px;cursor:pointer;"
        : "background:transparent;border:1px solid " + COR.texto + ";padding:6px 14px;cursor:pointer;", texto);
      b.id = "paybox-ff-" + id; b.type = "button";
      return b;
    }
    function entrada(id, rotulo, dica) {
      var bloco = el("label", "display:block;margin:0 0 8px 0;color:" + COR.secundario + ";", rotulo);
      var i = el("input", "display:block;width:100%;margin-top:2px;padding:4px 6px;background:" + COR.superficie + ";border:1px solid " + COR.borda + ";");
      i.id = "paybox-ff-" + id; i.placeholder = dica || "";
      bloco.appendChild(i);
      return bloco;
    }
    function opcao(id, valor, rotulo, marcado) {
      var bloco = el("label", "margin-right:12px;cursor:pointer;");
      var r = el("input", "margin-right:4px;");
      r.type = "radio"; r.name = "paybox-ff-modo"; r.id = "paybox-ff-" + id; r.value = valor; r.checked = !!marcado;
      bloco.appendChild(r); bloco.appendChild(doc.createTextNode(rotulo));
      return bloco;
    }

    var painel = el("div", "position:fixed;right:16px;bottom:16px;width:360px;z-index:2147483647;" +
                            "background:" + COR.fundo + ";border:1px solid " + COR.texto + ";padding:16px;" +
                            "box-shadow:0 2px 12px rgba(0,0,0,.15);");
    painel.id = ID_PAINEL;

    var topo = el("div", "display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;");
    topo.appendChild(el("strong", "font-weight:600;", "Associar notas da F&F"));
    var fechar = botao("fechar", "×", false);
    fechar.style.padding = "0 8px";
    topo.appendChild(fechar);
    painel.appendChild(topo);

    painel.appendChild(entrada("de", "Vencimento das despesas — de", "DD/MM/AAAA"));
    painel.appendChild(entrada("ate", "até", "DD/MM/AAAA"));
    painel.appendChild(entrada("limite", "Associar no máximo (vazio = todas)", ""));

    var modos = el("div", "margin:4px 0 12px 0;");
    modos.appendChild(opcao("modo-simular", "simular", "Simular", true));
    modos.appendChild(opcao("modo-associar", "associar", "Associar", false));
    painel.appendChild(modos);

    var acoes = el("div", "display:flex;gap:8px;margin-bottom:12px;");
    var iniciar = botao("iniciar", "Iniciar", true);
    var parar = botao("parar", "Parar", false);
    parar.disabled = true;
    acoes.appendChild(iniciar); acoes.appendChild(parar);
    painel.appendChild(acoes);

    var status = el("div", "color:" + COR.secundario + ";min-height:18px;margin-bottom:8px;");
    status.id = "paybox-ff-status";
    painel.appendChild(status);

    var resumo = el("pre", "white-space:pre-wrap;margin:0 0 12px 0;background:" + COR.superficie + ";padding:8px;border:1px solid " + COR.borda + ";");
    resumo.id = "paybox-ff-resumo";
    painel.appendChild(resumo);

    var baixarBotao = botao("baixar", "Baixar relatório (CSV)", false);
    baixarBotao.disabled = true;
    painel.appendChild(baixarBotao);

    return painel;
  }

  var api = null;

  function abrir() {
    var doc = janela.document;
    if (!/(^|\.)superlogica\.net$/.test(hostname)) {
      avisar("Abra o Superlógica (endereço terminado em superlogica.net) e clique no favorito de novo.");
      return null;
    }
    var existente = doc.getElementById(ID_PAINEL);
    if (existente) return existente;
    var painel = montarPainel(doc);
    doc.body.appendChild(painel);
    ligarPainel(painel);
    return painel;
  }

  function ligarPainel(painel) {
    function campo(id) { return painel.querySelector("#paybox-ff-" + id); }
    var sinal = null, ultimo = null, ultimoModo = "simular";

    campo("fechar").addEventListener("click", function () {
      if (sinal) sinal.parar = true;
      painel.remove();
    });
    campo("parar").addEventListener("click", function () {
      if (sinal) { sinal.parar = true; campo("status").textContent = "Parando depois da nota atual…"; }
    });
    campo("baixar").addEventListener("click", function () {
      if (ultimo) baixar(nomeRelatorio(ultimoModo), gerarCsv(ultimo.linhas));
    });
    campo("iniciar").addEventListener("click", function () {
      api.execucaoAtual = iniciar();
    });

    async function iniciar() {
      var modo = campo("modo-associar").checked ? "associar" : "simular";
      var textoLimite = campo("limite").value.trim();
      var limite = textoLimite ? Number(textoLimite) : null;
      if (textoLimite && !(Number.isInteger(limite) && limite > 0)) {
        campo("status").textContent = "Limite precisa ser um número inteiro maior que zero.";
        return;
      }
      if (modo === "associar" && !confirmar(
          "Associar de verdade as notas da F&F às despesas encontradas?\n\n" +
          "Recomendado: rodar Simular antes e conferir o relatório.")) {
        return;
      }
      sinal = { parar: false };
      campo("iniciar").disabled = true; campo("parar").disabled = false; campo("baixar").disabled = true;
      campo("resumo").textContent = "";
      var inicio = Date.now();
      try {
        var r = await processar({ modo: modo, deVencimento: campo("de").value, ateVencimento: campo("ate").value,
                                  limite: limite, sinalParar: sinal },
                                function (estado) { campo("status").textContent = textoProgresso(estado, inicio); });
        ultimo = r; ultimoModo = modo;
        campo("resumo").textContent = textoResumo(r, modo);
        campo("status").textContent = r.parouPor ? "Parou." : "Terminou.";
        campo("baixar").disabled = false;
      } catch (e) {
        campo("status").textContent = (e && e.message) || "Erro desconhecido.";
      } finally {
        sinal = null;
        campo("iniciar").disabled = false; campo("parar").disabled = true;
      }
    }
  }
```

Troque o `return {` final para guardar o objeto em `api`, que os tratadores de clique usam para expor `execucaoAtual`:

```js
  api = {
    abrir: abrir,
    execucaoAtual: null,
    processar: processar,
    logica: {
      ...
    }
  };
  return api;
}
```

Mantenha o conteúdo existente de `logica` sem mudança; o `...` acima só indica isso.

- [ ] **Step 5: Rodar e ver passar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 66 ok, 0 falhas`, código 0.

Se o teste "cantos retos e cobalto só no Iniciar" falhar por um elemento de formulário com raio padrão do navegador (ex.: o `input` radio), **não** relaxe o teste. Garanta que o `el()` aplica `border-radius:0` a ele (ele já aplica a todo elemento criado por `el`) e confira se o elemento foi criado sem `el`.

- [ ] **Step 6: Olhar o painel de verdade**

Se as ferramentas do navegador embutido estiverem disponíveis, abra `ferramentas/paybox_ff/testes.html` nele (preview_start com url `file:///C:/Users/Desktop/Pictures/projeto/ferramentas/paybox_ff/testes.html`) e, no console da página (javascript_tool), rode:

```js
criarPayboxFF(ambienteFalso(cenarioBasico())).abrir(); "aberto"
```

Tire um screenshot e confira:
- painel no canto inferior direito;
- cantos retos;
- só o Iniciar em cobalto;
- textos legíveis e nada cortado.

Feche o painel. Descreva no relatório o que viu. Sem as ferramentas do navegador, pule este passo e diga isso no relatório: o controlador faz a conferência visual.

- [ ] **Step 7: Commit**

```bash
git add ferramentas/paybox_ff/associar_ff.js ferramentas/paybox_ff/testes.js
git commit -m "Paybox F&F: painel na pagina com simular, associar, parar e relatorio

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Instalador do favorito, LEIAME e CLAUDE.md

**Files:**
- Create: `ferramentas/paybox_ff/instalar.html`
- Create: `ferramentas/paybox_ff/LEIAME.md`
- Modify: `CLAUDE.md`: nova seção logo antes de `## Planilha de Despesas do Superlógica (aba 3, v6.14.0)`.
- Modify: `ferramentas/paybox_ff/testes.js`: um caso novo.

**Interfaces:**
- Consumes: `criarPayboxFF` (Tasks 1–5).
- Produces: o favorito em si. O `href` é `javascript:` + `encodeURIComponent("(" + criarPayboxFF.toString() + ")({janela:window}).abrir();void 0;")`.

- [ ] **Step 1: Teste de que a função sobrevive ao `toString()`**

Acrescente ao fim de `ferramentas/paybox_ff/testes.js`:

```js
// ---- favorito ----

teste("criarPayboxFF recriada a partir do próprio texto funciona sozinha", async function () {
  //  É exatamente o que o favorito faz: o texto da função, sem nada em volta.
  //  Se ela depender de algo de fora, isto quebra aqui e não no Superlógica.
  var recriada = new Function("return (" + criarPayboxFF.toString() + ");")();
  var t = cenarioBasico();
  var r = await recriada(ambienteFalso(t)).processar(opcoes());
  igual(r.linhas.length, 2);
  igual(r.parouPor, null);
});
```

- [ ] **Step 2: Rodar**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 67 ok, 0 falhas`. O teste passa de primeira se as Tasks 1–5 respeitaram a regra de nada fora da função. Se falhar, a mensagem mostra o que ela referencia de fora: corrija em `associar_ff.js` e registre no relatório.

- [ ] **Step 3: Criar o instalador**

Crie `ferramentas/paybox_ff/instalar.html`:

```html
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Instalar — Associar notas da F&amp;F</title>
<style>
  body { font: 15px 'IBM Plex Sans','Segoe UI',system-ui,sans-serif; color: #1C1917; background: #FAFAF9;
         max-width: 720px; margin: 40px auto; padding: 0 16px; line-height: 1.5; }
  h1 { font-weight: 400; font-size: 28px; margin-bottom: 8px; }
  .favorito { display: inline-block; background: #003B8E; color: #FFFFFF; padding: 10px 18px;
              text-decoration: none; border-radius: 0; margin: 16px 0; }
  ol li { margin-bottom: 8px; }
  textarea { width: 100%; height: 90px; font: 12px Consolas, monospace; border: 1px solid #E7E5E4;
             background: #F5F5F4; border-radius: 0; box-sizing: border-box; }
  .aviso { border-left: 3px solid #1C1917; padding: 4px 12px; color: #57534E; }
</style>
</head>
<body>
<h1>Associar notas da F&amp;F no Paybox</h1>
<p>Um favorito que, clicado com o Superlógica aberto, associa cada nota fiscal da F&amp;F
   que está na fila do Paybox à sua despesa.</p>

<h2>Instalar (uma vez só)</h2>
<ol>
  <li>Mostre a barra de favoritos do navegador (Ctrl+Shift+B).</li>
  <li>Arraste este botão para a barra de favoritos:<br>
      <a class="favorito" id="favorito" href="#">Associar F&amp;F</a></li>
  <li>Se arrastar não funcionar: crie um favorito qualquer, edite, e cole no campo do
      endereço (URL) o texto abaixo.</li>
</ol>
<textarea id="codigo" readonly></textarea>

<h2>Usar, todo mês</h2>
<ol>
  <li>Suba as notas no Paybox e importe a planilha de despesas, como hoje.</li>
  <li>Abra o Superlógica e clique no favorito <strong>Associar F&amp;F</strong>.</li>
  <li>Informe o período de vencimento das despesas, deixe em <strong>Simular</strong> e clique
      em Iniciar. Baixe o relatório e confira.</li>
  <li>Mude para <strong>Associar</strong> e rode de novo.</li>
</ol>
<p class="aviso">Primeira vez: associe só <strong>1</strong> nota (campo "Associar no máximo")
   e confira a despesa no Superlógica — anexo, forma de pagamento, complemento e a
   classificação da Reinf — antes de rodar o lote inteiro.</p>
<p class="aviso">Quando o <code>associar_ff.js</code> for atualizado, abra esta página de novo e
   substitua o favorito antigo.</p>

<script src="associar_ff.js"></script>
<script>
  //  O favorito é a própria função criarPayboxFF, em texto, chamada com a
  //  janela da página onde ele for clicado.
  var codigo = "javascript:" + encodeURIComponent(
    "(" + criarPayboxFF.toString() + ")({janela:window}).abrir();void 0;");
  document.getElementById("favorito").href = codigo;
  document.getElementById("codigo").value = codigo;
</script>
</body>
</html>
```

- [ ] **Step 4: Conferir o instalador no navegador**

Se as ferramentas do navegador embutido estiverem disponíveis, abra `file:///C:/Users/Desktop/Pictures/projeto/ferramentas/paybox_ff/instalar.html` nele e confira pelo javascript_tool:

```js
[document.getElementById("favorito").href.slice(0, 40), document.getElementById("codigo").value.length]
```

Expected: o `href` começa com `javascript:(function%20criarPayboxFF` e o texto tem alguns milhares de caracteres. Tire um screenshot da página e descreva no relatório. Sem as ferramentas do navegador, pule e diga isso no relatório.

- [ ] **Step 5: Escrever o LEIAME**

Crie `ferramentas/paybox_ff/LEIAME.md`:

```markdown
# Associar notas da F&F no Paybox

Favorito do navegador que associa as NFS-e da F&F que estão na fila do Paybox
às suas despesas no Superlógica. **Não faz parte do Codificador** e não entra
no executável.

- Instalar e usar: abra `instalar.html` no navegador e siga as instruções.
- Testes: `python ferramentas/paybox_ff/rodar_testes.py` (roda `testes.html`
  no Edge sem janela).
- Por que existe e como decide: `docs/superpowers/specs/2026-09-27-associacao-paybox-ff-design.md`.

## O relatório

| Resultado | Quer dizer |
|---|---|
| associada | vinculada à despesa da coluna Despesa |
| associaria | (simulação) seria vinculada a essa despesa |
| fora | não foi tocada; a coluna Motivo diz por quê |

| Motivo | O que fazer |
|---|---|
| Despesa não encontrada no período | a despesa não foi importada, ou o vencimento está fora do período informado |
| Mais de uma despesa com o documento N | lançamento em dobro: resolver no Superlógica e associar à mão |
| Despesa já tem anexo | já foi associada antes; nada a fazer |
| Valor da despesa (X) difere da nota (Y) | conferir a nota: retenção ou lançamento com valor errado |
| Paybox não identificou o condomínio | associar à mão |
| Envelope com N notas / Não é NFS-e da F&F | arquivo fora do padrão; associar à mão |

## Quando parar

O favorito **para** em vez de seguir quando o Superlógica responde algo
inesperado (erro, formato diferente, sessão expirada). Rodar de novo é seguro:
o que já foi associado sai da fila. Se parar sempre no mesmo ponto com
"formato desconhecido", o Superlógica mudou a tela — é o `associar_ff.js` que
precisa de ajuste, a partir de um HAR novo da associação manual.

## Segurança

O código usa a sessão que a página do Superlógica já tem. Ele nunca lê, guarda
nem envia senha, cookie ou token, e só chama o próprio endereço do Superlógica.
HAR e `index.json` exportados do navegador carregam a sessão: não guarde no
repositório e apague depois de usar.
```

- [ ] **Step 6: Seção no CLAUDE.md**

Em `CLAUDE.md`, insira logo antes da linha `## Planilha de Despesas do Superlógica (aba 3, v6.14.0)`:

```markdown
## Associação das notas da F&F no Paybox — favorito do navegador (fora do Codificador)

A NFS-e não se associa sozinha no Paybox (ver a seção acima), e com ~1.800
notas da F&F por mês a associação manual virou gargalo. A saída foi um
**favorito do navegador** em `ferramentas/paybox_ff/`, que faz as mesmas
chamadas que a tela do Paybox faz quando a pessoa clica em Vincular.
**Fica fora do Codificador por decisão do usuário**: nada em `logica.py`, na
interface ou no executável.

**Dois caminhos testados antes, e descartados:**
- a coluna `etiqueta_paybox` da importação, com o link `sldocs.com.br` de cada
  nota: a despesa grava o link, mas a imagem continua na fila (teste com a
  NFS-e 11882, 2026-09-27);
- a API oficial, que depende de um App Token que só um administrador cria.

Do mesmo teste saiu a confirmação de que o importador **aceita a competência
no formato que o Codificador grava** (pendente da v6.21.0 para a F&F).

**Chave de casamento:** condomínio + fornecedor F&F + **número da NFS-e =
documento da despesa**; o valor é só conferência. Associa só com exatamente
uma candidata, sem anexo. Todo o resto vai para o relatório com o motivo.

**O que é da despesa sai da despesa.** A tela manda, ao vincular, a descrição
da nota como complemento e forma de pagamento 0. O favorito manda o
complemento e a forma de pagamento da própria despesa, para não depender do
comportamento (observado) de o Superlógica ignorar esses campos. A
classificação da Reinf vai como a tela manda, do cadastro do fornecedor.

**São chamadas internas, não a API documentada** — podem mudar sem aviso. Por
isso o código **para** diante de qualquer resposta que não reconheça, em vez
de pular. Spec: `docs/superpowers/specs/2026-09-27-associacao-paybox-ff-design.md`.
```

- [ ] **Step 7: Rodar tudo**

Run: `python ferramentas/paybox_ff/rodar_testes.py`
Expected: `RESUMO: 67 ok, 0 falhas`.

Run: `python -m unittest discover -s tests -p "test_*.py"`
Expected: `Ran 463 tests`, `OK` (o Codificador não mudou).

- [ ] **Step 8: Commit**

```bash
git add ferramentas/paybox_ff/instalar.html ferramentas/paybox_ff/LEIAME.md ferramentas/paybox_ff/testes.js CLAUDE.md
git commit -m "Paybox F&F: instalador do favorito, LEIAME e registro no CLAUDE.md

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Notas para quem executa

- **O primeiro uso real é do usuário**, na ordem do spec: simular a fila inteira, associar com limite 1 e conferir a despesa, depois ~20 e só então o lote. Nenhuma task roda nada contra o Superlógica de verdade.
- **Nunca abra, copie nem leia HAR ou `index.json` do usuário** para fazer fixture. As fixtures deste plano já têm o formato necessário.
- **Se a primeira chamada real falhar por falta de cabeçalho** (ex.: resposta HTML ou HTTP 403), o ajuste é só no `transportePadrao`. A captura mostra que a tela manda `X-Requested-With`, que já está incluído.
- **Contagem de testes esperada:** 9 → 24 → 41 → 57 → 66 → 67. A suíte Python continua em 463.
