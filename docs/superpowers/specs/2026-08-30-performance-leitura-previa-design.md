# Performance: leitura de PDF e prévia do documento — especificação

Data: 2026-08-30
Estado: aprovado, a implementar

## Objetivo

Reduzir o custo dos dois caminhos que o programa percorre com mais frequência:
**ler o texto de um PDF** (base da extração de NFS-e e da identificação por
CNPJ) e **renderizar a prévia do documento** no painel de resultado.

Os ganhos foram medidos antes de qualquer mudança, contra arquivos reais do
usuário — não são estimativa:

| Caminho | Hoje | Depois | Ganho |
|---|---|---|---|
| Ler texto de PDF | 117 ms/arq | 2,6 ms/arq | **44,7×** |
| → lote de 3347 notas (aba 2) | **7 min** | **9 s** | |
| Renderizar a prévia (410 px) | 202 ms | 26 ms | **7,7×** |
| Clique numa linha | congela a janela | não congela | |

## O que NÃO entra, e por quê

Medir primeiro derrubou a hipótese com que este trabalho começou. A suspeita
era a grade: `_estado_da_linha_protocolo` varre as três listas do resultado
para **cada** linha, e chama `_falta_id_sl`, que reconstrói o índice
código→CNPJ dos 772 condomínios a cada chamada. É de fato O(n²).

Só que a medição dá **52 ms para 600 linhas**, e os lotes reais têm ~15
arquivos. Otimizar isso seria trabalho sem retorno nenhum. **A grade fica como
está** — e este parágrafo existe para que a próxima pessoa que farejar o
O(n²) saiba que ele já foi medido e dispensado.

Também ficam de fora:

- **Imports preguiçosos** (adiar o `openpyxl`, 457 ms): economiza 450 ms uma
  vez por sessão. Não paga o risco de o `.spec` perder um import e matar um
  recurso em silêncio no executável — o mesmo modo de falha que o `CLAUDE.md`
  já registra sobre o `winocr`.
- **Paralelizar o OCR do lote**: já roda em thread com barra de progresso, e
  5,8 s para 15 protocolos não incomoda. Threads não ajudariam de todo modo:
  medido, 8 threads em `pypdf` deram 120 ms/arq contra 117 ms em série — é
  trabalho preso no GIL.

## 1. Trocar o extrator de texto

`extrair_texto_pdf` (`logica.py:357`) passa a usar PyMuPDF em vez de pypdf.

O **pypdf continua no projeto**: é ele que carimba os PDFs
(`logica.py:1258-1284`) e mede as páginas. A troca é exclusivamente na
**leitura de texto**. PyMuPDF já é dependência — nada novo para empacotar.

### Por que isto é seguro (a validação que sustenta a mudança)

Trocar o extrator de texto alimenta uma **planilha financeira**, então a
mudança foi validada contra os arquivos reais antes de ser proposta, na mesma
disciplina que pegou o bug do `candidatos_por_nome`:

- **1186 notas** (fevereiro/2026): 593 reconhecidas pelos dois extratores com
  **todos os campos idênticos**; 1 divergência; 0 reconhecidas por só um deles.
- **1188 PDFs** no caminho de identificação: **1161 CNPJs idênticos, 0
  divergências**.
- **0 arquivos** mudam de lado na regra dos 30 caracteres que decide cair para
  OCR — nenhum documento passa a ser tratado como escaneado, ou deixa de ser.

A única divergência em 1186 notas é o **PyMuPDF acertando**: em
`11047 Pallazo dei Visconti.pdf` o pypdf devolvia o nome do tomador truncado
(`PALAZZO DEI`, quebrado na linha) e o PyMuPDF lê inteiro
(`PALAZZO DEI VISCONTI`). Mais rápido e um pouco mais correto.

### Reserva quando o PyMuPDF falta

Se `FITZ_DISPONIVEL` for falso, a função **cai de volta no pypdf**. Assim ela
nunca fica sem resposta, e nesse cenário o comportamento é idêntico ao de
hoje. Custa três linhas e evita que um ambiente sem PyMuPDF perca a leitura de
texto inteira.

### Silenciar o ruído do MuPDF

Alguns PDFs fazem o MuPDF escrever `format error: cmsOpenProfileFromMem
failed` no stderr. É inofensivo, mas poluiria o `erros.log` e o log de
processamento. `fitz.TOOLS.mupdf_display_errors(False)` no import do módulo
resolve (API confirmada na versão em uso, PyMuPDF 1.28.0).

## 2. Prévia com largura exata

`renderizar_paginas_pdf` hoje rasteriza **sempre a 110 DPI** (910 px de
largura) e depois **reduz com LANCZOS** até a largura do painel — gera pixel
para jogar fora. Passa a rasterizar direto na largura pedida, com
`fitz.Matrix(z, z)` onde `z = largura / pagina.rect.width`.

O `Matrix` entrega a largura **exata** — conferido em 205, 410, 823 e 1640 px,
todos batendo no pixel. Isso é melhor do que ajustar o DPI, que é inteiro e
obrigaria a um resize de correção: assim o LANCZOS **some por completo**.

```
largura  205 px:  194 ms  →  20 ms   (9,8×)
largura  410 px:  202 ms  →  26 ms   (7,7×)   ← largura normal do painel
largura  820 px:  222 ms  →  78 ms   (2,8×)
```

### Corrige um defeito real de nitidez

Acima de ~220% de zoom a largura pedida passa dos 910 px do pixmap fixo, e
hoje o programa **amplia bitmap**. O `CLAUDE.md` afirma que "a 400% o texto
continua nítido" — hoje isso é falso. Com a rasterização na largura do alvo,
passa a ser verdade, e a documentação deixa de mentir.

## 3. Tirar o render da thread da interface

Hoje `_desenhar_previa` rasteriza **dentro do próprio clique**: a janela fica
congelada por 50–330 ms a cada linha selecionada, a cada dente da roda do zoom
e a cada ajuste do divisor. O lote já roda em thread; a prévia não.

Passa a:

1. **Cache cheio** → pinta na hora. É o caminho comum, e continua síncrono
   porque já é instantâneo.
2. **Cache vazio** → dispara a renderização numa thread e **mantém a imagem
   anterior na tela**. Mesmo princípio já adotado no arraste do divisor: não
   limpar antes da hora é o que evita a piscada.
3. A thread devolve **imagens PIL**; a thread principal converte em
   `PhotoImage` e pinta.

A divisão do passo 3 não é arbitrária: `ImageTk.PhotoImage` é objeto Tk e
**tem** que ser criado na thread principal. O que fica na thread é a parte
cara (rasterizar, 26–290 ms); o que sobra para a principal é a barata —
medido, 30 ms para 4 páginas, contra os 200 ms de hoje.

### Descartar resultado obsoleto

Um contador de geração é incrementado a cada pedido. Quando a thread termina,
o resultado só é pintado se a geração ainda for a corrente. Sem isso, clicar
depressa em várias linhas pintaria o documento errado na linha selecionada —
uma renderização lenta chegando depois de uma rápida.

Uma renderização por vez, com "o último pedido ganha": pedidos que chegam
durante uma renderização substituem o pendente em vez de enfileirar. Girar a
roda cinco dentes deve renderizar o zoom final, não os cinco.

## 4. Teto de memória no cache

O cache da prévia guarda **12 entradas**, contadas por unidade. Medido, o peso
de **uma** entrada:

| Largura | Memória (documento de 4 páginas) |
|---|---|
| 410 px | 2,9 MB |
| 1640 px | 45,7 MB |
| 3200 px | **173,8 MB** |

Janela maximizada com zoom em 400% num protocolo de 12 páginas passa de 1 GB.
**Isto já acontece hoje** — não é criado pela mudança, e é por isso que está
neste spec e não foi deixado de lado: seria irresponsável reescrever
exatamente esta função e passar ao largo.

O despejo passa a ser **por total de pixels**, não por contagem de entradas.
`LIMITE_CACHE_PREVIA` (12 entradas) é **substituído** por
`LIMITE_PIXELS_PREVIA`, não somado a ele — duas regras de despejo ao mesmo
tempo só confundiriam quem for depurar isso depois.

O teto é **48 milhões de pixels** (~144 MB em RGB). Esse número vem da
medição: no uso normal (410 px) uma entrada custa cerca de 1 M pixels, então
cabem umas 48 — bem mais que as 12 de hoje, o que torna o vaivém entre linhas
já vistas instantâneo. No extremo (3200 px, 58 M pixels) uma única entrada
estoura o teto sozinha, e é isso que a regra abaixo trata.

O teto é contado em pixels, e não em bytes, porque pixel é a unidade que o
código já manipula; a conversão é constante (3 bytes/pixel em RGB).

**A entrada recém-renderizada nunca é despejada**, mesmo que sozinha exceda o
teto: ela é a que está na tela. Nesse caso ela esvazia todas as outras e fica
sozinha no cache. A alternativa — recusar-se a guardar — faria o zoom máximo
re-renderizar a cada repintura, que é o oposto do objetivo.

## Testes

`tests/test_previa_documento.py` já cobre o contrato da renderização —
largura, proporção, limite de páginas, arquivo ruim, arquivo de 0 byte — e
pega regressão sozinho. Acrescentar:

- a largura pedida é a largura entregue, em vários níveis de zoom
- acima de 220% o pixmap **cresce de verdade** (a página tem mais pixels que
  antes), provando que não é upscale
- o teto de memória despeja de fato, e uma entrada acima do teto não trava o
  cache
- `extrair_texto_pdf` cai no pypdf quando o PyMuPDF falta
- pypdf e PyMuPDF concordam sobre as fixtures de NFS-e já existentes
  (`nfse_ff.txt`, `nfse_ff_retido.txt` e as DANFSe v2.0)

A parte de thread é interface e não entra em teste automatizado. É justamente
por isso que ela fica **fina de propósito**: nenhuma regra de negócio dentro
da thread, só rasterizar e devolver.

## Risco e o que observar depois

O risco concentra-se na troca do extrator, porque ela toca os três consumidores
de texto (identificação da aba 1, extração de NFS-e da aba 2, protocolos da
aba 3). A validação acima cobre os três com arquivos reais, mas ela é sobre
**os lotes que existem hoje**. Vale conferir o primeiro lote de NFS-e
processado depois da mudança contra a planilha do lote anterior: o formato
DANFSe já mudou uma vez sem aviso (v1.0 → v2.0, no meio de um mesmo dia), e o
extrator de texto é onde essa mudança apareceria primeiro.
