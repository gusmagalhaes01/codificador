# Associação automática das notas da F&F no Paybox — Design

**Data:** 2026-09-27
**Status:** aprovado para plano

## Problema

As NFS-e da F&F (cerca de **1.800 por mês**) não se associam sozinhas às
despesas no Paybox, e cada uma é vinculada à mão na fila. O motivo é estrutural
e já está documentado no CLAUDE.md ("Por que a NFS-e NÃO anexa sozinha no
Paybox"). A nota é classificada como `nfse` e lida pelo esquema do QR Code. O
assistente de anexos exige o campo Vencimento do arquivo preenchido, e esse
campo nunca é preenchido.

## Caminhos descartados (medidos, não presumidos)

- **Carimbo na nota:** a folha de uma `nfse` não é lida. Testado em 3 notas.
- **Coluna `etiqueta_paybox` da planilha de importação**, com o link
  `sldocs.com.br` de cada nota. Foi testado em 2026-09-27 com a
  `PGR 10141 Solar Rego Lopes` (NFS-e 11882). A despesa entrou certa e com o
  link gravado, mas **a imagem continuou na fila**. A coluna só guarda o link,
  não associa.
- **API oficial (App Token):** seria o caminho mais limpo, mas depende de um
  administrador criar o token. Não bloqueia este trabalho. Se o token vier, a
  associação pode migrar para a API.

O mesmo teste confirmou, de brinde, que o importador **aceita a competência no
formato que o Codificador grava**. Era um pendente da v6.21.0 para as notas da
F&F.

## Solução

Um **favorito do navegador (bookmarklet)** que, clicado com o Superlógica
aberto, percorre a fila do Paybox. Ele associa cada NFS-e da F&F à sua despesa
fazendo as **mesmas chamadas que a tela do Paybox faz** quando a pessoa vincula
à mão, capturadas por HAR em 2026-09-27.

**Fica fora do Codificador, por decisão do usuário.** Mora em pasta separada no
repositório, sem dependência nova, sem entrar no executável e sem mexer em
`logica.py` nem na interface.

### Por que bookmarklet

- Roda **dentro da página já logada**. O navegador usa a sessão sozinho, e o
  código nunca lê, guarda nem transmite senha, cookie ou token.
- Nada para instalar além do favorito, que é criado uma vez.
- Não depende da extração da aba 2: a fila já traz número da nota, CNPJ e
  valor, e a despesa já existe desde a importação.

**Custo assumido:** são chamadas internas, não a API documentada. O Superlógica
pode mudá-las sem aviso, e aí o arquivo precisa de ajuste. As proteções abaixo
existem para que uma mudança dessas **pare** o processo em vez de associar
errado.

## Arquivos

```
ferramentas/paybox_ff/
  associar_ff.js   — o código (lógica pura + chamadas + caixinha)
  instalar.html    — link para arrastar à barra de favoritos, com instruções
  testes.html      — página que roda os testes (abrir no navegador)
  testes.js        — os casos de teste
  rodar_testes.py  — roda testes.html no Edge sem janela; sai com erro se falhar
  LEIAME.md        — como usar, o que cada relatório significa, como atualizar
```

As fixtures dos testes são **extraídas e anonimizadas** dos HARs: sem o campo
`session`, sem cookies, sem cabeçalhos, sem nome de usuário. Nenhum HAR entra no
repositório.

## As chamadas (capturadas em 2026-09-27)

Todas vão para `https://<licença>.superlogica.net/condor/atual/...`, a mesma
origem da página. A licença é lida de `location.host`, nunca fixada no código.

As leituras são `POST` com corpo `application/x-www-form-urlencoded` e um único
campo `json`, contendo `{"params":[{...}],"url":"<url completa>"}`. É o formato
que a própria tela usa.

| Passo | Chamada | O que tira dela |
|---|---|---|
| Listar a fila | `ocr/getenvelopes?idCondominio=-2&pagina=N` | `data.envelopes[]` (cada um com `envelope_id`, `metadata.document-type`, `metadata.document-recipient`, `metadata.invoice-number`, `metadata.invoice-value`, `metadata.net-value`, `metadata.filename`) e `data.totaldepaginas` |
| Dados do envelope | `ocr/getenvelope?idEnvelope=<id>` | `data.invoices[]` (`type`, `number`, `amount`, `amount_before_taxes`), `data.customer` (`id_condominio_cond`, `st_fantasia_cond`), `data.issuer` (`document_number`, `id_contato_con` e a classificação da Reinf). É a **fonte de verdade** da nota. A lista da fila só serve para achar os envelopes |
| Buscar despesas | `despesas/index` com params `comStatus:"todas"`, `idCondominio`, `FAVORECIDOS:[<id_contato_con do issuer>]`, `pesquisa:<valor "84.65">`, `tipoFiltroData:"periodo"`, `dtInicio`/`dtFim` (MM/DD/AAAA), `itensPorPagina`, `pagina` | por despesa: `id_despesa_des`, `id_parcela_pdes`, `st_documento_des`, `st_cpf_con`, `id_contato_con`, `vl_valor_pdes`, `dt_despesa_des`, `fl_modelotrabalho_des`, `id_tipo_doc`, `id_forma_pag`, `st_complemento_pdes`, `arquivos`, `st_fantasia_cond`, `st_nome_con` |
| Associar | `ocr/vincularenvelopeadespesa`, corpo em campos de formulário soltos | resposta `status: "200"` |

A resposta de `getenvelopes` numa listagem sem filtro tem a forma do
`index.json` que o usuário exportou (`data` como lista). Numa listagem filtrada
por `pesquisa`, a forma é `data.envelopes`. **O código aceita as duas** e,
diante de qualquer outra forma, para (ver Proteções).

### Campos do `vincularenvelopeadespesa`, e de onde cada um vem

| Campo | Origem |
|---|---|
| `ID_CONDOMINIO_COND` | envelope (`id_condominio_cond`) |
| `ID_ENVELOPE_ENV`, `ID_ENVELOPEMENSAGEIRO_PDES` | envelope (`envelope_id`) |
| `ID_DESPESA_DES`, `ID_PARCELA_PDES` | despesa |
| `ID_CONTATO_CON` | despesa (`id_contato_con`) |
| `VL_VALOR_PDES`, `ST_VALORPRIMEIRAPARCELA`, `VL_DOCUMENTO_DES` | despesa (`vl_valor_pdes`), ponto decimal |
| `VL_VALOR_ENV` | o mesmo valor, **vírgula decimal** ("84,65"), como a tela manda |
| `ST_DOCUMENTO_DES` | despesa |
| `DT_DESPESA_DES` | despesa, só a data (MM/DD/AAAA) |
| `FL_MODELOTRABALHO_DES`, `ID_TIPO_DOC`, `ID_FORMA_PAG` | despesa |
| `ST_COMPLEMENTO_APRO` | despesa (`st_complemento_pdes`), **nunca** a descrição da nota |
| `ST_CPF_CNPJ_FORNECEDOR_ENV` | CNPJ da F&F |
| `ST_FANTASIA_COND` | envelope/despesa |
| `ST_NOME_CON` | despesa |
| `ST_CLASSIFICACAO_TRIBUTARIA`, `ID_CLASSIFICACAOTRIBUTARIA_DES`, `ST_CLASSIFICACAO_SERVICO_PRESTADO`, `ID_CLASSSERVICOPRESTADO_DES` | envelope, bloco `issuer` do `getenvelope` (`st_classificacao_tributaria`, `id_classificacao_tributaria`, `st_classificacao_servico_prestado`, `id_classificacao_servico_prestado`) |

**A classificação da Reinf vem do cadastro do fornecedor, exatamente como a
tela faz.** O HAR mostrou que os valores "99 Pessoas Jurídicas em Geral" /
"100000006 Preparação de dados para processamento" não são inventados pela
tela: saem do bloco `issuer` do `getenvelope`, que é o cadastro da F&F no
Superlógica. Mandar o mesmo que a associação manual manda produz o mesmo
resultado que ela. A despesa só tem os textos (`st_classificacao_*`), sem os
ids, então não daria para copiar dela. Se a classificação estiver errada, o
conserto é no cadastro do fornecedor, não aqui.
| `VALIDAR_CPF_CNPJ`=`1`, `FL_MARCAR_PARA_LANCAMENTO`=`0`, `despesa`=`on`, `salvar`=`Vincular` | fixos, como a tela manda |

**Por que copiar da despesa e não repetir o que a tela manda:** na captura, a
tela mandou `ST_COMPLEMENTO_APRO` com a descrição da nota ("ASSESSORIA OU
CONSULTORIA... FATURA NO. 358995..."), `ID_FORMA_PAG=0` (a despesa estava com
`8`) e uma classificação da EFD-Reinf ("100000006 Preparação de dados para
processamento") que a despesa não tinha. Conferido na despesa 56112, nem o
complemento nem a forma de pagamento mudaram. O Superlógica fica com o que está
no lançamento. Mesmo assim, o código não depende desse comportamento: se
depender e ele mudar, 1.800 despesas por mês seriam reescritas.

## Regra de casamento

Uma nota da fila entra no processo se `metadata.document-type == "nfse"` **e**
`metadata.document-recipient == "13736666000154"` (F&F). Todo o resto da fila é
ignorado e não aparece no relatório como erro.

Para cada nota:

1. Lê o envelope (`getenvelope`), de onde vêm a nota, o condomínio e o
   fornecedor. O envelope precisa ter **exatamente uma** nota, do tipo `nfse` e
   emitida pela F&F. Senão a nota fica de fora com o motivo.
2. Busca as despesas desse condomínio e desse fornecedor no período informado,
   com `pesquisa` pelo valor da nota. A primeira tentativa usa o `amount`. Sem
   candidata, tenta de novo com o `amount_before_taxes`, se for diferente. Lê
   todas as páginas da busca.
3. **Candidatas** são as despesas com `st_cpf_con == "13736666000154"` **e**
   `st_documento_des` igual ao número da nota, comparados sem zeros à
   esquerda.
4. O resultado:

| Situação | Ação | Motivo no relatório |
|---|---|---|
| exatamente 1 candidata, sem arquivo, valor = `amount` ou `amount_before_taxes` (tolerância R$ 0,005) | **associa** | — |
| 0 candidatas | fica | "Despesa não encontrada no período" |
| 2 ou mais | fica | "Mais de uma despesa com o documento N" |
| 1, mas já tem arquivo | fica | "Despesa já tem anexo" |
| 1, valor diferente | fica | "Valor da despesa (X) difere da nota (Y)" |
| envelope com 0 ou 2+ notas | fica | "Envelope com N notas" |
| nota do envelope não é NFS-e da F&F | fica | "Não é NFS-e da F&F" |
| envelope sem condomínio | fica | "Paybox não identificou o condomínio" |

O **número do documento é a chave**, e o valor é conferência. Por isso a
pesquisa é pelo valor (o filtro que a tela usa e que foi medido), e o número é
comparado no próprio código.

## Interface (a caixinha)

- Painel fixo no canto da página, no estilo Swiss do Codificador: cantos retos,
  cinza neutro e cobalto só no botão principal.
- **Campos:**
  - período de vencimento, De/Até, em DD/MM/AAAA;
  - "Associar no máximo [N]" (vazio = sem limite);
  - escolha **Simular** / **Associar**;
  - botão **Iniciar** (cobalto) e **Parar**.
- **Andamento:** "Lendo a fila… página 3 de 90", depois "Nota 312 de 1.800 —
  associadas 290 · fora 22 — faltam ~21 min".
- **No fim:** resumo por motivo e botão **Baixar relatório (CSV)**, com as
  colunas arquivo, nota, condomínio, valor, resultado, despesa e motivo.
  - Na simulação, o resultado diz "associaria à despesa X".
  - O CSV usa `;` e BOM UTF-8, para abrir certo no Excel em português.
- Clicar no favorito de novo **não abre um segundo painel**. Reaproveita o
  aberto.

## Proteções

- **Ritmo de pessoa:** uma chamada por vez, com pausa de ~1 s entre
  associações. Nada em paralelo.
- **Parada no inesperado:** para tudo e mostra onde parou quando:
  - a resposta HTTP não for 200;
  - o JSON vier com `status` diferente de `"200"`;
  - a resposta não tiver a forma esperada (campo que falta, `data` de tipo
    diferente).

  Não segue "pulando" o que não reconhece: se o Superlógica mudar algo, o certo é
  parar.
- **Sessão expirada:** uma resposta HTML no lugar de JSON (página de login) é
  tratada como parada, com a mensagem "Sessão expirou — entre de novo e rode
  outra vez".
- **Rodar de novo é seguro:** o que já foi associado sai da fila e não é
  reprocessado. Despesa que já tem arquivo também não é tocada.
- **Página errada:** clicado fora de `*.superlogica.net`, o favorito só avisa e
  não faz nada.
- **Nada sai da página:** nenhuma chamada a outro domínio e nada gravado no
  navegador além do painel. O relatório é gerado localmente, no clique de
  baixar.

## Pendente de conferência (primeiro uso real)

1. **Classificação da EFD-Reinf:** confirmar, na despesa associada na etapa 2
   abaixo, que ela ficou igual à de uma associação manual. Ela vem do cadastro
   do fornecedor, como na tela.
2. **Cabeçalhos exigidos:** se as chamadas feitas por `fetch` a partir do
   favorito precisam de algum cabeçalho que a tela adiciona, como
   `X-Requested-With`. Se precisarem, a primeira chamada falha e o processo para,
   sem risco.
3. **Formato da listagem sem filtro** (`data` lista ou `data.envelopes`): o
   código aceita os dois, e a simulação mostra qual veio.

**Ordem do primeiro uso:**

1. Simulação da fila inteira. O usuário confere o relatório.
2. Associação com limite **1**. O usuário abre a despesa e confere o anexo, a
   forma de pagamento, o complemento e a classificação da Reinf.
3. Associação com limite de ~20, conferida por amostragem.
4. Lote inteiro.

## Testes

- `testes.html` roda no navegador e testa só a **lógica pura**:
  - o filtro das notas da F&F;
  - a escolha da candidata, cobrindo todas as linhas da tabela de casamento;
  - a normalização do número do documento;
  - os dois formatos de valor;
  - a montagem do corpo do `vincular`, com os campos copiados da despesa e nunca
    a descrição da nota;
  - a conversão de datas DD/MM/AAAA ↔ MM/DD/AAAA;
  - o reconhecimento das duas formas de `getenvelopes`;
  - a detecção de resposta inesperada ou de sessão expirada;
  - a geração do CSV.
- As chamadas de rede ficam atrás de uma função única de transporte, que nos
  testes é trocada por respostas gravadas (fixtures anonimizadas). Assim o fluxo
  inteiro de simular e associar roda sem rede, e dá para conferir que a
  simulação **nunca** chama o `vincular`.
- Verificação: `python ferramentas/paybox_ff/rodar_testes.py` abre o
  `testes.html` no Edge sem janela (`--headless --dump-dom`) e sai com erro se
  algum teste falhar. Não há Node na máquina, e o Edge vem com o Windows. O mesmo
  `testes.html` também pode ser aberto à mão no navegador. A
  suíte Python do Codificador não é afetada.

## Fora do escopo

- Outros fornecedores e outros tipos de documento. O filtro da F&F é
  deliberado, mas o código deixa o CNPJ numa constante no topo, e generalizar é
  trabalho futuro.
- Integração com o Codificador ou com o executável.
- Uso da API oficial (depende do App Token).
