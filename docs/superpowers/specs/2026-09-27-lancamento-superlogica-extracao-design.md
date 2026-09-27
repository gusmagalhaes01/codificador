# Lançamento no Superlógica direto da extração (aba 2) — especificação

Data: 2026-09-27
Estado: aprovado, a implementar

## Objetivo

Fazer a aba 2 ("Extrair dados") gerar, junto da planilha de extração, o
**arquivo de importação de despesas do Superlógica** — o mesmo tipo de
arquivo que a aba 3 já gera para os Correios. Três fontes de documento:

- **Boletos do sindicato** (contribuição confederativa/assistencial)
- **DARF do DCTFWeb** (retenções; código de receita 5952 nas amostras)
- **Notas da F&F** (NFS-e: PCMSO, PGR, eSocial, exames...)

Um seletor na aba escolhe a **opção de lançamento** do lote — "FF - Exames
Medicos", "Sindicato - Contr. Assistencial", "DCTFWeb Complementar - Sem
assoc. de impostos" —, e a opção decide o que é fixo em todo lançamento
(fornecedor, favorecido, apropriação, tipo de documento, forma de pagamento).
O que varia por documento (condomínio, valor, vencimento, linha digitável...)
vem da extração.

**Uma opção por lote**, por decisão do usuário: cada pasta é de um tipo só.

## Por que é viável: a extração de hoje já faz quase tudo

Medido contra as amostras reais, antes de qualquer mudança:

| | Boletos do sindicato | DARF / DCTFWeb |
|---|---|---|
| Arquivos | 284 | 12 |
| Linha digitável lida (com DV) | **284** | **12** (arrecadação) |
| Valor | **284** | **12** |
| Vencimento | **284** | **0** |
| Condomínio pelo CNPJ | **284** | **12** |
| Divergência com o nome do arquivo | **0** | **0** |

O condomínio deduzido pelo CNPJ bateu com o código no nome do arquivo
(sindicato, `10007_CONFEDERATIVA_082026.pdf`) e com o CNPJ no nome do arquivo
(DARF, `GuiaPagamento_05860602000132_...pdf`) em **296 de 296** — dois sinais
independentes concordando. A F&F é a extração de NFS-e da v6.7.0, validada
em milhares de notas.

O único buraco é o **vencimento do DARF**: código de arrecadação não carrega
data. Ver seção 4.

## 1. As opções de lançamento: um arquivo de modelo por opção

Pasta **`modelos_superlogica/`** ao lado do programa (`pasta_base()`), com um
`.xlsx` por opção. **O nome do arquivo, sem extensão, é o texto do seletor.**
A linha 2 de cada arquivo é o molde, preenchida no Excel — o mesmo mecanismo
de `modelo_despesas.xlsx` na aba 3, que já funciona em produção.

Decisões:

- **Fornecedor e favorecido podem ir pelo ID do Superlógica** — o importador
  aceita. Tira o risco de nome digitado com uma letra trocada. A
  `conta_categoria` continua no formato `código nome`
  (`2.4.6 Correios - Postagem Simples`).
- **Vai haver várias opções da F&F** (PCMSO, PGR, eSocial, Exames...), cada
  uma com sua apropriação — exemplo do usuário: "FF - Exames Medicos" usa a
  mesma apropriação do PCMSO. Um arquivo por opção acomoda isso sem código.
- **O zip de instalação não traz modelos.** São do usuário; um pacote novo
  sobrescreveria as edições — o mesmo cuidado já registrado para o cadastro.
- **Pasta vazia ou inexistente**: o seletor mostra só "Nenhum" e um aviso
  dizendo onde colocar os modelos. A pasta é criada se não existir.

**Alternativa considerada e adiada:** guardar as opções no `config.json` e
editá-las numa tela de Configurações. O usuário prefere a longo prazo, mas
escolheu começar pelos arquivos. Se um dia migrar, os arquivos viram o ponto
de partida.

## 2. Na tela

Na aba 2, junto do botão de extrair:
**"Lançamento no Superlógica: [Nenhum ▾]"**.

- **"Nenhum"**: a aba funciona exatamente como hoje.
- **Uma opção escolhida**: ao terminar a extração, o programa pergunta a
  **chave** (sempre) e o **vencimento** (só se o lote tiver nota fiscal), e
  grava o arquivo de importação ao lado da planilha de extração.

As perguntas vêm **ao fim da extração**, e não no início como na aba 3. Lá o
vencimento é perguntado antes porque vai carimbado no PDF durante o
processamento; aqui nada é carimbado, e esperar o fim permite perguntar o
vencimento **só quando o lote de fato tem nota fiscal**. Cancelar uma
pergunta não perde a extração: a planilha de extração já foi gravada, só o
arquivo do Superlógica não é gerado.

## 3. O que cada documento preenche

| Coluna | Boleto do sindicato | DARF | Nota da F&F |
|---|---|---|---|
| `condomínio` (ID SL) | CNPJ do pagador | CNPJ | CNPJ do tomador |
| `valor` | código de barras | código de barras | valor do serviço |
| `vencimento` | código de barras | texto (seção 4) | o do lote |
| `linha_digitavel` | 47 dígitos | 48 dígitos | — |
| `numero_documento` | — | — | nº da NFS-e |
| `competencia` | — | — | competência da nota |

**Essas seis colunas são sempre escritas pelo programa, nunca herdadas do
molde** — "—" significa célula vazia, mesmo que o molde tenha algo ali. Foi
deixar o vencimento no modelo que fez o Superlógica gravar 01/01/1970 e
recusar lançamentos (v6.14.0); não se repete o caminho.

Todas as outras colunas vêm da linha 2 do modelo, sem nenhum tratamento:
fornecedor, favorecido, `conta_categoria`, `tipo_de_documento`,
`forma_de_pagamento` e **`complemento`**.

**O `complemento` não é escrito pelo programa**, por decisão do usuário: ele
o preenche à mão. Chegou a ser desenhado um complemento automático (o mês de
referência do boleto do sindicato, que vem impresso no documento), e foi
retirado — é código específico do layout de um boleto para economizar uma
digitação por lote.

**Cuidado que o usuário precisa saber:** se o complemento for preenchido na
**planilha gerada** aberta no Excel, é exatamente o gesto que já corrompeu a
coluna `vencimento` uma vez (v6.17.0). Preencher no **modelo** (linha 2) ou
no próprio Superlógica depois de importar não tem esse risco.

**`numero_documento` da F&F é o número da NFS-e** — é o que o usuário já
lança hoje, e é um dos critérios do grupo NOTA FISCAL do Paybox (ver "Por que
a NFS-e NÃO anexa sozinha no Paybox" no `CLAUDE.md`).

## 4. Vencimento do DARF: três ocorrências que têm de concordar

O código de arrecadação (48 dígitos, começando em 8) traz valor e DVs, mas
**não traz data**. A data está no texto nativo da guia, **três vezes**:

```
Data de Vencimento ... 18/09/2026
Pagar até: 18/09/2026
PA:08/2026 Vencimento:18/09/2026
```

Isso dá uma conferência de graça: **as ocorrências encontradas têm de ser
todas iguais**. Divergindo, ou nenhuma achada, o DARF fica de fora da
planilha do Superlógica com o motivo na observação — nunca se escolhe uma
das datas.

Ler data de texto nativo não contraria a regra "sem OCR na aba 2": a regra
é contra **OCR**, que troca dígito sem aviso. Texto nativo é o mesmo caminho
pelo qual a competência e a data de emissão das NFS-e já são lidas desde a
v6.7.0. E a tripla repetição dá ao vencimento do DARF uma proteção que a
data da NFS-e nem tem.

## 5. O que fica de fora — e é listado

Não entra no arquivo do Superlógica, e sai com o motivo na coluna
Observação da planilha de extração:

| Caso | Por quê |
|---|---|
| **Nota da F&F com retenção** | regra do usuário: é lançada à mão |
| Condomínio não identificado | não há `condomínio` para lançar |
| Condomínio **sem ID SL** | o importador não tem como achar o condomínio (10 dos 772 hoje) |
| DARF sem vencimento, ou com datas que não batem | seção 4 |
| Boleto sem vencimento no código (fator `0000` ou `9999`) | ver abaixo |
| Documento sem código legível | não há valor conferido |

**Boleto sem vencimento no código fica de fora, e não pega a data impressa
na folha.** O fator `9999` não é hipótese: o boleto real do Itaú de
referência (v6.18.2) tem 15/09/2026 impresso e `9999` na barra. Seria
tentador ler a data da folha como no DARF, mas o boleto não tem a tripla
repetição que protege o DARF — seria uma data solta, sem conferência. Nos
284 do sindicato não aconteceu nenhuma vez; é caso de exceção, e exceção vai
para o manual.

**Retenção é detectada por dois sinais independentes**, já lidos pela
extração de NFS-e: `valor do serviço ≠ valor líquido`, **ou** "Contrib.
sociais retidas" / "Prev. retida" preenchidos. Qualquer um dos dois manda a
nota para o manual. Nas notas de fevereiro/2026, 228 de 594 tinham retenção
— o manual não é caso raro.

Ao terminar, uma mensagem com a contagem: *"284 lançados, 3 ficaram de fora
— veja a Observação na planilha de extração."*

**Isto é deliberadamente diferente da aba 3**, que trava a geração inteira
por uma pendência. Lá todo protocolo tem de ser cobrado; aqui parte dos
documentos é manual **por regra do usuário**, e travar o lote por causa
deles não faria sentido.

## 6. Trava contra pasta misturada

Como a opção vale para o lote inteiro, uma NFS-e perdida numa pasta de
boletos do sindicato seria lançada com o fornecedor do sindicato.

Os documentos **reconhecidos** do lote têm de ser de um tipo só: nota
fiscal, boleto bancário ou arrecadação. Havendo mais de um tipo, o arquivo
do Superlógica **não é gerado** e a mensagem diz quais arquivos destoam. A
planilha de extração sai normalmente.

**Documento não reconhecido não conta como tipo.** O "Detalhamento do
Faturamento" vem misturado nas pastas da F&F (v6.7.0) e não é NFS-e nem
boleto; se contasse, toda pasta da F&F seria recusada.

## 7. O que muda no código

- `gerar_planilha_despesas` (`logica.py`) passa a aceitar, por lançamento,
  um dict `{coluna: valor}` com as colunas por documento da seção 3. **O
  formato atual `(id_sl, valor)` continua aceito**, para a aba 3 não mudar.
- Leitura do vencimento do DARF a partir do texto (seção 4), em `logica.py`.
- Montagem dos lançamentos a partir das linhas extraídas, com as exclusões
  da seção 5 e a trava da seção 6 — em `logica.py`, testável sem interface.
- Seletor, listagem da pasta de modelos e perguntas de fim de extração —
  interface.

**O que não muda:** a aba 3 e seu `modelo_despesas.xlsx`; a planilha de
extração da aba 2; a aba 2 com o seletor em "Nenhum".

## Testes

Fixtures de texto inline ou em `tests/dados/`, como as de NFS-e e boleto —
nunca PDF no repositório. Cobrir:

- vencimento do DARF: três ocorrências iguais → a data; divergentes → nada;
  nenhuma → nada
- nota com retenção fica de fora, pelos **dois** sinais separadamente
- condomínio sem ID SL fica de fora
- pasta misturada recusa; "Detalhamento do Faturamento" não conta como tipo
- as seis colunas por documento são escritas mesmo com o molde preenchido
  nelas, e o molde é copiado nas demais (inclusive `complemento`)
- `gerar_planilha_despesas` continua aceitando `(id_sl, valor)` — os testes
  existentes da aba 3 têm de passar sem alteração

## Validação contra os dados reais (obrigatória antes de liberar)

- **Sindicato**: os 284 de `C:\Users\Desktop\Music\sindicato\boletos` —
  lançamento para cada um, condomínio batendo com o código do nome do
  arquivo.
- **DARF**: os 12 de `C:\Users\Desktop\Music\inss` — vencimento lido nos 12,
  condomínio batendo com o CNPJ do nome do arquivo.
- **F&F**: um lote real de notas, conferindo que as com retenção ficam de
  fora e as sem retenção entram com o número da NFS-e.

## Riscos e o que não se sabe

- **O importador pode recusar o código de arrecadação** de 48 dígitos na
  coluna `linha_digitavel`. O `CLAUDE.md` registra que ela recusa o que não
  for código de boleto ("Código de barras inválido"). Teste obrigatório:
  importar **um** DARF de verdade antes de liberar. Se recusar, o DARF sai
  com a coluna vazia.
- **Formato da `competencia`** que o importador espera. O código já trata a
  coluna como data (`COLUNAS_DATA_DESPESAS`), mas a da NFS-e é uma data
  cheia (dia inclusive) e competência costuma ser mês. Conferir no mesmo
  teste de importação.
- **O Paybox pode passar a anexar os boletos sozinho** com a linha digitável
  no lançamento — plausível, mas **não testado**, e não é promessa desta
  entrega.
- **Base desatualizada:** o `main` local estava três semanas sem trazer o
  remoto quando este spec foi escrito. Antes de implementar, confirmar com o
  usuário e trazer o GitHub.
