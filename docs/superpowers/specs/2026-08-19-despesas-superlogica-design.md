# Planilha de Despesas do Superlógica a partir dos protocolos — especificação

Data: 2026-08-19
Estado: aprovado, a implementar

## Objetivo

Transformar o resultado de um lote de protocolos dos Correios na planilha de
importação de despesas do Superlógica, para que os lançamentos deixem de ser
digitados um a um.

Cada protocolo cobrável vira uma linha, com o campo `condomínio` preenchido
com o **ID SL** (o código do condomínio no Superlógica, guardado na 4ª coluna
do cadastro) e o campo `valor` com o valor apurado do protocolo. Os demais
campos se repetem em todas as linhas.

## O modelo vem do usuário, não do código

O arquivo `Despesas.xlsx` é o modelo de importação **baixado do próprio
Superlógica** — 32 colunas, usado para vários tipos de lançamento. Numa cópia
dele o usuário preencheu uma linha de exemplo com o que se repete no caso dos
Correios:

| Campo | Valor |
|---|---|
| `fornecedor` | DINAMICA SERVICOS POSTAIS E TELEMATICOS LTDA |
| `favorecido` | IMODATA ADMINISTRACAO COMPRA E VENDA DE IMOVEIS LTDA |
| `conta_categoria` | 2.4.6 Correios - Postagem Simples |
| `tipo_de_documento` | Outros |
| `forma_de_pagamento` | Trans. bancária |
| `chave` | 46 |

`vencimento`, `competencia`, `condomínio` e `valor` estão em branco.

**O programa lê esse arquivo como modelo a cada geração**, em vez de ter as
colunas e os valores no código. Motivo: o layout pertence ao Superlógica e
pode mudar, e a categoria ou o fornecedor podem mudar por decisão comercial.
Nos dois casos basta o usuário atualizar o próprio arquivo — nada de mexer no
programa e gerar executável novo.

**O arquivo é copiado, não reconstruído.** Abrir e preencher preserva o que
não se vê — formatos de célula, validações, colunas ocultas — que o importador
do Superlógica pode exigir. Reconstruir do zero perderia isso sem aviso.

## Como cada linha é montada

- As colunas `condomínio` e `valor` são localizadas **pelo nome no cabeçalho**,
  normalizando acento e caixa, nunca pela posição. Se o Superlógica reordenar
  ou renomear as colunas em volta, a geração continua correta.
- A **linha 2 é o molde**: seus valores são replicados em toda linha gerada.
  Ela é **substituída** pela primeira linha real — não pode sobrar no arquivo
  final como uma linha meio vazia.
- A partir da linha 2, uma linha por protocolo cobrável, na mesma ordem da
  planilha de protocolos.
- `condomínio` recebe o ID SL; `valor` recebe o valor apurado, gravado como
  **número**, não texto, para o importador não tropeçar.

**Mesmo condomínio em mais de um protocolo gera mais de uma linha.** No lote
multipágina de referência, dois protocolos são do 11161 MARILIA (36 e 31
unidades): saem como `496 / 138,60` e `496 / 119,35`. Cada lançamento continua
rastreável até o papel que o originou; somar os dois quebraria essa
correspondência.

## A trava

O botão **recusa gerar** enquanto houver protocolo cobrável incompleto. Dois
casos, com ações diferentes:

| Situação | O que falta | Onde se resolve |
|---|---|---|
| Pendência sem valor | a contagem não fechou e ninguém informou o valor | no painel, botão "Informar valor" |
| Condomínio sem ID SL | o cadastro não tem o código do Superlógica | na aba 4, Cadastro de Condomínios |

A mensagem lista **nome do arquivo e motivo**, agrupados por caso — sem isso,
"não gera nada" vira adivinhação.

Arquivos que **não são protocolo** não contam para a trava: nunca deveriam
virar despesa.

**Custo conhecido desta escolha, aceito pelo usuário:** um único condomínio sem
ID SL trava o lote inteiro. Hoje 10 dos 764 condomínios estão sem o campo, então
isso vai acontecer. O conserto é rápido (preencher na aba de Cadastro), mas
interrompe o trabalho.

## Onde mora

Botão **"Gerar planilha do Superlógica"** no painel de resultado da aba 3,
neutro (contornado, fundo transparente) — o cobalto continua reservado ao que
resolve pendência.

Ao clicar, pede o arquivo modelo e o destino, sugerindo salvar ao lado da
planilha de protocolos. Não reabre a planilha de protocolos: os dados do lote,
incluindo as pendências resolvidas à mão, já estão em `_ctx_protocolos`.

## Código

Em `logica.py`, sem dependência de interface:

- `gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos)` —
  `lancamentos` é uma lista de `(id_sl, valor)` na ordem de saída. Abre o
  modelo, localiza as colunas pelo cabeçalho, replica o molde, grava.
- `lancamentos_de_despesa(linhas, cadastro)` — monta a lista a partir das
  linhas do lote e devolve também o que travou, separado por motivo. Fica fora
  da tela para ser testável, no mesmo espírito de `resolver_protocolo_manual`.

A docstring de `carregar_cadastro` precisa mudar: hoje afirma que o ID SL é
"guardado só como referência: nada da identificação nem dos carimbos usa esse
campo". Deixou de ser verdade, e docstring que mente é pior que nenhuma.

## Testes

Sobre um modelo sintético montado no próprio teste:

- uma linha por lançamento, na ordem dada;
- os seis campos do molde repetidos em **todas** as linhas;
- a linha-molde não sobra em branco no arquivo final;
- `condomínio` recebe o ID SL e `valor` é gravado como número, não texto;
- as colunas são encontradas com acento e caixa diferentes no cabeçalho;
- modelo sem a coluna `condomínio` ou sem `valor` → erro claro, nunca um
  arquivo silenciosamente errado;
- lista de lançamentos vazia não gera arquivo pela metade;
- `lancamentos_de_despesa`: protocolo sem valor e condomínio sem ID SL entram
  na lista de travados, com o motivo certo; arquivo que não é protocolo é
  ignorado sem travar; mesmo condomínio duas vezes gera dois lançamentos.

## Fora de escopo

- Preencher `vencimento` e `competência`: estão em branco no modelo do usuário
  e continuam em branco.
- Somar protocolos do mesmo condomínio.
- Importar direto no Superlógica por API: aqui só se gera o arquivo.
- Lembrar o caminho do modelo entre execuções.
