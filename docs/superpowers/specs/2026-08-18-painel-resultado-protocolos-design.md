# Painel de resultado da aba 3, com valor informado à mão — especificação

Data: 2026-08-18
Estado: aprovado, a implementar

## Objetivo

Dar à aba 3 (Protocolos dos Correios) o mesmo tipo de tela de encerramento que
a aba 1 já tem: um painel que lista o que ficou pendente e deixa resolver ali
mesmo. A resolução aqui é **informar o valor em reais à mão** — o protocolo é
carimbado com esse valor e entra na cobrança.

Hoje, quando a contagem não confere, o arquivo sai do lote com uma linha de
observação na planilha e nada mais. Quem opera precisa abrir a planilha,
descobrir quais faltaram, calcular no papel e lançar por fora.

## O caso que motivou

Protocolo real do `11049 APART HOTEL`, 3 páginas: no pé da tabela está
impresso `Listando 76 unidades`, com o **76 riscado a caneta vermelha e 78
escrito embaixo**. O programa não conseguiu ler o dígito sob o risco,
`total_impresso` voltou vazio e o arquivo virou pendência — comportamento
correto, porque o número que ele trata como fonte da verdade estava
visivelmente adulterado.

Esse caso não tem solução automática: a correção está fora do que a máquina
lê, e a contagem de "Correio" concordava com o número impresso (76), não com
o corrigido (78). Só uma pessoa resolve. O painel é o lugar de resolver.

**Por que valor em reais e não quantidade de unidades:** nas folhas de
referência o que a pessoa escreve à mão é o valor (7,70 / 57,75 / 11,55 /
46,20). Digitar o mesmo número que já se escreveria à caneta é o caminho mais
curto e o menos sujeito a erro de transcrição. A consequência aceita é que a
coluna Unidades fica vazia nessas linhas e o carimbo não mostra a conta.

## Arquitetura: helpers compartilhados

Três helpers extraídos do `mostrar_resultado` da aba 1 (`identificacao_por_cnpj_6_0.py`),
passando a servir as duas abas:

| Helper | Responsabilidade |
|---|---|
| `_montar_painel_resultado(titulo, geometria)` | Cria ou reaproveita o `CTkToplevel`, aplica tema, título, tamanho e `transient`. |
| `_montar_faixa_cartoes(parent, cartoes)` | Recebe `[(valor, rotulo, destaque), ...]` e monta a linha de caixinhas. `destaque=True` sai em cobalto. |
| `_montar_tabela_resultado(parent, colunas, larguras)` | Devolve o `ttk.Treeview` já estilizado. |

A aba 1 passa a chamá-los em vez de montar tudo inline. **É código em produção:**
a refatoração é estrutural, sem mudança de comportamento, e isso precisa ser
verificado comparando a árvore de widgets gerada antes e depois — não só lendo
o diff.

## O painel da aba 3

`mostrar_resultado_protocolos(resultado)`.

**Faixa de cartões:** Protocolos · **Pendentes** (cobalto) · Total em R$. Sem
o cartão de tempo que a aba 1 tem — aqui o número que importa é o dinheiro.

**Tabela de pendentes, primeiro** (Arquivo | Condomínio | Motivo), com duas
ações por linha, acessíveis por duplo clique ou botão:

- **Abrir PDF** (`os.startfile`), botão neutro contornado;
- **Informar valor**, botão cobalto — é o que resolve a pendência, e a regra de
  cor do projeto reserva o cobalto exatamente para isso.

**Tabela de processados, depois** (Arquivo | Condomínio | Unidades | Valor).

**Terceira lista, curta: arquivos que não são protocolo**, sem nenhuma ação.
Não têm código nem contagem para informar, e oferecer "informar valor" ali
seria convidar a inventar cobrança sobre um documento alheio ao lote.

## Informar o valor

Janelinha no mesmo molde da que pede a tarifa: nome do arquivo e condomínio no
topo, um campo, aviso inline quando a entrada é inválida.

Aceita `300,30` e `300.30`. Recusa zero, negativo e **mais de duas casas
decimais** — a mesma validação da tarifa, pelo mesmo motivo: um valor com três
casas gera carimbo e planilha que não fecham quando alguém confere no papel.

Ao confirmar, em ordem:

1. **Carimba o PDF** — código do condomínio na lateral rotacionada, como
   sempre, e `R$ 300,30` no canto superior direito. Só o valor, sem a conta,
   porque não há conta a mostrar. Código não cadastrado carimba só o valor,
   como já acontece no fluxo normal.
2. **Move a linha** de pendentes para processados, onde ela aparece com a
   coluna Unidades vazia e o Valor preenchido. Na **planilha** essa linha leva
   ainda a observação `Valor informado manualmente` — a tabela do painel não
   tem coluna de observação, e é na planilha que fica o rastro de que o número
   veio de uma pessoa, já que o carimbo não diz isso.
3. **Atualiza os cartões** — Pendentes cai um, Total sobe.
4. **Regrava a planilha.**

**Se a regravação falhar** (tipicamente porque a planilha está aberta no
Excel), o valor informado **não pode se perder**: fica no estado do painel,
aparece o aviso que já existe hoje ("Se ela estiver aberta no Excel, feche e
tente de novo"), e a linha continua marcada como resolvida para ser regravada
na próxima tentativa. Regravar a cada correção foi decisão explícita do
usuário, com esse risco conhecido.

## Contexto guardado

A aba 3 passa a guardar `_ctx_protocolos` ao terminar um processamento, no
mesmo espírito do `_ctx_processamento` da aba 1: pasta de saída, dict de
aparência do carimbo, caminho da planilha e a lista de linhas. É o que permite
carimbar e regravar depois, fora do laço de processamento.

## Mudança em `logica.py`

`linha_planilha_protocolo` ganha um parâmetro opcional `valor_manual=None`.
Quando informado, preenche a coluna Valor e deixa **Unidades e Tarifa vazias
(`None`), nunca 0** — mesma regra já aplicada às pendências, porque 0
significaria "entregou zero unidades". A observação padrão passa a ser
`Valor informado manualmente`.

`salvar_planilha_protocolo` não muda: a linha de TOTAL já soma a coluna Valor,
então valores informados à mão entram no total sem tratamento especial.

## Testes

Na lógica, que é onde há cobertura automatizada:

- `linha_planilha_protocolo` com `valor_manual`: Valor preenchido, Unidades e
  Tarifa vazias, observação correta;
- a linha de TOTAL somando valores manuais junto com os calculados;
- a leitura do valor digitado: `300,30` e `300.30` produzem o mesmo `Decimal`;
  `300,305`, `0` e `-5` são recusados.

A refatoração dos helpers é coberta por regressão comparando a árvore de
widgets do painel da aba 1 antes e depois, com o resultado registrado.

O painel em si é verificação visual pelo usuário, como foi feito com o
carimbo — não há teste automatizado de interface neste projeto.

## Fora de escopo

- Tentar ler o valor manuscrito do papel. Nenhum dos três motores de OCR
  testados (winocr, Tesseract, RapidOCR) lê caligrafia de forma confiável — o
  RapidOCR leu `1.70` onde o papel dizia 7,70.
- Informar quantidade de unidades em vez de valor. Decidido pelo usuário.
- Mexer no painel da aba 1 além da extração dos três helpers.
- Melhorar a contagem por linha, que nos protocolos grandes desaba (1 de 31 num
  caso real). É problema separado, registrado à parte.
