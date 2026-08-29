# Planilha embutida com prévia do documento (aba 3) — especificação

Data: 2026-08-29
Estado: aprovado, a implementar

## Objetivo

Substituir as duas tabelas do painel de resultado dos protocolos (PENDENTES e
CALCULADOS) por **uma grade única e editável que É a planilha**, com a
**prévia do PDF ao lado**. Elimina duas idas e voltas do fluxo atual: abrir o
PDF num visualizador externo para descobrir de qual condomínio é o documento,
e abrir a planilha no Excel para conferir o lote.

## Por que uma grade só, e não uma terceira seção

A planilha já contém tudo que as duas tabelas mostram — `Arquivo`,
`Condomínio`, `Código`, `Unidades`, `Tarifa`, `Valor`, `Observação`. Manter as
duas tabelas E acrescentar a grade faria as mesmas linhas aparecerem duas
vezes na mesma tela.

A distinção pendente/calculado não some: vira **estado da linha** (um filtro e
um destaque visual), não duas tabelas separadas. Continua sendo verdade que
"tudo que precisa de ação fica num lugar só" — só que agora esse lugar é a
própria planilha.

## A regra que não pode ser quebrada: papel e planilha não podem divergir

O valor não vive só na planilha. Ele está **carimbado no PDF**: o bloco do
Paybox estampa `VALOR` e `VENCIMENTO` na folha, e o Superlógica lê o papel por
OCR para casar com o lançamento. O código e o nome do condomínio idem, no
carimbo lateral.

Editar a planilha sem recarimbar criaria divergência **silenciosa** — o Paybox
simplesmente não acha o lançamento, o arquivo cai na fila manual e ninguém
recebe erro (mesmo modo de falha já registrado no CLAUDE.md sobre o
vencimento).

Por isso: **toda edição que afeta o papel dispara o recarimbo do PDF**, pelo
mesmo caminho que "Informar valor" e "Escolher condomínio" já usam hoje. A
edição não é um atalho que fura a coerência — é um caminho mais direto para as
mesmas operações.

| Coluna | Editável | Efeito |
|---|---|---|
| `Arquivo` | não | é a identidade da linha |
| `Condomínio` | não (derivada) | muda junto com `Código` |
| `Código` | sim | recarimba (carimbo lateral) + regrava a planilha |
| `Unidades` | sim | recalcula o valor pela tarifa, recarimba, regrava |
| `Tarifa` | não | é do lote inteiro, não da linha |
| `Valor` | sim | recarimba (bloco do Paybox) + regrava |
| `Observação` | sim | só planilha; não afeta o papel |

Editar `Código` para um valor fora do cadastro é recusado na hora, com o
motivo — em vez de gravar e quebrar depois na geração das despesas.

Se o recarimbo falhar, a edição **não é aplicada**: mesma regra que
`resolver_protocolo_manual` já segue (não faz sentido marcar como resolvido um
arquivo que não saiu carimbado).

## Prévia do documento

- Painel à direita da grade, mostrando a **primeira página** do PDF da linha
  selecionada, em tamanho legível.
- Vale para **qualquer** linha, não só pendentes: nos calculados serve de
  conferência — dá pra bater o `Listando N unidades` impresso contra a
  contagem sem abrir arquivo nenhum.
- Miniatura pequena não resolve o caso de uso: o dado que se precisa ler é o
  cabeçalho (`W700A VILLARS (10005) Protocolo de Recebimento...`), então a
  prévia é de um documento por vez, não uma grade densa.
- Renderizada sob demanda (ao selecionar), com cache por arquivo. Não
  pré-renderiza o lote: 15 arquivos seria barato, 500 não.
- Documento ilegível (arquivo de 0 byte, PDF corrompido) mostra o motivo no
  lugar da imagem, não uma área em branco.

## O que NÃO muda

- `logica.py` continua sem dependência de interface. O que for testável
  (validação da edição, recálculo do valor) vai pra lá; renderizar imagem e
  desenhar widget fica na interface.
- Formato da planilha gravada (`COLUNAS_PROTOCOLO`, `salvar_planilha_protocolo`).
- O carimbo em si (`_carimbar_protocolo`, bloco do Paybox, carimbo lateral).
- O botão "Gerar planilha do Superlógica" e `lancamentos_de_despesa`.
- A aba 1 (boletos/notas) não é tocada nesta etapa.

## Fora de escopo

- Reordenar/dividir/juntar páginas (o que o PDF Arranger faz) — não é o
  problema aqui.
- Grade de miniaturas do lote inteiro.
- Edição da planilha depois de fechado o painel: a grade é do lote em
  andamento, não um editor de xlsx avulso.
