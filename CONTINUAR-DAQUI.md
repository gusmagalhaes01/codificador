# Continuar daqui — planilha de Despesas do Superlógica

Apagar este arquivo quando o trabalho terminar.

## O que fazer

```
git checkout despesas-superlogica
```

Pedir ao Claude: **"escreve o plano de implementação do spec de despesas e executa com subagente por task"**.

O spec já está aprovado em
`docs/superpowers/specs/2026-08-19-despesas-superlogica-design.md`. Não precisa
rediscutir o desenho — só planejar e implementar.

## Estado

- Branch `despesas-superlogica`: só o spec, nenhuma linha de código.
- `main` em **v6.13.0**, publicada com executável. 164 testes passando.

## O recurso, em uma frase

Botão **"Gerar planilha do Superlógica"** no painel de resultado da aba 3, que
transforma o lote de protocolos na planilha de importação de despesas: uma
linha por protocolo cobrável, `condomínio` = ID SL do cadastro, `valor` = valor
apurado, demais campos copiados de uma linha-molde.

## Fatos já verificados — não redescobrir

Conferidos contra os arquivos reais nesta máquina:

- **Modelo:** `C:\Users\Dell\Downloads\Despesas.xlsx`, 32 colunas.
  `condomínio` é a coluna 1 (com acento no cabeçalho: `'condomínio'`), `valor`
  é a coluna 10. Mesmo assim o código deve achá-las **pelo nome normalizado**,
  não pela posição — o layout é do Superlógica e pode mudar.
- **Linha 2 do modelo** tem seis células preenchidas: colunas 4 (fornecedor),
  5 (favorecido), 6 (conta_categoria), 7 (tipo_de_documento),
  11 (forma_de_pagamento) e 16 (chave = `46`). É o molde, e **é substituída**
  pela primeira linha real — não pode sobrar linha meio vazia.
- **`carregar_cadastro` já devolve `id_sl`** em cada registro (string, vazia
  quando não há). Nenhuma mudança necessária ali além da docstring, que ainda
  afirma que o campo "não é usado por nada".
- **`openpyxl` 3.1.5** copia estilo entre células com
  `celula._style = copy.copy(outra._style)` — testado, funciona.
- **`tests/cadastro_teste.py` ainda não tem `id_sl`.** Precisa ganhar o campo,
  com um dos condomínios deixado sem ele de propósito (sugestão: LAGO
  MAGGIORE), para cobrir o caso de cadastrado sem ID.

## Correção a fazer no spec durante o planejamento

O spec escreve `lancamentos_de_despesa(linhas, cadastro)`, recebendo as linhas
da planilha de protocolos. **Isso não funciona.** Nas linhas prontas, um
arquivo que não é protocolo e uma pendência do tipo "não foi possível ler o
documento" ficam idênticos — código, condomínio e valor todos vazios — mas um
deve travar a geração e o outro deve ser ignorado.

A função precisa receber o **`resultado` do painel**, que já separa
`processados`, `pendentes` e `ignorados`.

## Decisões fechadas (não rediscutir)

- Modelo lido do arquivo do usuário e **copiado**, nunca reconstruído.
- Mesmo condomínio em dois protocolos gera **duas linhas** (o 11161 MARILIA sai
  como `496 / 138,60` e `496 / 119,35`).
- **Não gera nada** enquanto houver protocolo cobrável incompleto — pendência
  sem valor ou condomínio sem ID SL. A mensagem lista quais e onde resolver.
- `vencimento` e `competência` continuam em branco, como no modelo.
- Botão neutro (contornado); cobalto é reservado ao que resolve pendência.

## Ponto em aberto, sem pressa

A trava para o lote inteiro por causa de **um** condomínio sem ID SL, e **10
dos 764 estão sem**. Vai acontecer. Três saídas: preencher na aba de Cadastro
quando travar; afrouxar a trava (gerar sem os incompletos, avisando); ou
preencher os 10 de uma vez a partir de um export do Superlógica, com
`tests/_importar_id_sl.py`, que é reexecutável.

## Lotes para teste (fora do repositório)

- `C:\Users\Dell\Downloads\TESTE CORREIO` — quatro de uma página, fecham em
  R$ 123,20 à tarifa de R$ 3,85.
- `C:\Users\Dell\Downloads\TESTE CORREIO\Nova pasta` — quatro multipágina; três
  fecham em R$ 488,95 e o do 11049 fica pendente de propósito (alguém riscou o
  total impresso e escreveu 78 à caneta).

## Pendências herdadas da v6.13.0

- Ninguém abriu o painel novo da aba 3 ainda — a aparência nunca foi conferida
  por uma pessoa.
- Confirmar que o Superlógica lê o carimbo novo, que passou a trazer só
  `R$ 46,20` sem a conta, justamente para ele não capturar a tarifa.
