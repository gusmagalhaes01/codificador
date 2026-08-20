# Continuar daqui — planilha de Despesas do Superlógica

Anotado em 2026-08-19. Apagar este arquivo quando o trabalho terminar.

## Onde parou

A branch **`despesas-superlogica`** tem só um commit: o spec, aprovado, em
`docs/superpowers/specs/2026-08-19-despesas-superlogica-design.md`.
Nenhuma linha de código foi escrita ainda.

A `main` está em **v6.13.0**, publicada com executável na
[release v6.13.0](https://github.com/gusmagalhaes01/codificador/releases/tag/v6.13.0).
158 testes passando.

## Para retomar

```
git checkout despesas-superlogica
```

E dizer ao Claude: *"escreve o plano de implementação do spec de despesas"*.
Daí segue o mesmo caminho da v6.13.0 — plano, execução com subagente por task,
revisão em cada uma, revisão ampla no fim.

## O que o recurso faz

Transforma o lote de protocolos dos Correios na planilha de importação de
despesas do Superlógica. Uma linha por protocolo cobrável, com `condomínio`
preenchido pelo **ID SL** (4ª coluna do cadastro) e `valor` com o valor apurado.
Os outros campos se repetem, copiados de uma linha-molde.

Botão **"Gerar planilha do Superlógica"** no painel de resultado da aba 3.

## Decisões já tomadas (estão no spec, não precisam ser rediscutidas)

- **O modelo vem do arquivo do usuário e é copiado, não reconstruído.** O
  layout é do Superlógica e pode mudar; copiar preserva formatos e validações
  que o importador pode exigir.
- **Colunas achadas pelo nome no cabeçalho**, não pela posição.
- **A linha 2 do modelo é molde e é substituída** pela primeira linha real —
  não sobra linha meio vazia no arquivo.
- **Mesmo condomínio em dois protocolos gera duas linhas** (o 11161 MARILIA sai
  como `496 / 138,60` e `496 / 119,35`), para cada lançamento continuar
  rastreável até o papel.
- **Não gera nada enquanto houver protocolo cobrável incompleto** — pendência
  sem valor ou condomínio sem ID SL. A mensagem lista quais e onde resolver
  cada caso.
- `vencimento` e `competência` continuam em branco, como no modelo.

## Ponto em aberto, para decidir com calma

A trava acima para o lote inteiro por causa de **um** condomínio sem ID SL, e
**10 dos 764 estão sem o campo**. Isso vai acontecer em algum momento. Três
saídas, nenhuma urgente:

1. deixar como está e preencher o ID SL na aba de Cadastro quando travar;
2. afrouxar a trava: gerar assim mesmo, deixando os incompletos de fora com
   aviso;
3. adiantar o problema preenchendo os 10 que faltam a partir de um export de
   condomínios ativos do Superlógica — foi assim que a coluna nasceu, via
   `tests/_importar_id_sl.py`, que é reexecutável.

## Arquivos de referência

- Modelo do Superlógica usado no desenho: `C:\Users\Dell\Downloads\Despesas.xlsx`
  (fora do repositório) — 32 colunas, linha 2 com os seis campos que se repetem.
- Lotes de protocolo para teste, também fora do repositório:
  `C:\Users\Dell\Downloads\TESTE CORREIO` (quatro de uma página, fecham em
  R$ 123,20 à tarifa de R$ 3,85) e `...\TESTE CORREIO\Nova pasta` (quatro
  multipágina; três fecham em R$ 488,95 e o do 11049 fica pendente de
  propósito, porque alguém riscou o total impresso e escreveu 78 à caneta).

## Pendências herdadas da v6.13.0

- **Ninguém abriu o painel novo da aba 3 ainda.** Os agentes exercitaram a tela
  programaticamente, mas a aparência não foi conferida por uma pessoa.
- **Confirmar que o Superlógica lê o carimbo novo.** Ele passou a trazer só
  `R$ 46,20`, sem a conta, justamente para o Superlógica não capturar a tarifa
  no lugar do total. Só usando dá para saber se resolveu.
