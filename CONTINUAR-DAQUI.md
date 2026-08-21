# Continuar daqui — planilha de Despesas do Superlógica

Apagar este arquivo quando o trabalho terminar.

## Como retomar

```
git checkout despesas-superlogica
```

E dizer ao Claude: **"leia CONTINUAR-DAQUI.md e continue"**.

## O que fazer primeiro, e é a única coisa que trava tudo

**Importar `teste E - vencimento pelo programa.xlsx` no Superlógica.**

Está em `C:\Users\Dell\Downloads\TESTE CORREIO\Nova pasta`. Foi gerado pelo
programa já corrigido, usando o `Despesas.xlsx` do Downloads (o modelo sem
vencimento) e a data vinda da janelinha nova. Duas linhas: condomínio 44 com
R$ 57,75 e condomínio 42 com R$ 11,55, ambas com vencimento 21/08/2026.

- **Se as duas entrarem certas:** o recurso está pronto. Falta só atualizar o
  `CLAUDE.md` com o achado da data e fechar a v6.14.0 (merge, tag, exe,
  release) — o mesmo caminho da v6.13.0.
- **Se algo entrar errado:** anotar exatamente o que o Superlógica mostrou e
  investigar antes de mexer no código.

## Estado

- Branch `despesas-superlogica`, **8 commits à frente** da `main`, tudo
  commitado. **203 testes passando.**
- `main` está na **v6.13.0**, publicada com executável.
- O recurso está funcional de ponta a ponta; falta a confirmação no
  Superlógica e a documentação.

## O que o recurso faz

Botão **"Gerar planilha do Superlógica"** no rodapé do painel de resultado da
aba 3. Fluxo: confere as travas → **pergunta o vencimento** → você escolhe o
modelo e onde salvar → gera uma linha por protocolo cobrável, com
`condomínio` = ID SL do cadastro e `valor` = valor apurado.

## O bug que dominou a sessão, e a lição

Você gerou uma planilha com três lançamentos e **só um entrou no Superlógica**,
com vencimento 01/01/1970.

Causa raiz: o Excel guarda data como número de série — `21/08/2026` é `46255` —
e só o **formato da célula** diz que aquilo é uma data. O modelo tinha a célula
em `General`, então o arquivo saía com o número cru. O Superlógica lia `46255`,
gravava a época do Unix e recusava as linhas.

Minha primeira suspeita foi o campo `chave`, repetido nas três linhas. **Estava
errada.** Só não virou "conserto" porque testamos antes de mexer: dois arquivos
variando só a `chave` e um variando só a data. O da data resolveu.

Duas correções entraram:

1. **Guarda:** número numa coluna de data (`vencimento`, `competência`,
   `liquidação`) vira data de verdade se estiver na faixa 2000–2099; qualquer
   outra coisa levanta erro com o nome da coluna, em vez de virar cobrança com
   data errada.
2. **O programa pergunta o vencimento**, em vez de você digitar no modelo. A
   data é dado do lote, como a tarifa — deixá-la no modelo obrigava a editar o
   arquivo toda vez, e era nessa edição que ela virava número. A guarda é a
   rede; perguntar é o conserto na origem.

## Decisões fechadas (não rediscutir)

- Modelo lido do arquivo do usuário e **copiado**, nunca reconstruído.
- Mesmo condomínio em dois protocolos gera **duas linhas**.
- **Não gera nada** enquanto houver protocolo cobrável sem valor ou sem ID SL;
  a mensagem lista quais e onde resolver cada caso.
- Vencimento perguntado a cada geração, sem valor sugerido.
- Campo de data aceita `21/08/2026`, `21-08-2026`, `21.08.2026`, `21/08/26`.
  Recusa `2026-08-21` de propósito: misturar convenções é como `03/04` acaba
  lançada com o mês trocado.

## Sobre trocar para CSV

Você levantou, e a ideia é boa: esse bug é **impossível** em CSV, onde o texto
é o valor e não existe metadado escondido. O argumento que me fez escolher
xlsx — preservar formatos e validações do modelo — **caiu**: conferi as 32
colunas do seu modelo e todas estão em `General`, sem validação nenhuma.

Mantive xlsx porque agora está provado funcionando, e trocar de formato abriria
três incógnitas (separador de campo, separador decimal, codificação por causa
do acento em "bancária") para fechar uma já resolvida. Se aparecer outro
problema de metadado invisível, a recomendação inverte. Para decidir, o que
falta é um CSV modelo exportado do próprio Superlógica.

## Arquivos de teste gerados (podem ser apagados)

Em `C:\Users\Dell\Downloads\TESTE CORREIO\Nova pasta`:

| Arquivo | O que é |
|---|---|
| `teste correios.xlsx` | o seu, com o defeito — vencimento como número |
| `teste A - chave diferente.xlsx` | hipótese descartada |
| `teste B - chave vazia.xlsx` | hipótese descartada |
| `teste C - data corrigida.xlsx` | você corrigiu a data à mão; importou certo |
| `teste D - gerado com a correcao.xlsx` | programa corrigindo um modelo quebrado |
| **`teste E - vencimento pelo programa.xlsx`** | **o que importa testar** |

## Pendências herdadas

- Ninguém abriu o painel da aba 3 e olhou — a aparência nunca foi conferida por
  uma pessoa.
- Confirmar que o Superlógica lê o carimbo novo do PDF, que passou a trazer só
  `R$ 46,20` sem a conta, justamente para ele não capturar a tarifa.
- 10 dos 764 condomínios estão sem ID SL. Vai travar um lote em algum momento;
  o conserto é preencher na aba de Cadastro, ou reimportar via
  `tests/_importar_id_sl.py`, que é reexecutável.
