# Protocolo dos Correios do sistema antigo ("telnet") — especificação

Data: 2026-09-03
Estado: aprovado, a implementar

## Objetivo

Fazer a aba 3 reconhecer, contar e cobrar o **protocolo de correspondência do
sistema antigo da Imodata** (chamado internamente de "telnet"), que hoje não é
reconhecido por nada no programa e cai como arquivo ignorado.

Ele chega **misturado**, na mesma pasta, com o protocolo novo (o do
Superlógica, que a v6.8.0 em diante já lê). Então o programa precisa
distinguir os dois sozinho, arquivo a arquivo.

## O documento

Escaneado, sem texto nativo — como o protocolo novo. Uma página. Estrutura:

```
*** IMODATA  * PROTOCOLO CORRESPONDENCIA CORREIO NORMAL   DATA: 11/08/2026 ***
    EDF:  MENESCAL                  * END:  AV N S DE COPACABANA 664
    REF:  BAL 07/2026               * CTR:  PB608111.121

                    R E L A C A O   D A S   U N I D A D E S
    LJ-01   LJ-02   LJ-08   ...   AP-1005   664B

  TOTAL ENVIADO PELO CORREIO..:  27 COD.: 1.1122.8)
```

Abaixo da folha vêm colados dois comprovantes térmicos dos Correios (AGF
DINÂMICA), com o valor pago.

### Como ele difere do protocolo novo

| | Protocolo novo | Telnet |
|---|---|---|
| Marcador | `Protocolo de Recebimento de Documento` | `PROTOCOLO CORRESPONDENCIA` |
| Código do condomínio | `W700A VILLARS (10005)` | `COD.: 1.1122.8)` — com pontos |
| Contagem | `Listando N unidades` | `TOTAL ENVIADO PELO CORREIO..: 27` |
| Unidades | `702 - Nome do morador` + coluna "Correio" | grade `LJ-01 AP-204 …`, sem nomes |
| Nome do condomínio | no mesmo campo do código | `EDF: MENESCAL`, campo separado |

## Escopo: só o telnet

A pasta de amostras (`Correios/telnet`, 13 arquivos) contém **dois** formatos:
7 telnet e 6 de um terceiro tipo, o **"Meus Correios"**
(`Cód. Condomínio: 10852-VILLA BRANCA`, `Qtde.`, `Classificação`).

O "Meus Correios" **fica fora deste spec**, por decisão do usuário: são SEDEX
com valor fixo, cobrados por outro critério, e não se misturam com as cartas
simples. Registre-se que ele é o formato **fácil** — o OCR lê o código em 6/6,
e o nome vem junto do código no mesmo campo, servindo de conferência. Quando
for a vez dele, começar daí.

## 1. Discriminador

Antes de extrair, decidir o formato pelo marcador presente no texto do OCR:

- `PROTOCOLO CORRESPONDENCIA` → telnet
- `Protocolo de Recebimento de Documento` → novo (caminho atual, intocado)
- nenhum dos dois → segue o fluxo de hoje (não é protocolo)

**Este é o ponto mais perigoso do spec, e o único ainda não medido.** Errar o
valor o usuário corrige na tela; classificar errado faz o programa aplicar a
extração errada e produzir um resultado plausível pelo motivo errado — não há
nada visivelmente estranho para a conferência humana pegar. Validar contra um
lote misto **antes** de liberar.

## 2. Extração

`extrair_dados_protocolo_telnet(texto, cadastro)`, em `logica.py`, devolvendo
código, nome conferido e contagem. Três decisões, todas vindas de medição
contra os 7 arquivos reais — nenhuma de intuição.

### 2.1 O código é achado pelo formato, não pelo rótulo

**Não ancorar em `COD.`.** O OCR estraga o rótulo muito mais do que os
dígitos: leu `DAS`, `D ÀS` e `COD. .` nos mesmos documentos em que os dígitos
saíram perfeitos. Ancorado no rótulo, o código saía em **4 de 7**; pelo
formato, em **7 de 7**.

O padrão é `1.DDDD` — todos os 772 códigos do cadastro têm exatamente 5
dígitos e começam em `1` (faixa 10002–11242), então `1\.\d{4}` é distintivo
sozinho. O que vier depois (`.8)`, `.7)`) é ignorado, por instrução do
usuário.

### 2.2 O total é ancorado em "PELO CORREIO", com folga curta

**Não ancorar em `ENVIADO`**: o OCR devolve `EWIADO` e `EFvTIADO`. `PELO
CORREIO` nunca falhou nas 21 leituras.

A folga entre a âncora e o número é de **6 não-dígitos, nunca 20**. Foi a
folga larga que, no arquivo 011, deixou capturar um `38` vindo de outro canto
da folha em vez do `3` correto — o número certo estava ali, mas o regex pulou
até um mais distante.

### 2.3 O nome é procurado no texto inteiro, não depois do `EDF:`

O `winocr` devolve a página **numa linha só e com as colunas fora de ordem** —
comportamento já registrado no `CLAUDE.md` a propósito da contagem do
protocolo novo. Por isso `EDF:` e o valor `MENESCAL` não ficam adjacentes:
saem coisas como `EDF : REF; AP-104 TOTAL CHATEAU FONTAINEBLEA`.

Ancorar no rótulo devolveu nome em 1 de 7. Procurar cada nome do cadastro
dentro do texto todo, por similaridade, achou o certo em **6 de 7 com score
1.00**.

**O nome conta como confirmado a partir de 0.90** (`difflib.SequenceMatcher`,
os dois lados normalizados — sem acento, só letras e espaço, em caixa alta).
O limiar não é delicado: nos 7 arquivos os confirmados deram **1.00** e o
único não confirmado deu **0.79**. Qualquer corte entre 0.80 e 0.99 daria o
mesmo resultado, e 0.90 fica no meio dessa folga.

Só se compara o nome do **código já escolhido** — não se usa o nome para
*achar* o condomínio. Isso é deliberado e segue a regra de ouro do projeto
registrada em "Cuidado: condomínios com nomes parecidos": nome serve para
confirmar ou levantar dúvida, nunca para identificar sozinho.

## 3. Três leituras e votação por maioria

Nenhuma configuração isolada lê os 7 arquivos. As três combinadas leem:

1. página inteira, 300 DPI
2. recorte do topo (0–35%), 400 DPI
3. recorte do topo (0–32%), 500 DPI

Código e total saem por **maioria** entre as três, não pela primeira que
casar. A votação não existe para perseguir 100% de acerto — existe porque foi
ela que transformou `38, 3, 3` em **3** no arquivo 011. Sem ela, `R$ 146,30`
chegaria à tela num condomínio de três unidades, e é o tipo de número que
passa despercebido num lote de 100.

**As duas regras de apuração, explícitas:**

- **Total:** exige **pelo menos 2 das 3 leituras concordando**. Com três
  valores diferentes, não há total — a linha vai para pendente. É o número que
  vira dinheiro, e uma leitura solitária não sustenta isso. Nos 7 arquivos,
  todos os totais preenchidos tiveram 2 ou 3 votos.
- **Código:** o mais votado entre as leituras que acharam um código **existente
  no cadastro**; basta 1 voto, porque o código não vira valor sozinho — ele
  ainda passa pela conferência do nome, que é quem marca a linha. Empate entre
  dois códigos diferentes → pendente, sem escolher nenhum.

Custo: ~3× o OCR de hoje por arquivo. A aba 3 já roda em thread com barra de
progresso.

## 4. Aceite: preencher e marcar, não barrar

**O usuário confere imagem por imagem.** Isso é premissa do desenho, dada por
ele, e é o que define a regra abaixo — que é deliberadamente mais frouxa que a
do protocolo novo.

| Situação | Comportamento |
|---|---|
| Código + nome + total concordam | preenche, linha limpa |
| Código e total, nome não confirma | **preenche**, observação `Nome não confirmado` |
| Sem total legível | pendente (não há o que preencher) |
| Código não encontrado no cadastro | pendente |

Uma versão anterior deste desenho **barrava** a linha sem confirmação de nome.
Foi abandonada: no arquivo 010, código e total estavam **ambos corretos** e a
linha iria para pendente por causa de um nome mal lido. Com conferência
humana, o portão custa mais do que rende.

A marcação continua existindo porque numa pilha de 100 todo mundo olha com
mais atenção o que está sinalizado. Ela diz **onde** o programa ficou em
dúvida, em vez de esconder a dúvida atrás de um valor com cara de certo.
Custo zero: a coluna `Observação` já existe e a grade já colore linha por
estado.

### Por que o cadastro NÃO serve de dígito verificador

Registrado porque contraria a intuição e vai tentar alguém no futuro: os
códigos vão de 10002 a 11242 e **772 dos ~1.241 números dessa faixa são
condomínios reais — 62% de densidade**. Um dígito errado no OCR tem grande
chance de produzir *outro condomínio existente*. "O código está no cadastro"
não é prova de nada. É por isso que o nome impresso, ainda que só como
marcação, tem valor.

## 5. O que não muda

Valor `unidades × tarifa`, carimbo lateral, bloco do Paybox, planilha de
cobrança, planilha de despesas, painel com grade e prévia. O telnet só
acrescenta **mais uma origem de `(código, contagem)`**; daí para frente o
fluxo é o mesmo do protocolo novo.

A conferência imagem por imagem já é servida pelo painel da v6.18.0 — grade e
documento lado a lado, com zoom e arraste.

## Validação feita

7 arquivos telnet reais, verdade conferida lendo a folha:

| Arquivo | Verdade | Voto do total | Código | Resultado |
|---|---|---|---|---|
| 007 | 27 · 11122 | 27 | 11122 ✓ | preenche |
| 008 | 8 · 11127 | não leu | 11127 ✓ | pendente |
| 009 | 10 · 11121 | 10 | 11121 ✓ | preenche |
| 010 | 8 · 10151 | 8 | 10151 ✓ | preenche, `Nome não confirmado` |
| 011 | 3 · 11100 | **3** (era 38) | 11100 ✓ | preenche |
| 012 | 5 · 11098 | 5 | 11098 ✓ | preenche |
| 013 | 1 · 11093 | 1 | 11093 ✓ | preenche |

**Código 7/7. Total 6/7 correto, 1 não lido. Nenhum valor errado preenchido.**
6 de 7 preenchidos (86%).

## Caminhos testados e descartados

Vale registrar, porque todos eram plausíveis e custaram medição:

1. **Ancorar o código em `COD.`** — 4/7. O OCR estraga o rótulo.
2. **Ancorar o nome em `EDF:`** — 1/7. Colunas embaralhadas.
3. **Folga de 20 não-dígitos após `CORREIO`** — capturou `38` no lugar de `3`.
4. **Usar a contagem da grade de unidades como conferidor** — descartado por
   decisão do usuário ("ignore a grade"). Medido antes: subconta em grades
   grandes (22–24 de 26 no arquivo 007), embora acerte nas pequenas.
5. **Ler o valor impresso no comprovante colado** — os comprovantes são
   térmicos e desbotados; o valor (`103,95`) **não foi lido em DPI nenhum**.
   O valor continua sendo `unidades × tarifa`.

## Testes

Fixtures de texto em `tests/dados/`, sintéticas, como as de NFS-e e as do
protocolo novo — nunca PDF real no repositório. Cobrir:

- o discriminador escolhe telnet, novo ou nenhum, para os três tipos de texto
- código extraído do padrão `1.DDDD` mesmo com o rótulo corrompido (`DAS`)
- total lido com `ENVIADO` corrompido (`EWIADO`, `EFvTIADO`)
- a folga curta **não** captura número distante (o caso do `38`)
- votação: `38, 3, 3` devolve `3`
- nome achado no texto embaralhado, longe do `EDF:`
- sem total legível → pendente; código fora do cadastro → pendente
- nome não confirmado → preenche **com** a observação

## Riscos e o que ainda não se sabe

- **O discriminador não foi medido** contra um lote misto. É o primeiro
  item a validar, antes de qualquer liberação.
- **São 7 amostras, de uma competência só.** A v6.18.1 nasceu exatamente
  desse tipo de ponto cego: o `CLAUDE.md` registra que "a quebra depende da
  largura da coluna no lote, não da versão do DANFSe". Sete arquivos não
  cobrem variação de layout, prefixo de unidade fora de `LJ`/`AP`/`SL`, nem
  telnet com mais de uma página — nenhum dos 7 tem.
- **O `.8)` depois do código** foi assumido como ignorável, por instrução do
  usuário. Se um dia significar algo, é aqui que está a suposição.
