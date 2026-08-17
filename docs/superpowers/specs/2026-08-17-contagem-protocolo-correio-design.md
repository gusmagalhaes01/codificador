# Contagem e cobrança dos Protocolos dos Correios (aba nova) — especificação

Data: 2026-08-17
Estado: aprovado, a implementar

## Objetivo

Aba nova que lê uma pasta de Protocolos de Recebimento de Documento
(Correios/Imodata), conta quantas unidades cada protocolo entregou, multiplica
pela tarifa por linha informada no lote, carimba o valor no PDF e gera uma
planilha com uma linha por protocolo e o total geral.

Hoje esse cálculo é feito à mão: alguém conta as linhas e escreve o valor a
lápis no canto do papel. Nos quatro protocolos de referência (`TESTE CORREIO`,
lote de 24/07/2026) os manuscritos são 7,70 / 57,75 / 11,55 / 46,20 — todos
iguais a `unidades × 3,85`.

Não altera nada da codificação de PDFs (aba 1) nem da extração de NFS-e
(aba 2). As abas de Cadastro e Logs só são renumeradas.

## Por que aqui o OCR é aceitável (e na aba 2 não é)

A aba 2 não usa OCR de propósito: valor e data não têm dígito verificador,
então uma leitura errada entraria numa planilha financeira sem ninguém
perceber. Aqui o resultado também é dinheiro, mas o documento traz o próprio
conferidor impresso no rodapé da tabela — `Listando 15 unidades`. Ele é a
fonte do valor; a contagem das linhas serve para confirmá-lo. Divergiu, o
arquivo não é carimbado.

É a mesma estrutura que sustenta a escada de DPI em `extrair_cnpj_tomador`:
não se confia no OCR, confia-se em um invariante que o OCR precisa satisfazer.

Esses protocolos são escaneados — os quatro de referência têm zero caractere
de texto nativo.

## Funções novas em `logica.py`

Nenhuma dependência de interface, como o resto do módulo.

### `extrair_dados_protocolo_correio(texto)`

Devolve `None` se o documento não for um protocolo (o marcador
`MARCADOR_PROTOCOLO_CORREIO` não aparece). Senão, um dict:

| Chave | Origem |
|---|---|
| `codigo` | `extrair_codigo_protocolo_correio` (já existe) |
| `condominio` | texto antes do `(código)` no cabeçalho, ex: `W700A VILLARS` |
| `total_impresso` | `Listando N unidades`, ou `None` se ilegível |
| `linhas_contadas` | regex das linhas de unidade |
| `entregas_contadas` | ocorrências de `Correio` na coluna Entrega |

Regexes (validadas contra o `winocr` real nos quatro protocolos de
referência — ver "Como o OCR realmente devolve o texto"):

```python
RE_LISTANDO = re.compile(r"Listando\s+(\d+)\s+unidade", re.IGNORECASE)
RE_UNIDADE = re.compile(r"\b\d{1,4}(?:\s*[-–—])+\s*[A-Za-zÀ-ÿ]")
RE_ENTREGA = re.compile(r"\bCorreio\b", re.IGNORECASE)
```

`RE_UNIDADE` é aplicada só no trecho **antes** do `Listando`, para não contar
números do rodapé (CEP, telefone) como unidades. `RE_ENTREGA` é aplicada no
texto inteiro; `\b` impede que "Correios" (plural, no rodapé) seja contado.

Dois conferidores em vez de um porque cada um falha de um jeito diferente. No
protocolo de 15 unidades o OCR leu `702 - - Enny Marins de Lima`, com traço
duplicado — daí o `(?:\s*[-–—])+` no lugar de um traço só.

### Como o OCR realmente devolve o texto

O `winocr` entrega a página inteira **numa única linha**, sem `\n`, e com as
colunas fora de ordem — os "Correio" da coluna Entrega vêm todos no fim, depois
do rodapé. Exemplo real (protocolo do VILLARS, 300 DPI):

```
W700A VILLARS (10005) Protocolo de Recebimento de Documento Unidade
401 - Carmen Lucia Vides Gomes 402 - Carmen Lucia Vides Gomes
Listando 2 unidades Imodata - Condomínios e Imóveis Rua Barata Ribeiro,
774 / 100 andar Copacabana } RJ -22.051-002 matriz@imodata.net -
(21) 3816-7800 Assinatura Entrega Correio Correio 170 1 de 1
```

Nenhuma regex pode depender de início de linha (`(?m)^`) nem da ordem das
colunas. O `170` no fim é o OCR tentando ler o `7,70` manuscrito — ruído
esperado, e mais um motivo para o valor vir do cálculo e nunca do papel.

### Tesseract foi testado e descartado (2026-08-17)

Com `pytesseract` + Tesseract 5.4 e o modelo `por` (tessdata_best), os quatro
protocolos deram 4/4 em todos os sinais — inclusive a contagem por linha, que
o winocr não consegue produzir porque não devolve quebras de linha. Empate no
que este recurso usa (`Listando` e `Correio`, ambos 4/4 nos dois motores).

Descartado pelo custo: 5 a 9× mais lento (3,88s contra 0,45s no documento
maior), ~50 MB de executável mais 8 MB de modelo dentro do zip distribuído, e
a quebra do princípio de motor nativo sem programa externo instalado. A
vantagem dele — dados por linha — só passa a valer se algum dia for preciso
saber *quais* unidades receberam, não apenas quantas.

Só com o modelo inglês (o padrão do instalador) o Tesseract lia o título como
"Protocolo de Recebimento de Dacentanio" e o documento nem era reconhecido
como protocolo em 1 dos 4 casos.

Conferência das três contagens nos quatro protocolos de referência:

| Protocolo | Esperado | `Listando` | `RE_UNIDADE` | `RE_ENTREGA` |
|---|---:|---:|---:|---:|
| -001 VILLARS | 2 | 2 | 2 | 2 |
| -002 ASTORIA | 15 | 15 | 15 | 15 |
| -003 CARMEM | 3 | 3 | 3 | 3 |
| -004 DIDEROT | 12 | 12 | 12 | 12 |

### `conferir_contagem_protocolo(dados)`

Regra explícita de aceite: aceita `total_impresso` quando ele existe **e** pelo
menos um dos dois conferidores bate com ele. Devolve `(aceito: bool, motivo:
str)`. Não aceita nada quando `total_impresso` é `None` — sem o número
impresso não há o que conferir, e contar linhas por OCR sozinho é chute com
cara de precisão.

### `valor_protocolo(unidades, tarifa)`

`Decimal` internamente, `ROUND_HALF_UP`, duas casas. Dinheiro não passa por
float; a conversão para float acontece só na escrita da planilha.

### `montar_texto_valor_protocolo(unidades, tarifa, valor)`

`"15 un × R$ 3,85 = R$ 57,75"`. Mostra a conta, não só o resultado, para que a
conferência no papel não precise refazer a multiplicação.

### `COLUNAS_PROTOCOLO`, `linha_planilha_protocolo`, `salvar_planilha_protocolo`

Trio espelhado do que já existe para as NFS-e, incluindo o acabamento:
números como número (não texto), formato de moeda, painel congelado e
autofiltro.

## Mudanças em funções existentes

Todas com o comportamento atual preservado por padrão — o mesmo cuidado da
v6.9.0, quando `criar_overlay` ganhou `angulo=0`.

**`extrair_texto_ocr(caminho, max_paginas=2, dpi=300)`** passa a aceitar
`max_paginas=None` para ler o documento inteiro. O default continua 2, então
nenhum chamador atual muda. Sem isso, um protocolo de 3 páginas perderia
linhas em silêncio.

**`criar_overlay(...)`** ganha `alinhamento="esquerda"` (`"esquerda"`,
`"centro"`, `"direita"`). O parâmetro `centralizado=True` continua válido e
equivale a `"centro"`, preservando os chamadores e os testes existentes. O
alinhamento à direita usa `stringWidth` para recuar o texto a partir de `x`.

**`processar_pdf(caminho_entrada, caminho_saida, texto, config,
carimbos_extras=None)`** monta internamente
`carimbos = [principal] + (carimbos_extras or [])` e aplica todos num único
overlay por página. Chamada sem `carimbos_extras`, é byte por byte o que é
hoje. Cada item de `carimbos_extras` é um dict com `texto`, `x`, `y`,
`alinhamento`, `angulo`, `tamanho` e `cor`.

## Fluxo por arquivo

1. `extrair_texto_pdf`. Se vier texto nativo, usa — não gasta OCR.
2. Senão, OCR de **todas** as páginas (`max_paginas=None`).
3. `extrair_dados_protocolo_correio`. `None` → linha na planilha com a
   observação `Não é um protocolo dos Correios`; o PDF não é carimbado.
4. `conferir_contagem_protocolo`. Recusou → relê no `proximo_dpi_maior` e
   confere de novo, **uma vez**, como já se faz com CNPJ inválido. Recusou de
   novo → pendente: não carimba, não inventa valor, e a observação diz o que
   foi lido (`Listando 15, contei 14`).
5. Aceito → `valor_protocolo`, carimba, grava na pasta de saída com o mesmo
   nome do arquivo original, e escreve a linha na planilha.

## Carimbos

Dois por página, em **todas** as páginas, num único passe de escrita.

**Lateral rotacionado (código)** — `montar_texto_protocolo_correio(codigo,
nome, cnpj)` com `angulo=90`, idêntico à v6.9.0. Nada muda no que a IA do
Superlógica já reconhece. Depende do cadastro, porque precisa de nome e CNPJ.

**Topo direito (valor)** — `montar_texto_valor_protocolo(...)`,
`alinhamento="direita"`, recuado 28pt da borda direita e do topo. É o espaço
em branco do documento; o rodapé foi descartado porque os protocolos de duas
páginas têm conteúdo lá embaixo.

Aparência fixa no código (mesma fonte, tamanho e cor do carimbo lateral). Não
vira opção de configuração — a aba não usa perfis, e ninguém pediu ajuste
fino (YAGNI).

**Código não cadastrado:** carimba só o valor e registra `Código não
cadastrado` na observação. O valor não depende do cadastro; perder a cobrança
por falta de um cadastro seria pior que entregar o PDF com um carimbo só.

## A planilha

| Arquivo | Condomínio | Código | Unidades | Tarifa | Valor | Observação |
|---|---|---|---|---|---|---|

`Condomínio` vem do cadastro quando o código existe lá; senão, do cabeçalho do
próprio documento. Linha de **total** no rodapé somando `Unidades` e `Valor`.
Arquivos pendentes entram com `Unidades`, `Tarifa` e `Valor` vazios — nunca 0,
que significaria "entregou zero".

## A aba

Entra como **3. Protocolos dos Correios**; Cadastro vai para 4 e Logs para 5
(mesma renumeração que a v6.7.0 já fez uma vez).

Segue a linguagem visual das outras: Swiss, cantos retos, cobalto só no botão
primário, sem jargão técnico no texto visível (nada de "OCR" ou "DPI").

- Três campos com botão `Trocar`: pasta com os protocolos, pasta de saída dos
  PDFs carimbados, destino da planilha.
- Botão cobalto `Calcular N protocolos`, com N dinâmico e desabilitado sem
  pasta, igual ao `Processar N PDFs` da tela Principal.
- Ao clicar, um popup pede a **tarifa por linha**, campo vazio (sem valor
  padrão, por decisão do usuário — a tarifa muda com o tempo: nos protocolos
  de referência aparecem 3,45 e 3,85). Aceita `3,85` e `3.85`; valida número
  positivo.
- Barra de progresso e resumo ao final, no formato da aba 2, com a opção de
  abrir a planilha.
- Só a pasta escolhida, sem subpastas. Os PDFs de entrada não são alterados.

## Testes

`tests/test_protocolo_correio_contagem.py`, sobre fixtures de texto em
`tests/dados/` capturadas do OCR real dos quatro protocolos de referência
(nunca a planilha real; cadastro de teste como nos outros).

- contagem correta nos quatro protocolos (2, 15, 3, 12 unidades);
- `total_impresso` ausente → recusa, sem valor;
- conferidores divergentes → recusa, com o motivo legível;
- `RE_LINHA_UNIDADE` falhando e `Correio` salvando a conferência;
- documento que não é protocolo → `None`;
- `valor_protocolo` em `Decimal`, sem centavo perdido (15 × 3,85 = 57,75);
- texto dos dois carimbos;
- `linha_planilha_protocolo` com código fora do cadastro;
- `criar_overlay` com `alinhamento="direita"` e o texto ainda extraível;
- regressão: `criar_overlay` e `processar_pdf` sem os parâmetros novos
  produzem o mesmo resultado de antes.

## Aceitação

O lote real `TESTE CORREIO` precisa fechar em **7,70 / 57,75 / 11,55 / 46,20**,
total **R$ 123,20**, batendo com os valores manuscritos. Validar contra os
PDFs reais antes de liberar — mesma disciplina que pegou o bug do
`candidatos_por_nome`.

Verificação visual do carimbo depois de implementado: gerar um protocolo
carimbado e conferir a olho, como foi feito com o carimbo lateral.

## Fora de escopo

- Ler o valor manuscrito do papel — o cálculo é a fonte, o manuscrito é só a
  referência histórica que validou a tarifa.
- Mexer na aba 1: protocolos misturados a boletos continuam sendo
  identificados como hoje, sem cálculo de valor.
- Configuração de aparência do carimbo do valor.
- Lembrar as pastas ou a tarifa entre execuções.
