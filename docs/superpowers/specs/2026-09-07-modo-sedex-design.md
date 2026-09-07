# Modo SEDEX e leitura do protocolo "Meus Correios" — especificação

Data: 2026-09-07
Estado: aprovado, a implementar

## Objetivo

Automatizar a parte automatizável dos envios avulsos dos Correios (SEDEX e
correio registrado): **o programa identifica o condomínio, carimba e monta a
planilha; a pessoa digita o valor.**

Duas peças, que juntas atendem o caso real:

1. **Leitor do formato "Meus Correios"** (gerado pelo sistema Agile), hoje
   não reconhecido — os arquivos caem como "não é protocolo" e são ignorados.
2. **Modo SEDEX**, um alternador de lote que desliga a contagem de unidades
   nos formatos que contam.

## Por que o valor é digitado, e não lido

Não é limitação aceita por preguiça — é medição. O valor desses envios vem
impresso num **comprovante escaneado**, e comprovante escaneado é exatamente o
que o OCR não lê. Está registrado na investigação do telnet: os comprovantes
térmicos colados nas folhas trazem `TOTAL: 27  103,95`, e o valor **não foi
extraído em DPI nenhum** (200, 300, 400). O `CLAUDE.md` já generaliza o
princípio na aba 2: valor e data não têm dígito verificador, então
`1.234,56` lido como `1.234,58` entraria numa planilha financeira sem
ninguém perceber.

Some-se que os valores **variam por envio** — não há tarifa fixa a
multiplicar, ao contrário das cartas simples.

## A decisão estruturante: modo e formato são ortogonais

```
FORMATO (novo / telnet / meus correios) → COMO ler o condomínio → o documento diz
MODO    (normal / SEDEX)                → SE conta unidades     → o documento NÃO diz
```

O mesmo envio SEDEX pode chegar nos três formatos. Nada na folha distingue um
protocolo de cartas simples de um de SEDEX — por isso o modo é **um alternador
que o usuário liga**, e não algo deduzido do documento.

| Formato | Modo normal | Modo SEDEX |
|---|---|---|
| novo (Superlógica) | conta unidades | não conta, valor à mão |
| telnet | conta unidades | não conta, valor à mão |
| **meus correios** | **não conta (intrínseco)** | não conta |

**"Meus Correios" nunca conta, em modo nenhum.** Esse formato não tem lista de
unidades — tem `Qtde. 1`, um envio por documento. Contar ali não faz sentido
em situação alguma, então isso é propriedade do formato e não do modo. Quem for
mexer nisso depois: não é um esquecimento, é deliberado.

## 1. Leitor do "Meus Correios"

### O documento

```
Correios                    Meus Correios              Página 1 de 1
ID     Data        Status   Cód. Condomínio      Qtde.  Classificação        ...
12700  13/08/2026  Gerado   10852-VILLA BRANCA   01     824 - EBCT - SEDEX   ...
                                                 Total: 1
```

Aparece em dois arranjos — tabela (como acima) e lista vertical
(`Correios [12704]` com os campos empilhados). São o mesmo sistema renderizado
em larguras diferentes; os rótulos são idênticos.

### Extração

`Cód. Condomínio: 10852-VILLA BRANCA` traz **código e nome no mesmo campo** —
conferência de graça, melhor do que a de qualquer outro formato: o código
resolve pelo cadastro e o nome ao lado confirma, sem precisar varrer a página.

Medido nos 6 arquivos reais: **código válido em 6/6**, nome batendo com o do
cadastro nos 6.

O regex precisa tolerar espaço no meio dos dígitos — o OCR devolve
`1 0193-MONACO` e `10520- SENADOR LEITE OITICICA`.

### Discriminador

`classificar_formato_protocolo` (`logica.py`) ganha o quarto valor
`"meus_correios"`, pelos marcadores `Meus Correios` / `Cód. Condomínio` /
`Correios [NNNNN]`. A regra de ambiguidade continua: dois marcadores de
formatos diferentes no mesmo texto → `"ambiguo"` → pendente.

## 2. Modo SEDEX

### Como liga

Alternador na aba 3, no mesmo lugar e formato do "Separar a saída em lotes de
N arquivos" — foi a referência pedida pelo usuário.

Rótulo: **"Modo SEDEX — não contar unidades, valor digitado à mão"**. O texto
diz o que o modo FAZ, não só o nome dele: quem liga sem saber precisa
entender a consequência antes de rodar o lote, e "Modo SEDEX" sozinho não
explica que a contagem some. Segue a convenção de vocabulário para leigos do
projeto — nada de "OCR" ou "DPI" em texto visível.

**Não persiste entre sessões.** O de lotes persiste porque é preferência do
sistema de destino; este é propriedade do lote que está na mesa agora. Nasce
desligado.

**Com o modo desligado e o lote todo em "Meus Correios"**, a tarifa ainda é
perguntada e simplesmente não é usada — esse formato não conta unidades de
qualquer jeito. É desperdício de uma pergunta, não erro; suprimir isso
exigiria saber o formato do lote antes de abrir o primeiro arquivo, o que o
programa não tem como fazer.

### O que ele muda

Só isto, e nada mais:

- não chama `conferir_contagem_protocolo` (nem a do telnet)
- não calcula valor — Unidades e Tarifa saem **vazias** na planilha, como já
  acontece no caso "Valor informado manualmente" da v6.13.0
- a **tarifa não é perguntada** no início do lote: não há o que multiplicar.
  Vencimento e chave continuam sendo, porque o bloco do Paybox e a planilha
  de despesas dependem deles.

### O que ele NÃO muda

Identificação do condomínio, carimbo lateral, separação em lotes, painel,
grade editável, planilha de cobrança, planilha de despesas. O modo é
**subtrativo**.

## 3. A linha nasce sem valor, e a máquina para isso já existe

A linha entra como pendente sem valor. O usuário digita na coluna Valor da
grade, e o caminho existente cuida do resto:

- `_recarimbar_linha` regrava o PDF com o valor **e** o bloco do Paybox —
  então o anexo automático volta a funcionar, o que não aconteceria se o
  valor fosse preenchido no Excel depois
- `reclassificar_registro` move a linha para `processados` (com valor →
  processados, sem valor → pendentes)
- `lancamentos_de_despesa` destrava e a planilha de despesas sai

Nada disso é novo. O caminho "carimba só a identificação agora, o valor entra
depois" nasceu na v6.17.1, quando `valor_do_carimbo` passou a devolver `None`
para valor desconhecido.

## 4. A observação carrega o serviço

O documento traz `Classificação: 824 - EBCT - SEDEX` ou
`590 - EBCT - CORREIO REG / AR`. Esse texto vai para a coluna **Observação**.

Motivo prático: é a informação que a pessoa precisa para saber **quanto**
digitar, e ela está no papel mas não estaria na planilha. Sem isso, conferir
linha a linha exigiria voltar ao PDF a cada valor.

Nos 6 arquivos medidos, **5 são `CORREIO REG / AR` e só 1 é SEDEX** — apesar
de o lote inteiro ser chamado de "os sedex". São serviços diferentes, com
preços possivelmente diferentes, e é por isso que o valor é por linha e não
um só para o lote.

**Quando já existe outra observação, a do serviço vem primeiro**, separada por
`"; "` — o mesmo separador que `linha_planilha_protocolo` já usa ao juntar a
sua. A ordem não é estética: numa linha pendente, o texto que explica *o que
fazer* ("Código do condomínio não confirmado") tem que sobreviver ao corte da
coluna, e vir por último não garante isso. O serviço é informação de apoio;
a pendência é ação.

**A classificação sai como está no documento** (`824 - EBCT - SEDEX`), sem
traduzir nem encurtar: é o texto que a pessoa vai cruzar com a tabela de
preços dos Correios, e reescrevê-lo cria uma diferença a explicar.

## Testes

Fixtures de texto inline, como as dos outros formatos — nunca PDF no
repositório. Cobrir:

- discriminador reconhece `"meus_correios"` nos dois arranjos (tabela e
  vertical) e não confunde com telnet nem com o novo
- código extraído com espaço no meio dos dígitos (`1 0193-MONACO`)
- código extraído com espaço depois do traço (`10520- SENADOR LEITE`)
- nome do documento conferido contra o do cadastro
- código fora do cadastro → pendente
- modo SEDEX: linha sai com Unidades e Tarifa vazias e sem valor
- modo SEDEX **não** altera o caminho dos outros formatos em modo normal
- a classificação do serviço chega à observação

## Riscos e o que não está validado

- **Só há amostra de um dos três formatos.** O usuário informou que SEDEX
  também chega em telnet e no formato do Superlógica, mas não tem exemplos
  hoje. O desenho cobre os três; a validação cobre um. É o mesmo ponto cego
  que produziu a v6.18.1 — lá, validar contra um lote de uma competência só
  escondeu uma variação de layout que quebrava a leitura inteira.
- **São 6 arquivos, de um mês só.** Não cobrem variação de layout, serviço
  fora de SEDEX/REG-AR, nem documento com mais de uma linha na tabela — todos
  os 6 têm `Qtde. 1` e uma linha.
- **O modo desligado num lote SEDEX, ou ligado num lote normal**, falha de
  forma ruidosa nos dois sentidos (tudo cai em pendente), não silenciosa. Foi
  o que dispensou uma trava mais elaborada.
