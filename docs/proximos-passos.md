# Onde parei — 2026-08-24

Nota para retomar do zero, em outra máquina, sem o contexto da conversa.

## O que ficou pronto e commitado

**Condomínios identificados por CPF** (`2aa6e80`, `fa83330`, `a63e0b9`, `2a74b82`)
A chave do cadastro passou a ser "documento": 14 dígitos (CNPJ) ou 11 (CPF do
síndico). 8 condomínios já cadastrados por CPF, vindos do export do Superlógica.

**Painel dos protocolos sem piscar** (`2dc3572`)
Informar valor deixou de remontar o painel inteiro. Verificado na janela: não
pisca mais.

**Bloco do Paybox nos protocolos** (`0bcaee1`)
Os protocolos passam a sair com um bloco carimbado (fornecedor, CNPJ do
fornecedor, vencimento, valor) que faz o Paybox anexar o documento à despesa
sozinho. O vencimento passou a ser perguntado no **início** do lote, junto da
tarifa. Detalhes e o porquê de cada decisão estão na seção "Anexo automático no
Paybox" do `CLAUDE.md`.

## O que falta — em ordem de prioridade

### 1. Rodar o fluxo completo da aba 3 no app (não testado ponta a ponta)

O bloco do Paybox foi validado **fora** do programa: os PDFs de teste foram
montados à mão e importados no Superlógica, e associaram. O caminho dentro do
app (pergunta do vencimento → carimbo → planilha) só foi verificado por teste
automatizado e por um carimbo isolado. Falta rodar de verdade:

```bash
python identificacao_por_cnpj_6_0.py
```

Na aba "3. Protocolos", processar um lote e conferir:
- a pergunta do **vencimento** aparece logo depois da tarifa, no início;
- o PDF carimbado sai com o bloco de 4 linhas legível, sem cobrir conteúdo;
- "Gerar planilha do Superlógica" **não** pergunta o vencimento de novo, e a
  coluna `vencimento` da planilha traz a mesma data que está carimbada;
- importar a planilha, depois mandar os PDFs ao Paybox, e confirmar que
  anexaram sozinhos.

### 2. Conferir a posição do bloco num protocolo com recibo colado

`X_BLOCO_PAYBOX` / `Y_BLOCO_PAYBOX` (`identificacao_por_cnpj_6_0.py`) põem o
bloco na faixa em branco abaixo da tabela de unidades. Isso foi validado num
protocolo simples (VILLARS). **Alguns protocolos trazem um recibo dos Correios
colado no meio da folha** — o do JARDIM DOS ACANTUS é um deles. Nesses, o bloco
pode cair em cima do recibo. Conferir e, se precisar, ajustar as duas
constantes; nada mais depende delas.

### 3. Verificar a aba de Cadastro na janela (pendente desde a mudança do CPF)

Não existe teste automatizado de interface no projeto. Conferir na janela:
- documento com 11 dígitos (ex: `529.982.247-25`) é aceito;
- `123` é recusado, com mensagem citando CNPJ **e** CPF;
- **selecionar um condomínio cadastrado por CPF, trocar o documento por um CNPJ
  e salvar → a tabela tem que continuar com UMA linha, não duas.** É a correção
  de duplicação; se aparecerem duas, ela falhou.

### 4. Gerar o executável

O `dist/` está desatualizado — nada das mudanças acima chegou a quem usa o
`.exe`. Duplo clique em `gerar_exe.bat` (instala dependências, roda os testes e
só empacota se passarem).

## Pendências fora do código

- **Apagar os lançamentos de teste no Superlógica** (VILLARS, R$ 7,70, com e sem
  Pix). Um deles ficou com `chave_pix` preenchida com um BR Code de chave
  aleatória inexistente — não é pagável, mas não deve ficar lá.
- Pasta `C:\Users\Dell\Downloads\TESTE QR PAYBOX` e
  `C:\Users\Dell\Downloads\TESTE CODIGO BARRAS` são só de teste, podem ir fora.

## Coisas que já foram descartadas — não refazer

- **QR Code / código de barras para associação**: não funciona. Code128, I2of5 e
  Code39 nem são lidos; QR é lido mas não associa, nem com BR Code Pix válido.
  A associação é por valor + vencimento + fornecedor.
- **`etiqueta_paybox` como identificador livre**: exige URL do `sldocs.com.br`.
- **Colar link do sldocs manualmente**: destrói o ganho, é o mesmo trabalho
  manual de hoje.
- **Ler o número escrito à caneta** no protocolo do 11049: decisão de deixar
  como está, resolve pelo "Informar valor".
- **LIMA BARROS com CPF de 14 dígitos** (zeros à esquerda na máscara de CNPJ):
  decisão de ignorar. Ele não é identificado pelo conteúdo; cai no match por
  código/nome do arquivo.
