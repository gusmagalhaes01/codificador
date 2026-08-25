# Onde parei — 2026-08-25

Nota para retomar do zero, em outra máquina, sem o contexto da conversa.

## O que ficou pronto, commitado e testado no app

**Condomínios identificados por CPF**
A chave do cadastro passou a ser "documento": 14 dígitos (CNPJ) ou 11 (CPF do
síndico). 8 condomínios já cadastrados por CPF, vindos do export do Superlógica.

**Painel dos protocolos sem piscar**
Informar valor deixou de remontar o painel inteiro. Confirmado na janela.

**Anexo automático no Paybox**
Os protocolos saem com um bloco carimbado no pé da página (fornecedor, CNPJ do
fornecedor, vencimento, valor) que faz o Paybox anexar o documento à despesa
sozinho, acabando com a associação manual. Confirmado ponta a ponta contra o
Superlógica real. O porquê de cada decisão está na seção "Anexo automático no
Paybox" do `CLAUDE.md`.

**Tarifa, vencimento e chave perguntados no início do lote**
As três vão direto para a planilha de despesas. Ninguém precisa mais abrir o
arquivo no Excel antes de importar — foi abrindo para trocar a `chave` que a
coluna `vencimento` perdeu o formato de data e o Superlógica gravou 01/01/1970
nos lançamentos, sem acusar erro.

## O que falta

### 1. `git push`

Há commits locais que ainda não subiram. Conferir com `git status -sb`.

### 2. Verificar a aba de Cadastro na janela

Pendente desde a mudança do CPF. Não existe teste automatizado de interface no
projeto. Conferir:
- documento com 11 dígitos (ex: `529.982.247-25`) é aceito;
- `123` é recusado, com mensagem citando CNPJ **e** CPF;
- **selecionar um condomínio cadastrado por CPF, trocar o documento por um CNPJ
  e salvar → a tabela tem que continuar com UMA linha, não duas.** É a correção
  de duplicação; se aparecerem duas, ela falhou.

### 3. Gerar o executável

O `dist/` está desatualizado — nada disso chegou a quem usa o `.exe`. Duplo
clique em `gerar_exe.bat` (instala dependências, roda os testes e só empacota
se passarem).

## Pendências fora do código

- **Apagar os lançamentos de teste no Superlógica** (VILLARS, R$ 7,70, com e sem
  Pix). Um deles ficou com `chave_pix` preenchida com um BR Code de chave
  aleatória inexistente — não é pagável, mas não deve ficar lá.
- Pastas `C:\Users\Dell\Downloads\TESTE QR PAYBOX` e
  `C:\Users\Dell\Downloads\TESTE CODIGO BARRAS` são só de teste, podem ir fora.

## Coisas que já foram descartadas — não refazer

- **QR Code / código de barras para associação**: não funciona. Code128, I2of5 e
  Code39 nem são lidos; QR é lido mas não associa, nem com BR Code Pix válido.
  A associação é por valor + vencimento + fornecedor, em texto na página.
- **`etiqueta_paybox` como identificador livre**: exige URL do `sldocs.com.br`.
- **Colar link do sldocs manualmente**: destrói o ganho, é o mesmo trabalho
  manual de hoje.
- **Ler o número escrito à caneta** no protocolo do 11049: decisão de deixar
  como está, resolve pelo "Informar valor".
- **LIMA BARROS com CPF de 14 dígitos** (zeros à esquerda na máscara de CNPJ):
  decisão de ignorar. Ele não é identificado pelo conteúdo; cai no match por
  código/nome do arquivo.

## Armadilha conhecida

**Não abrir a planilha de despesas no Excel antes de importar.** Se precisar
abrir por algum motivo, conferir depois se a coluna `vencimento` continua como
data (e não como número tipo `46255`). Quando ela vira número cru, o
Superlógica grava 01/01/1970 e **não avisa** — o erro só aparece no lançamento.
