# Condomínios identificados por CPF — especificação

Data: 2026-08-24
Estado: aprovado, a implementar

## Objetivo

Permitir cadastrar e identificar condomínios cujo documento não é um CNPJ, mas
o **CPF** de uma pessoa física (tipicamente o síndico) — que é o que aparece
impresso no campo do tomador/pagador dos boletos e notas desses condomínios.

Hoje isso é impossível em dois pontos: `_adicionar_ou_atualizar`
(`identificacao_por_cnpj_6_0.py`) recusa qualquer coisa que não tenha 14
dígitos, e `carregar_cadastro` (`logica.py`) pula em silêncio a linha da
planilha cujo CNPJ não normaliza. O documento existe e está no papel; só não
cabe no cadastro.

## A decisão central: a chave do cadastro é o documento, não o CNPJ

O cadastro continua sendo um dict, mas a chave passa a ser um **documento
normalizado**: 14 dígitos (CNPJ) ou 11 (CPF).

Isso funciona sem ambiguidade porque **o comprimento distingue os dois
sozinho** — nenhum CPF pode ser confundido com um CNPJ, nem o contrário, em
nenhum ponto do programa. Cada um tem seu próprio dígito verificador, então
uma leitura errada de OCR continua sendo detectável nos dois casos. Não há
chave sintética, placeholder, nem linha que possa se sobrescrever.

### O que foi descartado, e por quê

A primeira ideia foi gravar `00.000.000/0000-00` para os condomínios sem CNPJ
e derivar uma chave interna a partir do código. Foi abandonada quando ficou
claro que **esses condomínios têm CPF e ele sai nos documentos**: com o CPF
real, a identificação automática volta a funcionar para eles, enquanto o
placeholder os deixaria dependendo para sempre do nome do arquivo.

Registrado porque o placeholder tem uma armadilha que pode ressurgir: com o
cadastro chaveado pelo documento, dois condomínios gravados como
`00.000.000/0000-00` **se sobrescrevem em silêncio** ao carregar — o segundo
apaga o primeiro sem nenhum aviso. Isso não é hipotético: o export de
condomínios ativos do Superlógica de 2026-08-20 trazia 8 CNPJs exatamente
assim. Qualquer retorno a essa ideia precisa resolver a colisão primeiro.

Também foi descartado rechavear o cadastro inteiro pelo **código**. É
conceitualmente atraente (os 765 códigos da planilha são únicos e nenhum está
vazio), mas obrigaria a mudar todo `cadastro.get(documento)` do caminho de
identificação — a parte mais crítica e mais testada do programa — sem nenhum
ganho visível. Guardar o documento como chave preserva esse caminho intacto.

## Cadastro

`cpf_valido` entra ao lado de `cnpj_valido` em `logica.py`: mesmo formato de
função, algoritmo próprio (pesos 10..2 e 11..2), rejeitando os 11 dígitos
iguais antes de conferir os dois DVs.

`formatar_cnpj` vira `formatar_documento`: 11 dígitos saem como
`000.000.000-00`, 14 como `00.000.000/0000-00`, e qualquer outra coisa
continua saindo crua, como hoje.

O cabeçalho da coluna A da planilha passa de `CNPJ` para `CNPJ / CPF`.
`carregar_cadastro` lê por posição e nunca por nome, então renomear não quebra
planilha nenhuma — inclusive as já existentes, que continuam carregando.

No formulário da aba de Cadastro, o campo aceita 11 **ou** 14 dígitos e recusa
o resto, com a mensagem citando os dois. Continua obrigatório: todo condomínio
tem um documento ou o outro.

### Correção incluída: editar o documento duplicava o registro

Hoje, mudar o CNPJ de um registro existente **cria um segundo** em vez de
atualizar o primeiro, porque o CNPJ é a chave do dict e a chave antiga
permanece. `_adicionar_ou_atualizar` passa a remover a chave antiga quando ela
muda.

Entra nesta mudança porque é exatamente o fluxo que ela cria: um condomínio
cadastrado por CPF que depois obtém CNPJ próprio. Sem a correção, ele passaria
a existir duas vezes.

## Achar o CPF no documento

Em dois níveis, com pesos deliberadamente diferentes.

### No campo semântico, sempre

Os padrões de `extrair_cnpj_tomador` (`CO-ESTIPULANTE`, `PAGADOR ... CNPJ/CPF`,
`TOMADOR`, `EMPREGADOR`, `CONTRATANTE/CLIENTE`) passam a aceitar um valor de 11
dígitos além do de 14. O rótulo `PAGADOR ... CNPJ/CPF:` já existe hoje — só o
valor esperado ali é que era de 14 dígitos, então o CPF já vinha sendo lido e
descartado.

Aqui o rótulo é a garantia: se o documento diz que aquele é o pagador, é ele.

### No fallback genérico, só se estiver cadastrado

Quando nenhum rótulo foi lido, um CPF solto no texto **só vira candidato se já
for um condomínio do cadastro**. CNPJ continua entrando como hoje.

A assimetria é proposital. Num boleto, um CNPJ solto tende a pertencer a
alguma empresa envolvida na cobrança; um CPF solto costuma ser de uma pessoa
qualquer — síndico, sacador avalista, quem assinou o documento. Aceitar todo
CPF válido da página encheria o processo de falso positivo, e falso positivo
aqui é documento carimbado com o código do condomínio errado.

O custo, explícito: um condomínio novo identificado por CPF **não** vira
pendente "documento não cadastrado" automaticamente a partir do fallback. Ele
cai para o caminho de código/nome do arquivo, que continua valendo como rede.
Quem for afrouxar essa regra depois precisa saber que o ganho é esse e o risco
é o carimbo errado.

### NFS-e

`extrair_dados_nfse` procura hoje apenas `CNPJ_REGEX` no bloco do tomador
(`logica.py`). Passa a procurar os dois formatos — senão a planilha de notas
desses condomínios sairia sistematicamente sem código.

## Carimbos

O carimbo de rodapé e o de canto superior nunca imprimiram o documento; não
mudam.

`montar_texto_protocolo_correio`, que hoje monta `código nome - documento`,
passa a omitir o documento quando ele é um CPF: sai `10005 VILLARS`, sem o
traço. O PDF carimbado circula e vai para o Superlógica, e estampar o CPF de
uma pessoa física nele é diferente de estampar o CNPJ de um condomínio. Com
CNPJ o carimbo continua idêntico ao de hoje.

## Testes

`tests/cadastro_teste.py` ganha um condomínio identificado por CPF, no mesmo
espírito do LAGO MAGGIORE sem `id_sl`: um caso real representado no cadastro
congelado, para que o resto da suíte o atravesse naturalmente.

Casos novos:

- `cpf_valido`: CPF válido, DV errado, 11 dígitos iguais, comprimento errado.
- `formatar_documento`: 11 e 14 dígitos, e entrada que não é nem um nem outro.
- `carregar_cadastro`/`salvar_cadastro`: ida e volta pela planilha preservando
  um CPF, e planilha antiga só com CNPJ continuando a carregar.
- `extrair_cnpj_tomador`: acha o CPF depois de um rótulo (`PAGADOR ... CNPJ/CPF`);
  no fallback genérico, **ignora** CPF fora do cadastro e **aceita** o que está
  nele; CNPJ não muda de comportamento em nenhum dos dois.
- `montar_texto_protocolo_correio`: com CNPJ imprime o documento, com CPF omite.
- Nenhum CPF casa com uma chave de CNPJ, e vice-versa.

## Fora de escopo

- Condomínio sem documento nenhum (nem CNPJ nem CPF). Foi descartado acima e
  não existe caminho para ele: o campo continua obrigatório.
- Extrair IBS/CBS ou qualquer outro campo novo da NFS-e.
- Rechavear o cadastro pelo código.
