# Identificação de PDFs por CNPJ — Contexto do Projeto

## O que o programa faz

App desktop (Windows) que processa boletos e notas fiscais (NFS-e) em PDF,
identifica a qual condomínio cada documento pertence (via CNPJ do tomador) e escreve
o código do condomínio no PDF (rodapé ou canto superior), comparando com um cadastro
em planilha `.xlsx`.

Arquivo principal: `identificacao_por_cnpj_6_0.py` — só a interface CustomTkinter
(classe `App`, redesign Swiss — ver seção "Redesign visual 6_0") e o `main`. Toda a
lógica de negócio (extração/validação de CNPJ, busca por nome, desempate, config
e predefinições, leitura de PDF/OCR, carimbo no PDF, persistência da planilha,
extração de dados das NFS-e) vive em `logica.py` — módulo sem nenhuma dependência de interface, importado com
`from logica import (...)` no topo do arquivo principal. É esse módulo que os
testes (`tests/`) importam diretamente (`import logica as app`). A versão anterior
`identificacao_por_cnpj_5_3.py` (Tkinter/ttk, monolítico, não modularizado) segue
no repo como referência. A lógica de identificação começou byte-idêntica entre 5_3
e 6_0 (o 6_0 nasceu como redesign exclusivo de interface), mas divergiu depois em
3 pontos, só no 6_0 — ver "Divergências de lógica só no 6_0 (pós-redesign)".
Cadastro de condomínios: `cadastro_condominios.xlsx` (colunas: CNPJ, Código,
Nome, **ID SL**). "ID SL" é o código do condomínio no **Superlógica** — outro
número, sem relação com o código interno (ex: KLOSTERS é `10004` aqui e `44`
lá). Nada da identificação nem dos carimbos usa esse campo — quem usa é a
geração da planilha de despesas do Superlógica (aba 3, v6.14.0), que preenche
com ele a coluna `condomínio` do arquivo de importação. Planilha antiga de 3 colunas continua carregando (campo
vazio) — mas cuidado: `salvar_cadastro` reescreve a planilha inteira, então
qualquer coluna nova precisa ser gravada lá também, senão a primeira edição
pela aba de Cadastro apaga a coluna de todos os condomínios em silêncio.
Preenchido a partir do export "condomínios ativos" do Superlógica via
`tests/_importar_id_sl.py` (uso único, reexecutável). O cruzamento é pelo
**código**, não pelo CNPJ: no export de 2026-08-20 vários CNPJs vinham
malformados (8 deles como `00.000.000/0000-00`), e cruzar por CNPJ casava só
632 dos 757 — contra 747 pelo código, sem nenhum conflito de nome. Os 10 sem
correspondência (não estão no export de ativos) ficam com o campo vazio.
Configuração da interface: `config.json` (ao lado do script, gitignored) — 3
predefinições de lote (FedCorp/F&F/Notas Diversas, ver seção "Predefinições de
lote") + tema; criado na 1ª execução com defaults.
Logs: `processamento.log` (histórico de sessões) e `erros.log` (exceções não
tratadas capturadas pelo handler global — ver seção "Robustez da interface"),
ambos ao lado do script.
Repositório: [gusmagalhaes01/codificador](https://github.com/gusmagalhaes01/codificador) (privado).

## Fluxo de identificação

1. Extrai texto do PDF via `pypdf` (texto nativo).
2. Se texto < 30 caracteres (PDF escaneado/imagem), tenta OCR via `winocr`
   (motor nativo do Windows — nada de programa externo instalado). Tenta `pt-BR`,
   cai para `en-US` se o idioma não estiver instalado.
3. Busca o CNPJ do tomador em `extrair_cnpj_tomador()`, em duas etapas:
   - **Campo semântico** (prioridade): procura rótulos como `CO-ESTIPULANTE`,
     `PAGADOR`, `TOMADOR`, `EMPREGADOR`, `CONTRATANTE/CLIENTE` seguidos do CNPJ.
   - **Fallback genérico**: varre todos os CNPJs do texto e exclui os conhecidos
     por não serem o condomínio (ver regra de negócio abaixo).
4. Casa o CNPJ encontrado com a planilha de cadastro e escreve o código no PDF.

## Regra de negócio importante — CNPJs que NUNCA são o condomínio tomador

Nos boletos da FedCorp existem 3 CNPJs em jogo, só um é o condomínio de verdade:

| CNPJ | Papel | É o tomador? |
|---|---|---|
| `35.315.360/0001-67` | FedCorp Administradora (emitente) | Não |
| `12.184.361/0001-14` | Imodata Empreendimentos (estipulante) | Não |
| *(varia por documento)* | Condomínio (co-estipulante) | **Sim** |

Esses dois primeiros estão hardcoded em `CNPJS_INTERMEDIARIOS` e são sempre
ignorados, tanto na busca por campo quanto no fallback genérico — isso é
necessário porque em alguns boletos escaneados o OCR não consegue ler o rótulo
`CO-ESTIPULANTE` (sai tudo concatenado sem quebra de linha), então sem essa
lista o fallback pegava o CNPJ da Imodata por engano.

**Se aparecer um novo intermediário conhecido** (outra empresa de cobrança/gestão
que não seja o condomínio), adicionar o CNPJ dele em `CNPJS_INTERMEDIARIOS`.

## CPF do síndico: por que a regra é diferente da do CNPJ

Condomínio sem CNPJ próprio é identificado pelo **CPF do síndico**, que sai
impresso no campo do tomador/pagador. A extração trata os dois em dois níveis,
com pesos deliberadamente diferentes:

- **Depois de um rótulo** (`PAGADOR ... CNPJ/CPF`, `TOMADOR`,
  `CO-ESTIPULANTE`...), o CPF é sempre aceito: o rótulo é a garantia de que
  aquele documento é o do pagador.
- **Na varredura genérica**, sem rótulo legível, o CPF **só vira candidato se
  já estiver no cadastro** (parâmetro opcional `cadastro` de
  `extrair_cnpj_tomador`). Num boleto, um CNPJ solto tende a ser de alguma
  empresa envolvida na cobrança, mas um CPF solto costuma ser de uma pessoa
  qualquer — síndico, sacador avalista, quem assinou. Aceitar todo CPF válido
  da página encheria o processo de falso positivo, e falso positivo aqui é PDF
  carimbado com o código do condomínio errado.

**Custo assumido:** um condomínio novo identificado por CPF **não** vira
pendente "não cadastrado" por esse caminho — ele cai no match por código/nome
do arquivo, que continua valendo como rede. Quem for afrouxar a regra depois
precisa saber que o ganho é esse e o risco é o carimbo errado.

`CPF_REGEX` e `CPF_FLEX` (`logica.py`) começam com o lookbehind `(?<!\d)` e
terminam com o lookahead `(?![\d\-/])`, que barra a continuação do número — de
propósito: sem eles, os 11 primeiros dígitos de um CNPJ colado sem separador
(OCR) poderiam casar como CPF, e uma coincidência de checksum viraria carimbo
errado.

**O CPF fica fora do carimbo do protocolo dos Correios**
(`montar_texto_protocolo_correio`): com CNPJ o carimbo é
`"{código} {nome} - {CNPJ}"`, com CPF sai só `"{código} {nome}"`. O PDF
carimbado circula e vai para o Superlógica, e estampar o CPF de uma pessoa
física nele é diferente de estampar o CNPJ de um condomínio.

## Cuidado: condomínios com nomes parecidos

A planilha tem pares de condomínios com nomes quase idênticos e CNPJs diferentes.
Exemplo real que já causou confusão:
- `05.695.194/0001-00` → código `10490` → "CENTRO COM CONDE DE BONFIM RES"
- `29.361.458/0001-58` → código `10491` → "CENTRO COM CONDE DE BONFIM"

Por isso a identificação **sempre** deve ser feita pelo CNPJ, nunca por
similaridade de nome sozinha.

## Limitações do ambiente de teste (importante para quem for testar fora do Windows)

- `winocr` só funciona no Windows (motor de OCR nativo). Em outro SO (ex: Linux),
  para simular o pipeline, usar `pymupdf` (renderizar página) + `pytesseract`
  (com o idioma `por` instalado via `apt-get install tesseract-ocr-por`).
  A lógica de regex/extração é a mesma — só troca o motor de OCR.
- PDFs de exemplo testados com sucesso (após o fix): `ARAUJO_LIMA_QUITADO_05_26.pdf`
  e `ARAUJO_LIMA_NF_05_26.pdf`, ambos resolvendo para CNPJ `40.338.774/0001-41`
  (código `10695`, ARAUJO LIMA).

## Histórico de decisões

- **v6.18.1 — a nota inteira era descartada porque "NFS-e" quebrava a linha
  no meio do rótulo**: nas notas da F&F de agosto/2026 (DANFSe v2.0), o
  rótulo do cabeçalho cai na borda da coluna e o hífen fica numa linha e o
  "e" na seguinte — `NÚMERO DA NFS-\ne\n \n12815`. São dois blocos de texto
  em alturas diferentes dentro do PDF, então **qualquer** extrator quebra ali;
  não é falha do `pypdf` nem nota escaneada, o texto nativo está íntegro.
  `campo_danfse` montava `re.escape(rotulo)` e exigia o rótulo inteiro numa
  linha, então `numero` vinha `None` e `extrair_dados_nfse` desistia na
  primeira checagem (linha `if not numero: return None`). O estrago era
  desproporcional à causa: a nota era descartada pelo **mesmo caminho** do
  "Detalhamento do Faturamento" — documento que de fato não deve ser lido —,
  e nada na tela distinguia os dois. `padrao_rotulo` (`logica.py`) passou a
  gerar o padrão do rótulo aceitando **uma** quebra dentro de "NFS-e" no fim
  do texto. Mesma ideia do prefixo `"EMITENTE DA NFS-"` em `SECOES_DANFSE`,
  com uma diferença: lá basta parar antes da letra que varia, aqui o "e"
  precisa ser **consumido**, porque o valor do campo vem depois dele. A folga
  é `[ \t]*\n?[ \t]*`, nunca `\s*` — com quebras livres, um rótulo de campo
  vazio casaria com a palavra iniciada em "e" da linha de baixo (ex:
  "emissão") e devolveria o valor errado, exatamente o motivo pelo qual
  `campo_danfse` já recusava `\s*` solto.
  **Por que os 668/668 da v6.11.0 não pegaram isto:** nas notas de julho, a
  mesma DANFSe v2.0 trazia `NÚMERO DA NFS-E` numa linha só (ver a fixture
  `nfse_v2_ibs.txt`). A quebra depende da largura da coluna no lote, não da
  versão do DANFSe — então validar contra um lote grande de uma competência
  só não cobre este eixo de variação.
  Três rótulos eram afetados, todos os que terminam em "NFS-e" no cabeçalho:
  Número (descartava a nota), Competência e Data de emissão (célula vazia).
  Validado contra os 2.220 PDFs reais de agosto/2026: 1.994 notas lidas sem
  nenhum campo vazio, 226 recusadas — **todas** "Detalhamento do Faturamento",
  que é o comportamento correto. Em 1.994 notas o código deduzido do CNPJ do
  tomador bateu com o código no nome do arquivo em 100% dos casos.
  Fixture `tests/dados/nfse_v2_rotulo_quebrado.txt`, testes em
  `tests/test_extracao_nfse.py` (`TestRotuloQuebradoNoHifen`,
  `TestPadraoRotulo`).

- **O valor saiu do topo direito: com o bloco do Paybox ele aparecia duas
  vezes**: desde a v6.16.0 o bloco do Paybox já traz `VALOR: R$ ...`, então
  o carimbo do topo virou repetição na mesma folha. Some pelo mesmo motivo
  que a conta ("12 un × R$ 3,85 = ...") saiu na v6.13.1 — **duas ocorrências
  de "R$" na página dão ao OCR do Superlógica a chance de capturar a
  errada**. Agora o carimbo do topo só entra quando o bloco do Paybox NÃO
  vai ser gravado (lote sem vencimento informado), para o papel nunca ficar
  sem valor nenhum. Conferido em protocolo de página única: uma única
  ocorrência de "R$" no PDF carimbado.

- **Botão "Tela cheia" nos painéis de resultado**: os painéis são
  `CTkToplevel` com `transient(self)`, e no Windows isso **tira os botões de
  minimizar/maximizar da barra de título**, deixando só o fechar — não havia
  como aproveitar a tela inteira para ver planilha e documento lado a lado. O
  alternador vive em `_montar_barra_janela`, chamado por
  `_montar_painel_resultado`, então as abas 1 e 3 ganham juntas. Usa
  `state("zoomed")` (maximizar) e **não** `-fullscreen`: o fullscreen do Tk
  esconde a barra de título junto, e sem ela não sobra nem o botão de fechar.
  F11 alterna; Esc só restaura — Esc não pode fechar, há trabalho não salvo
  no painel. Rótulo em texto puro, sem glifo: `⛶` (U+26F6) não é garantido
  nas fontes usadas.

- **Dar o valor pela CÉLULA não movia a linha de pendente para calculado**:
  existiam dois caminhos para informar o valor — o botão "Informar valor", que
  movia o registro entre as listas do `resultado`, e a edição da célula na
  grade, que gravava o valor e deixava o registro em `pendentes`. No segundo
  caso a linha continuava marcada como pendente e **travava a geração da
  planilha de despesas**, sem nada na tela explicando. `reclassificar_registro`
  (`logica.py`) passou a ser o único critério: **com valor é `processados`, sem
  valor é `pendentes`**, venha o valor de onde vier. Testes em
  `tests/test_reclassificar.py`.
- **Divisor arrastável entre a planilha e o documento**: `ttk.PanedWindow`
  horizontal, como o painel de visualização do Explorer — puxar para a
  esquerda estreita a grade e a página cresce junto, sem mexer no zoom. A
  largura da página deixou de ser constante e passa a acompanhar o painel
  (`_largura_previa`). O redesenho é **adiado** (`ESPERA_REDESENHO_PREVIA`):
  arrastar dispara muitos `<Configure>` seguidos e renderizar em todos
  travaria a interface; enquanto espera, o canvas segue mostrando a imagem
  anterior, então não pisca.

- **Editar o código na grade exige ID SL, como o botão já exigia**: o botão
  "Escolher condomínio" recusa condomínio sem o campo, mas a edição da célula
  aceitava — a linha continuava pendente e **nada na tela explicava por quê**.
  São 10 dos 772 condomínios, então acontecia de verdade. `validar_edicao_
  protocolo` passou a conferir o ID SL e devolve a mesma frase do botão.
- **Botão "Remover linha" no painel dos protocolos**: saída para o arquivo que
  não tem como ser resolvido (PDF corrompido ou de 0 byte, documento que nem
  devia estar na pasta) e que, sem isso, ficaria pendente para sempre travando
  a geração das despesas. **Não apaga o PDF do disco** — some da planilha e do
  painel, o arquivo continua onde estava, e a confirmação diz isso.
  `remover_linha_do_lote` (`logica.py`) **reindexa** `indice_linha` de todo
  registro depois do removido: é posição em `ctx["linhas"]`, então sem
  reindexar o registro seguinte passaria a apontar para a linha errada e uma
  edição posterior escreveria no vizinho. O total em reais também perde o
  valor da linha que saiu, senão o cartão TOTAL divergiria da planilha.
  Botão contornado, não cobalto: remover não RESOLVE a pendência, só a tira
  do caminho. Testes em `tests/test_remover_linha.py`.

- **Correção — escolher o condomínio antes de informar o valor**: o carimbo
  calculava o valor incondicionalmente, então anotar o condomínio de um
  pendente ainda sem contagem caía em `valor_protocolo(None, tarifa)` e
  estourava `int() argument must be ... not 'NoneType'`. O PDF não era
  recarimbado e o papel ficava **sem identificação nenhuma** — o oposto do
  que a ação pretendia. Agora `valor_do_carimbo` (`logica.py`) devolve `None`
  quando o valor é desconhecido, e nesse caso carimba-se só a identificação
  lateral: o valor e o bloco do Paybox entram depois, quando for informado.
  Sem cadastro **e** sem valor não há o que escrever, e aí a gravação é
  recusada com mensagem — gravar cópia sem carimbo nenhum faria o arquivo de
  saída parecer processado. Testes em `tests/test_valor_do_carimbo.py`.

- **Grade editável com prévia do documento (aba 3)**: o painel de resultado
  deixou de ter três tabelas (pendentes/calculados/ignorados) e passou a ser
  **uma grade que É a planilha**, com a página do documento ao lado. A
  distinção virou cor da linha + filtro. Toda edição que muda o que está
  IMPRESSO no papel (código, unidades, valor) **dispara o recarimbo**, e se
  ele falhar a edição não é aplicada — sem isso, papel e planilha divergiriam
  em silêncio e o Paybox deixaria de anexar o documento. "Observação" é a
  única editável que não mexe no papel. Ver o spec
  `docs/superpowers/specs/2026-08-29-planilha-embutida-previa-design.md`.
- **A prévia mostra o documento INTEIRO, não só a 1ª página**: protocolo dos
  Correios costuma ter 2 ou 3 páginas, e a lista de unidades continua na
  segunda — mostrar só a primeira escondia justamente o que se quer conferir.
  `renderizar_paginas_pdf` (`logica.py`) devolve todas as páginas até
  `LIMITE_PAGINAS_PREVIA` (12) e informa o total real, para o cabeçalho poder
  avisar quando há página não mostrada. As páginas ficam empilhadas no canvas,
  com um vão entre elas.
- **Zoom pela roda do mouse e arraste para mover**: a roda amplia direto, sem
  tecla nenhuma, e a navegação pelo documento é por arraste
  (`scan_mark`/`scan_dragto` com `gain=1` — o padrão, 10, faz a página
  disparar e passar do ponto). O zoom preserva o ponto que está sob o cursor,
  em vez de voltar ao topo. Ampliar re-renderiza do PDF, não amplia bitmap,
  então a 400% o texto continua nítido.
- **A prévia usa `tk.Canvas` + `ImageTk.PhotoImage`, nunca `CTkLabel` +
  `CTkImage`**: além de o canvas rolar nos dois eixos (o que permite o zoom
  até 4x), `CTkLabel.configure(image=None)` **não limpa a imagem** — o
  `_update_image` do CustomTkinter só age quando `_image` é `CTkImage` ou
  não-`None`, então o label interno seguia apontando para um `pyimageN` cuja
  `CTkImage` já tinha sido coletada, dando `image "pyimage13" doesn't exist`.
  A ordem em `_desenhar_previa` também importa: limpar o canvas ANTES de
  soltar a referência da imagem anterior.
- **Dois campos que faltavam no contexto do lote**: `_ctx_protocolos` nunca
  guardou a `tarifa` (embora `_acao_escolher_condominio` já lesse
  `ctx.get("tarifa")`, sempre recebendo `None`), e os `ignorados` do painel
  só tinham `arquivo` e `motivo`. Sem a tarifa não há como recalcular o valor
  ao editar as unidades; sem `caminho`/`indice_linha` nos ignorados a prévia
  não funcionaria naquelas linhas.

- **v6.17.1 — correção: resolver o valor e depois escolher o condomínio**: o registro
  criado por `_acao_informar_valor` não carregava `caminho` nem
  `indice_linha`, e o item voltava para a tabela de pendentes (por desenho:
  `_listas_do_painel_protocolos` mostra junto dos pendentes todo processado
  ainda sem ID SL). "Escolher condomínio" então estourava
  `KeyError('caminho')` no recarimbo e, logo depois,
  `KeyError('indice_linha')` **fora** do try/except — este último caindo no
  handler global. A montagem do registro virou `registro_painel_resolvido`
  (`logica.py`), que carrega adiante `caminho`, `indice_linha` e
  `pasta_destino` — este último para o recarimbo cair no MESMO "Lote NN" em
  vez de abrir outro. Testes em `tests/test_resolucao_manual_painel.py`.
  A lição que fica: um dict que transita entre as duas listas do painel
  precisa dos campos das ações das DUAS, não só das da lista onde nasceu.

- **v6.17.0 — tarifa, vencimento e chave perguntados no início do lote**: a
  `chave` da importação passou a ser perguntada junto da tarifa e do
  vencimento, e vai direto para a planilha (`gerar_planilha_despesas`,
  parâmetro `chave`). Motivo concreto: ela muda a cada importação e era
  editada à mão na planilha já gerada — e foi abrindo o arquivo no Excel só
  para isso que a coluna `vencimento` perdeu o formato de data e voltou a ser
  número cru, fazendo o Superlógica gravar 01/01/1970 nos lançamentos **sem
  acusar erro nenhum**. Com os três dados vindo do programa, não há mais razão
  para abrir a planilha antes de importar. `chave=None` mantém o que estiver
  no modelo.

- **v6.16.0 — anexo automático no Paybox (bloco carimbado)**: os protocolos
  dos Correios passam a sair com um bloco de texto — fornecedor, CNPJ do
  fornecedor, vencimento e valor (`montar_bloco_paybox`, `logica.py`) — que faz
  o Paybox anexar o documento à despesa **sozinho**, eliminando a associação
  manual. O vencimento passou a ser perguntado no **início** do lote, junto da
  tarifa, porque o carimbo acontece antes da planilha existir; `_acao_gerar_despesas`
  reusa esse mesmo valor em vez de perguntar de novo. Ver "Anexo automático no
  Paybox" abaixo.

- **v6.15.0 — condomínios identificados por CPF**: alguns condomínios não têm
  CNPJ próprio e aparecem nos boletos e notas com o **CPF do síndico**. A chave
  do cadastro deixou de ser "CNPJ" e passou a ser **documento normalizado**: 14
  dígitos (CNPJ) ou 11 (CPF). Funciona sem ambiguidade porque o comprimento
  distingue os dois sozinho, e cada um tem seu próprio dígito verificador
  (`cnpj_valido` e `cpf_valido`, `logica.py`). `formatar_cnpj` virou
  `formatar_documento`, formatando pelo comprimento. A coluna A da planilha
  passou a se chamar `CNPJ / CPF`; `carregar_cadastro` lê por posição e nunca
  pelo nome do cabeçalho, então planilha antiga continua abrindo. Junto veio
  uma correção: editar o documento de um registro existente **criava um
  segundo** em vez de atualizar o primeiro, porque o documento é a chave — é
  exatamente o fluxo de quem estava cadastrado por CPF e depois obteve CNPJ.
  Ver o spec `docs/superpowers/specs/2026-08-24-cadastro-por-cpf-design.md`.
- **Fix (versão 5_3)**: adicionado `CNPJS_INTERMEDIARIOS` para excluir FedCorp e
  Imodata também no fallback genérico, não só no filtro por campo semântico.
  Corrigiu caso onde o OCR não lia o rótulo `CO-ESTIPULANTE` e o script pegava
  o CNPJ da Imodata por engano.
- **Versão 6_0 (redesign)**: reforma gráfica completa em `identificacao_por_cnpj_6_0.py`
  — reestruturação em 3 telas (Principal / Configurações / Resultado) **e** nova
  identidade visual Swiss International Style, migrando de Tkinter/ttk para
  **CustomTkinter**. Executada como redesign **exclusivamente de interface**: nenhuma
  função de lógica de identificação foi tocada (as 14+ funções são byte-idênticas ao
  5_3) e o formato de `processamento.log`/`erros.log`/planilha foi preservado. Ver
  "Redesign visual 6_0".
- **v6.7.0 — extração das NFS-e para planilha (aba 2)**: nova aba que lê as
  notas de uma pasta e gera um `.xlsx` com os dados de cada uma (identificação,
  tomador, código, valores, ISSQN e retenções federais). Lógica em `logica.py`
  (`extrair_dados_nfse` e companhia), interface em `_montar_aba_extracao`,
  testes em `tests/test_extracao_nfse.py`. Decisão central: **não usar OCR
  nessa funcionalidade** — ver seção "Extrair dados das NFS-e para planilha"
  para o raciocínio e os números da validação. Nada da codificação de PDFs foi
  alterado; as abas de Cadastro e Logs só foram renumeradas (3 e 4).
- **v6.6.0 — sugestão por nome nos pendentes "Não foi possível ler"**: quando a
  extração de CNPJ do conteúdo não acha nenhum candidato (escaneado sem CNPJ
  legível), `candidatos_por_nome()` (`logica.py`) tenta sugerir condomínios por
  similaridade de nome (arquivo + texto OCR) e o pendente vira "Nome parecido
  encontrado", habilitando o botão "Escolher" que antes ficava indisponível
  (só "Abrir PDF"). Ao contrário de `buscar_por_nome_arquivo`, não aplica o
  corte de ambiguidade que bloqueia decisão automática — aqui a escolha final é
  sempre manual, então mostrar os candidatos parecidos (ex: os dois "CONDE DE
  BONFIM" juntos) é o objetivo, não um bug. Mantém a regra de nunca identificar
  **automaticamente** só por nome — só o funcionário decide, depois de abrir o
  PDF. Ver "Cuidado: condomínios com nomes parecidos".
- **v6.8.0 — Protocolo de Recebimento de Documento (Correios/Imodata)**: novo
  tipo de documento reconhecido, além de boleto/NFS-e — recibo de entrega escaneado,
  sem CNPJ nenhum, mas com o código do condomínio já pronto no texto (ex:
  `W700A VILLARS (10005) Protocolo de Recebimento de Documento...`).
  `extrair_codigo_protocolo_correio()` (`logica.py`) detecta o marcador e
  extrai o código automaticamente pelo conteúdo (não é uma predefinição —
  funciona misturado com boletos normais no mesmo lote); resolve o CNPJ via
  o mesmo índice reverso código→CNPJ de `buscar_por_nome_arquivo`. Código não
  cadastrado vira pendente "Código não cadastrado" (`cnpj: None`, já que o
  documento não tem CNPJ pra oferecer). Validado com 4 arquivos reais contra
  a planilha real antes de liberar, mesma disciplina que pegou o bug do
  `candidatos_por_nome`. Ver spec
  `docs/superpowers/specs/2026-08-13-protocolo-correio-design.md`.
- **v6.9.0 — carimbo lateral rotacionado no protocolo dos Correios**: os
  documentos identificados via `extrair_codigo_protocolo_correio` passam a
  receber um carimbo diferente do rodapé/canto superior normais — uma linha
  única (`"{código} {nome} - {CNPJ}"`, via `montar_texto_protocolo_correio`)
  rotacionada 90° (lê de baixo pra cima), colada na margem direita,
  verticalmente centralizada. Automático só pra esse tipo de documento, não
  depende de `modo_texto`. `criar_overlay()` ganhou o parâmetro opcional
  `angulo` (padrão `0`, comportamento antigo preservado) pra isso. Testado
  com um PDF real (protocolo do VILLARS) e confirmado que a IA do
  Superlógica conseguiu reconhecer o código carimbado. Ver spec
  `docs/superpowers/specs/2026-08-13-carimbo-lateral-protocolo-design.md`.
- **v6.10.0 — contagem e cobrança dos Protocolos dos Correios (aba 3)**: aba
  nova que conta as unidades entregues em cada protocolo, multiplica pela
  tarifa informada no lote (perguntada a cada processamento — ela muda com o
  tempo: nos protocolos de referência aparecem 3,45 e 3,85), carimba o valor
  no canto superior direito e gera uma planilha com o total. Ao contrário da
  aba 2, **aqui o OCR é usado** — o que torna isso aceitável é o
  `Listando N unidades` impresso no documento, que funciona como dígito
  verificador da contagem, com a contagem de linhas e a de "Correio" servindo
  de conferência. Divergiu, relê em qualidade maior e, persistindo, vira
  pendente sem carimbo. Abas de Cadastro e Logs renumeradas para 4 e 5. Ver
  spec `docs/superpowers/specs/2026-08-17-contagem-protocolo-correio-design.md`.
- **v6.11.0 — suporte à DANFSe v2.0 na extração de NFS-e**: a Prefeitura do
  Rio passou a emitir DANFSe v2.0 (reforma tributária, campos novos de
  IBS/CBS), convivendo no mesmo lote com notas v1.0 — não foi uma troca
  limpa (ex: às 10h e às 11h do mesmo dia, formatos diferentes). Isso
  quebrava `extrair_dados_nfse` por completo: os rótulos de identificação e
  os títulos de seção saem em CAIXA ALTA na v2.0 (vs Título Normal na v1.0),
  a seção do tomador foi renomeada de "TOMADOR DO SERVIÇO" para "TOMADOR /
  ADQUIRENTE", e o campo "Valor do Serviço" virou "Valor da Operação /
  Serviço". `bloco_secao`/`campo_danfse` (`logica.py`) continuam sensíveis a
  maiúsculas de propósito — uma primeira tentativa de busca "sem diferenciar
  maiúsculas" quebrou as notas v1.0, porque colide com "Código de Tributação
  Municipal" (um rótulo de campo que contém as mesmas palavras da seção
  real, só que em Título Normal); em vez disso, os poucos rótulos que
  realmente mudam de caixa entre versões (ex: "EMITENTE DA NFS-") usam uma
  forma sem a letra final que varia. **IBS/CBS não são extraídos** — os
  campos novos da reforma tributária, decisão deliberada por ora. Validado
  contra os 668 PDFs reais de um lote misto v1.0/v2.0: 668/668 reconhecidas.
- **v6.12.0 — rotação do log e separação da saída em lotes**: duas coisas
  independentes. (1) `rotacionar_log()` (`logica.py`) descarta as sessões
  mais antigas do `processamento.log` quando ele passa de 5 MB
  (`LIMITE_TAMANHO_LOG`) — antes crescia pra sempre e deixava a aba de Logs
  lenta. Sessões são os blocos separados por linha em branco que
  `_salvar_sessao_no_log` já gravava; a mais recente nunca é apagada, mesmo
  que sozinha passe do limite. (2) Checkbox "Separar a saída em lotes de N
  arquivos" logo abaixo do botão primário das abas 1 e 3 (as duas que geram
  PDF), porque o Superlógica só aceita um punhado de arquivos por envio: a
  saída sai em subpastas `Lote 01`, `Lote 02`... via `caminho_do_lote()`.
  O índice usado é o de arquivos **efetivamente carimbados**, não o do laço
  — assim pendentes não ocupam vaga e cada lote sai cheio. A preferência
  fica no nível raiz do `config.json` (não dentro da predefinição): depende
  do sistema de destino, não do tipo de documento. Desligada por padrão.
- **v6.14.0 — planilha de Despesas do Superlógica (aba 3)**: botão "Gerar
  planilha do Superlógica" no painel de resultado, que transforma o lote em
  arquivo de importação de despesas — uma linha por protocolo cobrável, com
  `condomínio` preenchido pelo **ID SL** do cadastro e `valor` pelo valor
  apurado. O modelo vem do arquivo do usuário (baixado do próprio Superlógica)
  e é **copiado, não reconstruído**: preserva formatos, validações e colunas
  ocultas que o importador pode exigir, e sobrevive a mudanças de layout que
  não são nossas. As colunas `condomínio` e `valor` são achadas pelo nome
  normalizado, nunca pela posição. A linha 2 do modelo é o molde, com os
  campos que se repetem (fornecedor, categoria, forma de pagamento), e é
  substituída pela primeira linha real. **Mesmo condomínio em dois protocolos
  gera duas linhas**, de propósito: cada lançamento continua rastreável até o
  papel que o originou. O **vencimento é perguntado a cada geração**, como a
  tarifa: quando ficava no modelo, a data virava número de série em célula
  `General` e o Superlógica gravava 01/01/1970 e recusava os lançamentos. Ver
  "Planilha de Despesas do Superlógica" abaixo e o spec
  `docs/superpowers/specs/2026-08-19-despesas-superlogica-design.md`.
- **v6.13.0 — painel de resultado na aba 3, com valor em reais digitado à
  mão**: a aba dos Protocolos dos Correios passa a encerrar com painel
  (Protocolos / Pendentes / Total), no mesmo molde da aba 1 e reaproveitando
  os mesmos três helpers de UI (`_montar_painel_resultado`,
  `_montar_faixa_cartoes`, `_montar_tabela_resultado`,
  `identificacao_por_cnpj_6_0.py`). A pendência criada quando a conferência
  não aceita a contagem (ver "Regra de aceite" acima) se resolve pelo botão
  "Informar valor" — em reais, não em unidades, porque é o número que a
  pessoa já escreveu à caneta na folha do protocolo. A linha resolvida assim
  sai da planilha (`linha_planilha_protocolo`, parâmetro `valor_manual`) com
  Unidades e Tarifa vazias e observação `Valor informado manualmente`; o
  carimbo no PDF mostra só o valor, sem a conta, então o papel sozinho não
  distingue valor calculado de valor digitado — só a planilha guarda essa
  diferença. Motivado por um caso real: no protocolo do `11049 APART HOTEL`
  o `Listando 76 unidades` impresso foi riscado à caneta e trocado por `78`,
  e a contagem de "Correio" concordava com o número riscado (76) — nenhum
  conferidor automático pega esse tipo de rasura. Validado contra o lote
  real de 4 protocolos multipágina (`C:\Users\Dell\Downloads\TESTE
  CORREIO\Nova pasta`, fora do repo): 3 calculados (36, 31 e 60 unidades,
  R$ 138,60 + R$ 119,35 + R$ 231,00 = R$ 488,95 à tarifa de R$ 3,85) e 1
  pendente, o do 11049, exatamente como esperado.

## Melhorias implementadas (a partir da versão 5_3)

1. **Match por nome de arquivo primeiro** (`buscar_por_nome_arquivo`): antes de
   abrir o PDF, tenta casar o nome do arquivo (normalizado — sem acento, sem
   dígitos, maiúsculo) contra a coluna "Nome" do cadastro via
   `difflib.SequenceMatcher`. Só aceita o match se o score for ≥ 0.72
   (`LIMIAR_SCORE_NOME`) **e** o 2º colocado ficar pelo menos 0.08
   (`LIMIAR_DIFERENCA_AMBIGUA`) mais distante — caso contrário considera
   ambíguo (ex: os dois "CONDE DE BONFIM") e cai para a extração de CNPJ do
   conteúdo do PDF. Controlado pelo checkbox "Tentar identificar pelo nome do
   arquivo..." na aba 1 (ligado por padrão).
2. **Validação de dígitos verificadores do CNPJ** (`cnpj_valido`): todo CNPJ
   candidato agora passa pelo checksum oficial antes de ser aceito, tanto no
   match por campo semântico quanto no fallback genérico em
   `extrair_cnpj_tomador`. Quando o OCR foi usado e nenhum CNPJ válido restou
   (provável erro de leitura), o processamento re-tenta automaticamente em um
   DPI mais alto (`proximo_dpi_maior`, escada 72→150→200→300→400), uma única
   vez por arquivo.
3. **OCR por região** (`extrair_texto_ocr_regiao`): opcional, configurável na
   aba 1 ("OCR por região"). Permite recortar uma área fixa da 1ª página
   (frações 0.0–1.0 de x0,y0,x1,y1) e rodar OCR só nela antes do OCR de
   página inteira — mais rápido e menos sujeito a carimbos como "QUITADO"
   fora da área de interesse. Só é usado quando resulta em exatamente 1
   candidato; caso contrário cai para o OCR de página inteira normalmente.
   Calibração é visual: botão "🖱 Selecionar região no PDF..."
   (`_selecionar_regiao_visualmente` / `_montar_janela_selecao_regiao`) abre
   um PDF de exemplo em tamanho real (com barra de rolagem, já que a página
   não cabe inteira na tela) e o usuário desenha o retângulo clicando e
   arrastando o mouse; os campos x0,y0,x1,y1 são preenchidos sozinhos.
   **Os valores padrão (y0=0.15, y1=0.35) continuam sendo um chute inicial,
   não calibrados contra um boleto real da FedCorp** — calibrar antes de usar.
   Desligado por padrão até o usuário calibrar.
4. **Tipo de serviço com quebra de linha**: o campo (aba 1, item 4) virou uma
   caixa de texto multi-linha (`self.txt_tipo_servico`, um `tk.Text`) em vez
   de um `Entry` de uma linha só — dá pra apertar Enter e quebrar o texto.
   `criar_overlay()` agora desenha cada linha do texto empilhada (a primeira
   no topo, as seguintes abaixo, respeitando centralização), em vez de forçar
   tudo numa linha só no rodapé do PDF. O log de progresso mostra as quebras
   trocadas por " / " pra não virar múltiplas linhas na tela de log.

## Robustez da interface (janela pequena / crashes)

- A aba "1. Processamento" cresceu bastante com os itens de OCR por região e
  passou a não caber em telas menores — o conteúdo da aba agora fica dentro
  de um canvas rolável (scroll com a roda do mouse ou barra lateral), então o
  botão "▶ Processar PDFs" nunca fica inacessível. A janela principal também
  passou a ser redimensionável (`resizable(True, True)`, com `minsize`).
- `App.report_callback_exception` foi adicionado: qualquer exceção não
  tratada num callback (clique de botão etc.) antes caía no comportamento
  padrão do Tkinter de tentar imprimir no stderr — em ambientes sem console
  (ex: script associado ao `pythonw`), isso podia derrubar a janela inteira
  sem aviso nenhum. Agora o erro é gravado em `erros.log` (ao lado do script)
  e mostrado num popup, sem fechar o app.

## Redesign visual 6_0 (implementado — `identificacao_por_cnpj_6_0.py`)

Reforma gráfica que separa **operação** (o que o funcionário faz todo dia) de
**configuração** (o que se define uma vez) e transforma o resultado num painel
acionável. Migração de Tkinter/ttk para **CustomTkinter** (o Notebook de abas
foi mantido e estilizado; tabelas continuam em `ttk.Treeview` estilizado, pois
o CustomTkinter não tem widget de tabela).

**3 telas:**
- **Principal** (aba 1, sem scroll): cabeçalho "Codificador" + botão
  "⚙ Configurações" + alternador de tema; campos de pasta de entrada/saída com
  botão "Trocar"; botão cobalto largo "Processar N PDFs" (N = contagem de `*.pdf`,
  dinâmico; desabilitado sem pasta); linha explicativa em linguagem humana. Sem
  jargão técnico.
- **Configurações** (`_abrir_configuracoes`, modal `CTkToplevel` com
  `CTkScrollableFrame`): Empresa (CNPJ emitente, validado por checksum ao salvar),
  Identificação (checkbox nome do arquivo), Leitura de boletos escaneados
  (checkbox + "Qualidade de leitura" Rápida 72 / Normal 200 / Máxima 300 →
  mapeia o antigo DPI; grupo recolhível "Leitura por região" com o seletor visual
  `_selecionar_regiao_visualmente`), Texto no PDF (posição, tipo de serviço
  multi-linha, tamanho, cor validada `#RRGGBB`, e opção "Renomear o arquivo de
  saída com o código do condomínio" → `nome_saida_com_codigo`, ex: "10002 - ARAUJO
  LIMA QUITADO.pdf"). Padrão snapshot-ao-abrir /
  restaura-no-Cancelar / persiste-no-Salvar. **Nenhum rótulo usa "OCR" ou "DPI"** —
  vocabulário humano ("boletos escaneados", "qualidade de leitura").
- **Resultado** (`mostrar_resultado(resultado)`, painel `CTkToplevel`): substitui
  o antigo `messagebox` + log. 3 cartões (Processados / Pendentes / Tempo, só o
  Pendentes em cobalto); tabela de **pendentes primeiro** (colunas Arquivo | Motivo,
  motivos humanos: "CNPJ não cadastrado" / "Não foi possível ler" / "Dois CNPJs
  possíveis") com ações por linha (duplo clique **ou** botão): **Cadastrar** (abre a
  aba de cadastro pré-preenchida com CNPJ + nome sugerido), **Abrir PDF**
  (`os.startfile`), **Escolher** (popup de candidatos → reprocessa só aquele
  arquivo via `_reprocessar_arquivo`, usando `self._ctx_processamento`); depois a
  tabela de processados (Arquivo | Código | Origem humana). O `processamento.log`
  técnico continua idêntico, gravado como antes.

**Identidade visual (Swiss International Style):** cantos retos
(`corner_radius=0` em tudo), **cor de destaque única cobalto** (`TEMA_*["acento"]`;
sem verde/amarelo/vermelho de status — pendências usam o próprio cobalto), hierarquia
por tom sólido (não por opacidade, que o CustomTkinter não suporta em texto),
tipografia grotesca em peso leve (IBM Plex Sans se instalada, senão Segoe UI, via
`familia_fonte()`), grid de 8px, hairlines de 1px. Paleta em dois dicts de módulo
`TEMA_CLARO`/`TEMA_ESCURO` (stone), trocados por `aplicar_tema`/`_recolorir`; sem
hex soltos no código. Tema padrão segue o Windows (`darkdetect`), com alternador
manual persistido em `config.json`.

**Regra de ouro do cobalto:** só no botão "Processar", no cartão/label "Pendentes",
e nos botões que **resolvem** um pendente (Cadastrar/Escolher/"Usar este CNPJ").
Botões neutros (Trocar, Abrir PDF, Cancelar) são contornados, fundo transparente.

## Predefinições de lote (FedCorp / F&F / Notas Diversas)

A Tela Principal tem 3 botões ("PREDEFINIÇÃO") que trocam de uma vez todas as
configurações de um tipo de lote — CNPJ emitente, identificação, carimbo, tipo
de serviço — em vez de reconfigurar tudo toda sessão. Editáveis pelo próprio
app (não fixas no código): cada perfil é um dict dentro de
`config.json["predefinicoes"]`, com `predefinicao_ativa` apontando qual está
em uso. `carregar_config()` migra automaticamente um `config.json` de antes
desse recurso (campos soltos no nível raiz) para dentro do perfil "fedcorp",
sem perder a configuração que já existia.

- **FedCorp** (`cnpj_emitente=35.315.360/0001-67`, rodapé, `CIPAA`): fluxo
  original — `usar_match_nome_arquivo=True` tenta primeiro
  `buscar_por_nome_arquivo` (ver abaixo), senão cai para
  texto/OCR/`extrair_cnpj_tomador` do conteúdo do PDF.
- **F&F** (`cnpj_emitente=13.736.666/0001-54`, rodapé, tipo de serviço
  editável por lote — PGR/PCMSO/ESOCIAL/EXAMES/TREINAMENTOS, o usuário troca
  na aba/pasta que for processar): os arquivos já chegam nomeados com o
  código do condomínio embutido (ex: `PGR 10004 Klosters.pdf` → 10004) — quase
  sempre resolve na primeira etapa (código exato), sem precisar abrir o PDF.
- **Notas Diversas** (`cnpj_emitente` vazio — lotes de administradoras
  variadas, ex: Imodata; canto superior, só o código): usa o pipeline de CNPJ
  normal, mas sem emitente fixo pra excluir a priori. Tem
  `preferir_cadastrado_em_ambiguo=True`: quando sobra mais de um CNPJ
  candidato (tipicamente o emitente da nota, não cadastrado, + o condomínio
  tomador, cadastrado), prefere o(s) que já estão no cadastro — resolve o
  caso comum sem precisar de emitente fixo.

**`buscar_por_nome_arquivo` faz DUAS coisas** (a que resolver primeiro,
ganha — não combina): (1) procura números no nome do arquivo que batam
*exatamente* com um código do cadastro (usado pela F&F, mas funciona pra
qualquer perfil com `usar_match_nome_arquivo=True`); (2) se não achar código,
cai pro fuzzy match original por nome do condomínio (usado pela FedCorp).
Não é um "modo" separado por perfil — é o mesmo mecanismo pra todos, e cada
perfil só liga/desliga `usar_match_nome_arquivo`.

Configurações → "Salvar" grava no perfil ativo (`Configurações — Editando:
<nome>` no título do modal), não mais num config plano. Snapshot/Cancelar
continuam do jeito que eram. Os botões de predefinição mostram
"(modificado)" quando as `tk.Var` vivas divergem do que está salvo no perfil
(`_perfil_ativo_modificado`) — só indicativo, não bloqueia nada.

## Extrair dados das NFS-e para planilha (aba 2, v6.7.0)

Contrapartida do carimbo: em vez de **escrever** o código no PDF, **lê** os dados
das notas e gera um `.xlsx`. Nasceu do fluxo F&F — as notas daquele lote são
DANFSe da prefeitura do Rio, com texto nativo e rotulado.

**Interface** (`_montar_aba_extracao`): pasta com as notas + destino da planilha
(sugerido automaticamente ao escolher a pasta) + botão primário cobalto
"Extrair dados de N notas" (mesmo papel do "Processar" na tela Principal — é a
extensão natural da regra do cobalto: o botão primário da tela). Só a pasta
escolhida, sem subpastas. Não altera os PDFs.

**Por que aqui NÃO se usa OCR (decisão deliberada):** CNPJ tem dígito
verificador, então uma leitura errada é detectável — é isso que sustenta a
escada de DPI em `extrair_cnpj_tomador`. Valor e data não têm nada disso:
`1.234,56` lido como `1.234,58` entraria numa planilha financeira sem ninguém
perceber. Nota sem texto nativo é reportada como "Nota escaneada — não foi
possível ler", nunca chutada.

**Como a leitura é feita** (`logica.py`):
- `bloco_secao` + `SECOES_DANFSE` recortam o documento por seção antes de
  procurar o rótulo. Necessário porque rótulos se repetem — "Valor do Serviço"
  aparece em TRIBUTAÇÃO MUNICIPAL **e** em VALOR TOTAL DA NFS-E, e o
  "Nome / Nome Empresarial" do EMITENTE vem antes do mesmo rótulo no TOMADOR.
- `campo_danfse` aceita só `Rótulo\n \nValor` e `Rótulo\nValor` (esta segunda
  no bloco de tributação federal). **Não usar `\s*` solto**: campos vazios
  existem no DANFSe (ex: "Benefício Municipal") e um regex frouxo devolveria o
  **rótulo seguinte** como se fosse o valor. Os conversores
  (`converter_valor_br`/`converter_percentual`) são a última defesa — devolvem
  `None` para qualquer coisa que não seja número.
- `linha_planilha_nfse` monta a linha; o **código vem sempre do cadastro pelo
  CNPJ do tomador**, nunca por semelhança de nome. CNPJ não cadastrado → código
  vazio + observação, nunca um código inventado.
- `salvar_planilha_nfse` grava valores como `float` e datas como `date`/
  `datetime` (não texto), com formato de moeda/data, painel congelado e
  autofiltro — dá pra somar e filtrar no Excel direto.

**Retenções federais — por que estas colunas e não PIS/COFINS:** a planilha traz
"Prev. retida" (*Contribuição Previdenciária - Retida*) e "Contrib. sociais
retidas" (*Contribuições Sociais - Retidas*), que é o **agregado de
PIS+COFINS+CSLL retidos** (os 4,65%). São elas que descontam da nota: somadas,
batem com "Total das Retenções Federais" e explicam a diferença entre valor do
serviço e valor líquido. Os campos *PIS/COFINS - Débito Apuração Própria*, que
existem no DANFSe, **não** entram na planilha de propósito — são débito da
própria empresa, não descontam nada da nota, e lado a lado com a retenção
seriam confundidos numa conferência de valores. Campo sem retenção vem como
"-" no DANFSe e vira **célula vazia, nunca 0** (0 significaria "reteve zero").
Conferido nos 3.347 PDFs de julho/2026: `serviço − retenções = líquido` fecha
nas 3.134 NFS-e, das quais 34 têm retenção de contribuições sociais.
- Documento que não é DANFSe (ex: "Detalhamento do Faturamento", que vem no
  mesmo lote) → `extrair_dados_nfse` devolve `None` e a linha sai só com nome e
  motivo.

**Validação contra os dados reais** (~2.700 PDFs de julho/2026): conferência
campo a campo contra uma nota lida à mão deu 0 divergências; em 200 PDFs, 186
NFS-e reconhecidas sem nenhum campo vazio e com `ISSQN = BC × alíquota`
fechando em todas; e em 1.955 arquivos o código deduzido do CNPJ bateu com o
código já presente no nome do arquivo em 1.948 de 1.951 (99,85%).

**Divergências reais achadas nos dados** (não são bugs do programa):
- `10871 Esperança` (PCMSO/PGR/ESOCIAL) são na verdade notas do **10872 PAIVA**
  — trazem no tomador o CNPJ `86.846.763/0001-73`, que no cadastro é o 10872. O
  cadastro está coerente; o **nome dos arquivos** é que está errado, provável
  troca entre códigos vizinhos. Bom exemplo de por que se identifica por CNPJ.
- `PCMSO 11095 Serra Azul.pdf` e `PCMSO 10710 Martinica.pdf` estão em pastas
  PCMSO mas são "Detalhamento do Faturamento", não NFS-e.

## Contagem dos Protocolos dos Correios (aba 3)

Contrapartida do carimbo de código (v6.8.0/v6.9.0): além de identificar o
condomínio, essa aba conta quantas unidades receberam correspondência num
"Protocolo de Recebimento de Documento" e transforma isso em cobrança —
unidades × tarifa do lote, carimbadas no canto superior direito e somadas
numa planilha.

**Formato de saída do `winocr` (por que a extração não pode assumir linhas):**
o motor devolve a página inteira **numa linha só**, sem quebra nenhuma, e com
as colunas fora de ordem — os "Correio" da coluna Entrega saem todos
agrupados no fim do texto, não intercalados com as unidades a que pertencem.
Por causa disso nenhuma regex de contagem pode depender de início de linha
(`(?m)^`); `RE_UNIDADE_PROTOCOLO` e `RE_ENTREGA_PROTOCOLO`, em `logica.py`,
casam padrões soltos no texto corrido.

**O traço duplicado:** o OCR às vezes lê o traço entre o número da unidade e
o nome do morador em dobro (ex: `"702 - - Enny Marins de Lima"`, visto no
protocolo real do ASTORIA). Foi isso que motivou `(?:\s*[-–—])+` em vez de um
único `[-–—]?` em `RE_UNIDADE_PROTOCOLO` — sem o `+`, essas linhas não
batiam e a contagem saía sistematicamente abaixo do impresso.

**Regra de aceite** (`conferir_contagem_protocolo`): o `Listando N unidades`
impresso no documento é o valor que manda e funciona como dígito
verificador — sem ele, nada é aceito, mesmo que os dois conferidores
(contagem de linhas e contagem de "Correio") concordem entre si, porque
contar por OCR sozinho é chute com cara de precisão e o resultado vira
dinheiro cobrado. Com o `Listando` presente, basta que **um** dos dois
conferidores bata com ele para aceitar.

**Motores de OCR testados:** além do `winocr` (motor usado em produção),
foram testados o Tesseract 5.4 com modelo português e o RapidOCR (modelos
PP-OCR via ONNX) contra o mesmo lote de aceitação. Os três empatam em 4/4 nos
sinais usados (`Listando`, contagem de linhas, contagem de "Correio"), mas o
`winocr` é de 5 a 10× mais rápido e não pesa nada na distribuição (é motor
nativo do Windows, nenhum modelo pra empacotar) — por isso continua sendo o
motor principal. O RapidOCR ficou como reserva **opcional**
(`extrair_texto_rapidocr`, em `logica.py`), fora do `requirements.txt` e do
`.spec`, acionada por `extrair_texto_escaneado` só quando o `winocr` não
existe ou quebra. Se um dia for preciso saber **quais** unidades receberam, e
não só quantas, o RapidOCR é o motor a revisitar — ao contrário do `winocr`,
ele devolve blocos de texto com coordenadas, então dá para reconstruir a
ordem espacial das colunas em vez de só contar ocorrências no texto corrido.

**Lote de aceitação** (`C:\Users\Dell\Downloads\TESTE CORREIO`, 4 protocolos
reais, fora do repo — os valores vieram de números manuscritos nas próprias
folhas): 2 unidades / R$ 7,70, 15 / R$ 57,75, 3 / R$ 11,55 e 12 / R$ 46,20,
com tarifa de R$ 3,85 — total R$ 123,20. As fixtures de teste automatizado
(`tests/test_protocolo_contagem.py`) são sintéticas, sem nome de morador,
seguindo a mesma disciplina das fixtures de NFS-e.

**O carimbo de valor traz só o número, nunca a conta (v6.13.1):** era
`"12 un × R$ 3,85 = R$ 46,20"` e passou a ser `"R$ 46,20"`. O motivo não é
estético — o Superlógica lê esse carimbo, e com dois valores em reais na mesma
linha ele pode capturar a **tarifa** em vez do total. Depois da mudança o PDF
carimbado tem uma única ocorrência de "R$" e a tarifa não aparece em nenhum
lugar do texto extraível. Efeito colateral: o carimbo do valor calculado ficou
idêntico ao do informado à mão, então a distinção entre os dois existe só na
planilha (colunas Unidades e Tarifa preenchidas ou vazias, mais a observação).

**Qualidade de leitura fixa, não herdada da predefinição:** a aba 3 lê sempre
na melhor qualidade (`QUALIDADE_LEITURA_PROTOCOLO = 300`, em
`identificacao_por_cnpj_6_0.py`), ignorando a que estiver escolhida na
predefinição ativa. Antes ela herdava esse ajuste da aba 1, e uma revisão
mostrou que isso quebrava o lote inteiro quando a predefinição estava em
"Rápida". Medição nos quatro protocolos de referência:

| Documento | Rápida (72) | 150 | Normal (200) | Máxima (300) |
|---|---|---|---|---|
| -001 VILLARS | não reconhecido | ok | ok | ok |
| -002 ASTORIA | pendente | ok | ok | ok |
| -003 CARMEM | pendente | pendente | aceito por 1 conferidor | ok |
| -004 DIDEROT | não reconhecido | pendente | ok | ok |

A 300 os três sinais concordam nos quatro documentos, e o custo é de cerca de
meio segundo por página. A aba 1 continua usando a qualidade da predefinição —
lá o CNPJ tem dígito verificador e boa parte dos boletos tem texto nativo, o
que não vale para nenhum protocolo. Por isso `montar_config_atual` devolve só
aparência do carimbo: qualidade de leitura não entra nesse dict.

**Margem de segurança real da conferência:** o desenho de
`conferir_contagem_protocolo` falha para o lado seguro (rejeita quando os
conferidores não concordam com o `Listando`), mas a redundância observada no
lote real é menor do que "dois de dois". No protocolo `-003` (3 unidades), a
200 DPI, `Listando 3` foi confirmado por só **um** dos dois conferidores —
contagem de linhas deu 1 (subcontou), contagem de "Correio" deu 3 (bateu). A
aceitação passou porque basta um dos dois, não porque os dois concordaram.
Quem for endurecer essa regra (ex.: exigir os dois conferidores) precisa
saber que isso teria recusado um protocolo real do lote de aceitação.

**Painel de resultado (v6.13.0):** a aba 3 passou a encerrar com o mesmo
painel da aba 1 — cartões de resumo (Protocolos / Pendentes / Total) e
tabelas de calculados/pendentes — em vez do `messagebox` antigo.
`_montar_painel_resultado`, `_montar_faixa_cartoes` e `_montar_tabela_resultado`
(`identificacao_por_cnpj_6_0.py`) são os três helpers compartilhados entre as
duas abas; só o conteúdo interno muda.

A pendência se resolve pelo botão "Informar valor", digitando o **valor em
reais**, não a quantidade de unidades — porque é o que a pessoa já escreve à
caneta na folha do protocolo (nos protocolos de referência os manuscritos são
7,70 / 57,75 / 11,55 / 46,20; ver "Lote de aceitação" acima). Pedir a
quantidade obrigaria a pessoa a fazer de novo, de cabeça, a mesma conta que já
está escrita na folha.

A linha resolvida manualmente sai da planilha com Unidades e Tarifa vazias e
observação `Valor informado manualmente` (`linha_planilha_protocolo`, param.
`valor_manual`); o carimbo no PDF mostra **só o valor** em reais, sem a conta
por trás — então olhando só o papel carimbado ninguém distingue um valor
calculado (unidades × tarifa) de um valor digitado à mão. O rastro dessa
diferença existe **só na planilha**, nunca no PDF.

**O caso que motivou o recurso:** no protocolo real do `11049 APART HOTEL`, o
`Listando 76 unidades` impresso na folha foi riscado a caneta e substituído
por `78` escrito embaixo. A correção manual está fora do que a máquina lê —
ela continua vendo `76` — e, nesse documento, a contagem de "Correio" batia
exatamente com o número riscado (76), não com o corrigido (78): se o risco
estivesse mais leve ou não tivesse sido feito, os dois conferidores
concordariam entre si e o programa cobraria o valor errado com total
confiança. Não existe conferidor automático capaz de pegar esse caso — só uma
pessoa lendo a folha física resolve, e é exatamente para isso que existe o
"Informar valor".

## Protocolo do sistema antigo — "telnet" (aba 3, v6.19.0)

Segundo formato de protocolo dos Correios, do sistema antigo da Imodata.
Chega **misturado** com o novo na mesma pasta, então `formato_do_protocolo`
decide arquivo a arquivo pelo marcador: `PROTOCOLO CORRESPONDENCIA` (telnet)
ou `Protocolo de Recebimento de Documento` (novo). Com os **dois**
marcadores presentes devolve `None` de propósito — aplicar a extração
errada produz resultado plausível pelo motivo errado, e a conferência
humana não tem o que estranhar na tela.

**v6.19.1 — o ambíguo escapava da própria regra que o CLAUDE.md
prometia.** A interface (`_e_telnet`) traduzia esse `None` sempre como
"não é telnet", então um arquivo ambíguo caía no caminho do protocolo
novo e a extração rodava normalmente — a promessa acima nunca se
cumpria de fato. `classificar_formato_protocolo` (`logica.py`) é a
versão que não funde "nenhum marcador" com "os dois juntos": devolve
`"telnet"`, `"novo"`, `"ambiguo"` ou `None`. `formato_do_protocolo`
continua existindo, agora como wrapper fino que funde `"ambiguo"` em
`None` — mantido assim porque é o contrato que os testes e o resto do
código já esperavam; quem precisa da distinção usa a nova função
diretamente. Na aba 3, o ambíguo agora vira pendente com a observação
"Documento ambíguo: parece dois protocolos no mesmo arquivo", sem passar
pela extração de nenhum formato.

**Três decisões de extração, todas contra a tentativa óbvia**, e cada uma
medida nos 7 arquivos de referência:

| Ancorar em | Resultado | Por quê |
|---|---|---|
| rótulo `COD.` | 4/7 | o OCR leu `DAS`, `D ÀS` — estraga o rótulo, não os dígitos |
| **formato `1.DDDD`** | **7/7** | todo código tem 5 dígitos e começa em 1 (10002–11242) |
| `ENVIADO` | falha | sai `EWIADO`, `EFvTIADO` |
| **`PELO CORREIO`** | **6/7** | nunca falhou nas 21 leituras |
| rótulo `EDF:` (nome) | 1/7 | o winocr embaralha as colunas: rótulo e valor não ficam juntos |
| **texto inteiro** | **6/7 a 1.00** | varre em vez de ancorar |

**A folga do total é de 6 não-dígitos, nunca 20.** Com folga larga o regex
pulava o número certo e capturava outro do canto da folha: leu `38` no
lugar de `3` num condomínio de três unidades — R$ 146,30 em vez de
R$ 11,55.

**Três leituras de OCR com votação por maioria** (página inteira a 300, topo
a 400 e a 500). Nenhuma configuração sozinha lê os sete. A votação existe
por esse mesmo caso do `38, 3, 3 → 3`. O **total** exige 2 votos, porque
vira dinheiro; o **código** aceita 1, porque ainda passa pela conferência do
nome. **Empate no topo não elege ninguém** — escolher por ordem de chegada
seria arbitrário.

**O cadastro NÃO serve de dígito verificador aqui**, ao contrário do CNPJ.
772 dos ~1.241 números da faixa 10002–11242 são condomínios reais — 62% de
densidade —, então um dígito errado tende a produzir *outro condomínio
existente*. "O código está no cadastro" não prova nada. É por isso que o
nome impresso é cruzado com o do cadastro (limiar 0.90; os confirmados dão
1.00 e o único não confirmado deu 0.79).

**O aceite é deliberadamente mais frouxo que o do protocolo novo**:
`conferir_contagem_telnet` **preenche e marca** em vez de barrar, porque o
usuário confere imagem por imagem. Uma versão anterior barrava a linha sem
confirmação de nome e mandava para pendente um arquivo cujo código e total
estavam ambos corretos. Sem total legível ou sem código, aí sim é pendente
— com a frase **neutra** "Código do condomínio não confirmado no
protocolo": esse `if` cobre três causas que `apurar_protocolo_telnet` já
funde no mesmo `codigo=None` (nenhum código lido, código lido mas fora do
cadastro, empate entre dois códigos), e uma frase que afirmasse uma causa
específica contradizia a que `linha_planilha_protocolo` acrescenta logo
depois ("Código não identificado no documento") sempre que a causa real
fosse outra — as duas ficavam lado a lado na mesma célula da planilha se
contradizendo (v6.19.1).

**A marcação só é útil se aparecer, e isso exigiu três consertos juntos
(v6.19.1):** uma linha calculada com observação (telnet sem nome
confirmado, ou "Erro ao carimbar" de qualquer um dos dois formatos) tinha
a MESMA cor que uma linha limpa (nenhuma tag), ficava fora do filtro
"Pendentes" (o único gesto de "mostre o que precisa de mim"), e — mesmo
que alguém clicasse na linha — a coluna Observação ficava fora da tela,
porque a grade não tinha rolagem horizontal (7 colunas somam ~835px contra
os ~655px do painel da grade numa janela de 1180px com o divisor em 3:2).
Sozinho, nenhum dos três resolvia: cor sem filtro esconde a linha atrás de
"Tudo"; filtro sem cor não diferencia "precisa decidir" de "só conferir";
e os dois sem rolagem não mostram o *texto* da marca de qualquer forma.
`_estado_da_linha_protocolo` ganhou o estado `"atencao"` (calculado com
observação), com cor própria na grade — nunca o cobalto de `acento`, que é
reservado ao que *resolve* uma pendência — e alcançável pelo filtro
"Pendentes" (mesmo gesto que processado-sem-ID-SL já usava, em vez de um
5º botão).

**Nada a jusante mudou**: o dict tem `codigo` e `condominio`, que é tudo que
`linha_planilha_protocolo` e `_carimbar_protocolo` consultam. Valor,
carimbo lateral, bloco do Paybox, planilhas e painel são os mesmos.

**Um terceiro formato existe e está fora do escopo:** o "Meus Correios"
(`Cód. Condomínio: 10852-VILLA BRANCA`, `Qtde.`, `Classificação`), que são
SEDEX de valor fixo, cobrados por outro critério. Ele é o formato **fácil**
— o OCR lê o código em 6/6 e o nome vem no mesmo campo, servindo de
conferência. Quando for a vez dele, começar daí.

**A segunda leitura (topo@400) que decide o formato é reaproveitada, não
repetida (v6.19.1).** A classificação (`_classificar_arquivo_protocolo`,
na interface) dá uma segunda chance no recorte do topo a 400 DPI quando a
primeira leitura não achou nenhum marcador; esse texto agora é passado
adiante para `_leituras_telnet`, que o usa como uma das três leituras da
votação em vez de pagar o mesmo OCR de novo. Antes disso acontecia o
oposto do que o docstring de `_leituras_telnet` prometia — e todo arquivo
sem marcador na primeira leitura, **incluindo documento com texto nativo
que nem é protocolo**, pagava um OCR a 400 DPI que antes não pagava.

### Dois limites conhecidos, aceitos por desenho — não "endurecer" sem reler isto

**A conferência de nome é cega em dois casos.** `_melhor_score_do_nome`
usa uma janela do tamanho do nome do cadastro, então **um nome que seja
subcadeia do texto impresso pontua 1.00** — o score não sabe que só bateu
por estar contido em algo maior. Medido contra os 772 nomes reais do
cadastro: existem **8 pares de códigos que diferem em um dígito e se
confirmam mutuamente com score ≥ 0.90** — inclusive `10490 CENTRO COM
CONDE DE BONFIM RES` / `10491 CENTRO COM CONDE DE BONFIM`, o mesmo par que
este CLAUDE.md já destaca em "Cuidado: condomínios com nomes parecidos", e
`10174 JAB I` / `10175 JAB II`. Ou seja: o erro de OCR que a conferência
existe para pegar — um dígito trocado no código — é justamente onde ela é
cega, nesses pares. Além disso, **147 dos 772 nomes têm ≤ 6 caracteres e
28 têm ≤ 4** (`ARCO`, `AVRO`, `BONFIM`); para esses, a janela curta casa
com quase qualquer coisa numa folha cheia de `AP-104` e nome de rua. A
falha é **positiva em silêncio**: confirma o nome e não marca a linha —
não é "confirma errado e é óbvio", é "confirma errado e não avisa
ninguém". O texto acima ("o limiar não é delicado... 1.00... 0.79") é
verdadeiro nos 7 arquivos de referência e **irrelevante para esta falha**
— ela não é de limiar, é de janela.

**A "maioria de 2 em 3" tem só dois votantes independentes de verdade, não
três.** Das três leituras, duas são o MESMO recorte (topo 0–35% a 400 DPI
e topo 0–32% a 500 DPI) — só variam a qualidade, não o que é lido. Contra
um erro sistemático do cabeçalho (reflexo no papel, inclinação, dígito
borrado bem no campo do total), essas duas leituras tendem a concordar
entre si e derrotar a leitura de página inteira, que é a única
genuinamente diferente. O quórum efetivo é **um olhar independente, não
dois**. O caso medido e citado acima (`38, 3, 3 → 3`) deu certo porque o
valor discrepante estava justamente na leitura de página inteira; a ordem
inversa (`3, 38, 38`) elegeria **38** com "dois votos" e a linha sairia
limpa, sem marcação nenhuma. Quem quiser um 2-de-3 de verdade precisa
variar **o que a terceira leitura olha** (outra região, outro ângulo), não
só o DPI do mesmo recorte.

Spec: `docs/superpowers/specs/2026-09-03-protocolo-telnet-design.md`.
Validadores manuais: `tests/_validar_discriminador.py` e
`tests/_validar_telnet.py` (dependem de PDFs fora do repositório).

## Modo SEDEX e o protocolo "Meus Correios" (aba 3, v6.20.0)

Terceiro formato de protocolo, gerado pelo sistema **Agile**, e um modo de
lote que desliga a contagem de unidades. São coisas independentes:

```
FORMATO (novo / telnet / meus correios) → COMO ler o condomínio → o documento diz
MODO    (normal / SEDEX)                → SE conta unidades     → o documento NÃO diz
```

O mesmo envio SEDEX chega nos três formatos, e nada na folha distingue um
protocolo de cartas simples de um de SEDEX — por isso o modo é um alternador
que o usuário liga, e não algo deduzido do papel.

| Formato | Modo normal | Modo SEDEX |
|---|---|---|
| novo (Superlógica) | conta unidades | não conta, valor à mão |
| telnet | conta unidades | não conta, valor à mão |
| **meus correios** | **não conta (intrínseco)** | não conta |

**"Meus Correios" nunca conta unidades, em modo nenhum** — esse formato traz
`Qtde. 1`, um envio por documento. Isso é propriedade do formato, não do
modo, e não é esquecimento.

**O valor é digitado, e isso foi medido, não presumido.** Ele vem impresso
num comprovante escaneado, e a investigação do telnet mostrou que o valor
desses comprovantes (`TOTAL: 27  103,95`) **não sai em DPI nenhum** (200, 300,
400). Vale aqui o mesmo princípio já registrado na aba 2: valor e data não têm
dígito verificador, então uma leitura errada entra na planilha sem ninguém
perceber. Digitando na **grade do painel**, `_recarimbar_linha` regrava o PDF
com valor e bloco do Paybox — o anexo automático volta a funcionar, o que não
aconteceria preenchendo no Excel depois.

**A leitura é a mais fácil dos três formatos:** `Cód. Condomínio:
10852-VILLA BRANCA` traz código e nome no mesmo campo, então o código resolve
pelo cadastro e o nome ao lado confirma. Medido: **6/6 códigos** nos arquivos
de referência.

**Dois detalhes de regex que são load-bearing:**

- O marcador exige `COND` logo depois de `Cód` (`C[OÓ]D[.\s]{0,3}COND...`).
  O telnet traz `COD.: 1.1122.8)` na linha do total; um marcador que casasse
  só `COD.` tornaria **todo telnet ambíguo** e o lote inteiro cairia em
  pendente.
- O valor da Classificação é cortado no rótulo seguinte, e **não é sempre
  "Histórico"**: o winocr embaralha as colunas, e num dos seis arquivos vem
  `ID 12699 Contar: 1 Data ...`. Sem o corte, a observação carregaria meia
  página junto.

**A Classificação vai para a observação** (`824 - EBCT - SEDEX`), como está no
documento, sem traduzir — é o texto que a pessoa cruza com a tabela de preços
dos Correios para saber quanto digitar. É best-effort: não sendo legível, a
observação sai sem ela e nada é bloqueado.

**Nos 6 arquivos de referência, 5 são `590 - EBCT - CORREIO REG / AR` e só 1
é `824 - EBCT - SEDEX`**, apesar de o lote ser chamado de "os sedex". São
serviços diferentes, com preços possivelmente diferentes — é por isso que o
valor é por linha e não um só para o lote.

**O alternador não persiste entre sessões**, ao contrário do de lotes: aquele
é preferência do sistema de destino, este é propriedade do lote que está na
mesa. Deixá-lo ligado faria um lote de cartas simples sair inteiro sem
contagem.

**O que de fato torna `tarifa=None` seguro em modo SEDEX não é o código
antigo ser robusto a tarifa ausente — é `unidades` nunca chegar preenchido**
nos três ramos (telnet, protocolo novo, meus_correios) enquanto o modo está
ligado. Testado: `valor_protocolo(5, None)` estoura `InvalidOperation` — ou
seja, se algum ramo do roteamento algum dia deixasse passar `unidades`
preenchido com `tarifa=None`, a thread de processamento cairia sem
tratamento. `linha_planilha_protocolo`, `valor_do_carimbo` e
`validar_edicao_protocolo` só parecem "prontos para `tarifa=None`" porque a
entrada que recebem em modo SEDEX já chega com `unidades=None` — não porque
foram escritos pensando nisso. Quem mexer no roteamento depois precisa
preservar essa invariante: `unidades` sempre `None` quando o modo SEDEX
estiver ligado, nos três ramos.

**Não validado:** SEDEX nos formatos telnet e Superlógica. O usuário informou
que existem, mas não tinha exemplos na época. O desenho cobre os três; a
medição cobriu um. É o mesmo ponto cego que produziu a v6.18.1 — lá, validar
contra um lote de uma competência só escondeu uma variação de layout que
quebrava a leitura inteira.

Spec: `docs/superpowers/specs/2026-09-07-modo-sedex-design.md`.
Validador manual: `tests/_validar_meus_correios.py`.

## Anexo automático no Paybox (aba 3, v6.16.0)

O Paybox anexa um documento à despesa sozinho quando casa **valor, vencimento
e fornecedor** — não por código de barras, QR Code nem Pix. Como o Protocolo de
Recebimento de Documento não traz nenhum dos três por conta própria, o
Codificador carimba um bloco com eles (`montar_bloco_paybox`).

**Isso foi determinado por experimento contra o Superlógica real, e a hipótese
inicial estava errada.** A suspeita era que o Paybox casasse pelo código de
pagamento detectado na imagem. Foram testados quatro tipos de código impressos
no mesmo protocolo: **Code128, Interleaved 2 of 5 e Code39 não são lidos** — o
campo "Códigos detectados" volta vazio —, e **só o QR Code é reconhecido**. Mas
mesmo um QR com **BR Code Pix válido**, que o Superlógica extraiu corretamente
(apareceu o campo `pix-copia-e-cola` com o código inteiro), **não fez o Paybox
associar**. Passou a associar quando o **vencimento** entrou na página. Um teste
final sem QR nenhum, só com os três campos em texto, associou igual — então o
Pix é dispensável, e com ele sai o risco de deixar um código pagável em cada
lançamento.

**O Superlógica faz OCR da imagem da página**, não lê a camada de texto: a
prova é o nome do condomínio ter voltado como "CONDOMINIO DO EFICIO ..." em vez
de "EDIFÍCIO", erro que só OCR comete. Por isso o bloco é carimbado em
**negrito, corpo 11** — maior que o resto —, e por isso **não adianta** embutir
dado invisível (metadado do PDF, texto oculto): tem que estar visível na folha.

**Campos que o Superlógica extrai e não servem para associar:** ele lê e
reporta `Documento do pagador` (o CNPJ do condomínio, vindo do carimbo lateral)
e o `pix-copia-e-cola`. Nenhum dos dois participa do casamento.

**Três campos da planilha de importação recusam identificador inventado**, cada
um com sua validação: `linha_digitavel` exige código de barras de boleto
(`"Código de barras inválido"`), `chave_pix` exige BR Code EMV
(`"Chave BR Code inválida"`, e precisa de `tipo_de_chave_pix = 6` para QRCode),
e `etiqueta_paybox` **não é uma etiqueta**: exige uma URL do `sldocs.com.br`,
isto é, o link do documento já subido. Nenhum serve como gancho para um id
nosso.

**O vencimento carimbado e o da planilha têm que ser o mesmo.** Se divergirem,
o Paybox não acha o lançamento e o arquivo cai na fila manual sem erro nenhum —
falha silenciosa. É por isso que ele é perguntado uma vez só, no início do
lote, e `_acao_gerar_despesas` reusa o valor guardado em `_ctx_protocolos`.

**A despesa precisa existir ANTES do arquivo chegar ao Paybox.** O assistente
não age retroativamente sobre a fila: importar a planilha primeiro, mandar os
PDFs depois.

**O CNPJ do fornecedor é fixo no código** (`FORNECEDOR_PROTOCOLO_CNPJ`,
`logica.py`), por decisão do usuário — trocar de empresa de postagem exige
alterar o programa e gerar executável novo. As alternativas descartadas foram
uma coluna no `modelo_despesas.xlsx` e um campo em Configurações.

**A posição do bloco** é o **canto inferior direito**, escolhido pelo usuário
sobre um protocolo real: acima do rodapé da Imodata e do "1 de 1", e livre do
carimbo lateral (que é centralizado verticalmente e não chega ao rodapé). A
faixa abaixo da tabela de unidades foi descartada porque alguns protocolos
trazem um recibo dos Correios colado ali no meio da folha.

`MARGEM_DIREITA_BLOCO_PAYBOX`/`Y_BLOCO_PAYBOX`/`TAMANHO_BLOCO_PAYBOX`
(`identificacao_por_cnpj_6_0.py`) controlam isso. **A janela vertical é
estreita e o `Y` foi medido, não estimado:** num protocolo de página cheia a
tabela de unidades desce até ~81 e o rodapé da Imodata começa logo abaixo.
Varrendo as 3 páginas de um protocolo cheio (10380 ROXY) e um curto (10005
VILLARS), a posição livre mais baixa ficou em 74–77, e o valor usado é 77.
Descer mais encosta no rodapé e no "1 de N"; subir encosta nas linhas de
assinatura. Só as linhas estreitas (CNPJ/vencimento/valor) descem à altura do
rodapé — elas ficam à direita dele porque o bloco é alinhado à direita. A margem é medida da **borda
direita**, com o texto alinhado à direita, pelo mesmo motivo do carimbo de
valor: x fixo cairia fora da folha em página menor que A4. O corpo é **10**, e
não 11, porque a linha do fornecedor (~265pt, a mais larga) precisa caber sem
alcançar o carimbo lateral.

**As regras ficam em `Despesas > Paybox > Configurações gerais > ASSISTENTE
PARA ANEXOS AUTOMÁTICOS`**, exigem o perfil "Paybox - Alteração (1572)" e valem
para a licença inteira, não por condomínio. O protocolo é classificado pelo
Superlógica como `"outro"`, grupo que exige o mínimo de uma regra marcada.

## Por que a NFS-e NÃO anexa sozinha no Paybox — e por que carimbar não resolve

Investigado em 2026-09-01, a propósito de replicar nas notas da F&F o carimbo
que funciona nos Correios. **A conclusão é que não dá, e o motivo é
estrutural** — não é ajuste de posição, tamanho ou conteúdo do carimbo.

**A classificação do documento decide como o Superlógica o lê**, e isso muda
tudo:

| | Protocolo dos Correios | NFS-e da F&F |
|---|---|---|
| Classificação | `outro` | `nfse` |
| De onde vêm os campos | **OCR da folha** | **esquema canônico da NFS-e**, via QR Code |
| Carimbar texto novo adianta? | **sim** — é a única fonte que ele tem | **não** — o que não está no esquema é descartado |

O protocolo funciona justamente por ser um documento que o Superlógica **não
sabe classificar**: sem esquema, ele cai no OCR da página, que é onde o
carimbo está. Numa NFS-e ele lê o QR (aparece em "Códigos detectados", com o
link da consulta pública da `nfse.gov.br`), preenche um esquema fixo
(`access-key`, `invoice-number`, `invoice-value`, `net-value`, `emission-date`,
`iss-value`, `Documento do pagador`, `Documento do destinatário`) e **não olha
a folha**.

**O assistente exige o campo Vencimento preenchido no arquivo**, e o layout da
NFS-e não tem esse campo. Como a extração é por esquema, ele nunca é
preenchido — então **nenhuma NFS-e associa sozinha**. Não é intermitente: é
por construção. Isso combina com o achado já registrado acima, de que o Pix
válido não bastou e a associação só começou quando o vencimento entrou na
página.

### A prova

Nota do `10048 ALVORADA` (NFS-e 12367, R$ 21,31) contra a despesa `54927`:

| Critério marcado no grupo NOTA FISCAL | No documento | No lançamento |
|---|---|---|
| CNPJ do condomínio | `01183801000100` | `id_condominio_cond: 64` |
| CNPJ do fornecedor | `13736666000154` | `st_cpf_con: 13736666000154` |
| Número da nota fiscal | `12367` | `st_documento_des: "12367"` |
| Valor igual ao do documento | `21.31` | `vl_valor_pdes: "21.31"` |

Os quatro batem, o lançamento estava aberto (`fl_liquidado_pdes: "0"`), o tipo
era fiscal (`id_tipo_doc: "1"`, "Nota Fiscal") e o PDF foi enviado **depois** de
a despesa existir. Mesmo assim `arquivos: []`. **Preenchendo o Vencimento do
arquivo à mão, associou na hora** — é esse teste que isola a causa.

### O bloco completo dos Correios também foi testado, e também foi ignorado

Depois da conclusão acima, o bloco **no formato exato dos Correios** (negrito,
corpo 10, alinhado à direita, com fornecedor, CNPJ, `VENCIMENTO:` e `VALOR:`)
foi carimbado em duas notas — `10500 WALQUIRIA` (NFS-e 12934) e
`10049 PIRATININGA` (NFS-e 12739) — usando `processar_pdf`/`criar_overlay` do
próprio programa, para que um resultado negativo não pudesse ser atribuído a
um render improvisado. A posição foi **medida** na nota (`y=94`), não
reaproveitada dos Correios (`y=77`), porque o rodapé da NFS-e é outro.

| Teste | Bloco na folha | Vencimento preenchido à mão | Associou |
|---|---|---|---|
| ALVORADA | não (só uma linha `Vencimento: ...`) | não | **não** |
| ALVORADA | não | **sim** | **sim** |
| Walquiria / Piratininga | **sim, formato dos Correios** | não | **não** |
| Walquiria / Piratininga | sim | **sim** | **sim** |

**A variável que decide é sempre o campo Vencimento do arquivo, e nunca o
carimbo.** Não tente de novo com outro formato, outra posição ou outro corpo
de letra: o que falha não é a legibilidade, é o fato de a folha não ser lida.

### Quatro hipóteses foram testadas e descartadas antes desta

Vale a pena registrá-las, porque todas eram plausíveis e custaram experimento:

1. **Carimbar o vencimento na nota** — feito, com carimbo grande e legível
   (`Vencimento: 05/09/2026`). Ignorado, pelo motivo acima.
2. **A regra casaria por vencimento** — não: "Data do documento" está
   **desmarcada**, e esse campo é a data de emissão, outra coisa.
3. **`numero_documento` vazio no lançamento** — não: vai preenchido com o
   número da nota, e confere.
4. **Grupo de regras separado para `nfse`** — não existe. Os tipos de
   documento do Superlógica são 9 (`1 Nota Fiscal`, `2 Imposto`, `3 Fatura`,
   `4 Recibo`, `5 Cupom Fiscal`, `6 Outros`, `7 Folha de pagamento`,
   `8 Apólice`, `9 DANFE`), e o grupo "NOTA FISCAL" é mesmo o que se aplica.

### Consequências

- **Não implementar o bloco do Paybox nas notas da F&F.** Seria carimbar numa
  folha que, para essa classe de documento, não é lida.
- **O carimbo do código + tipo de serviço CONTINUA, e não tem nada a ver com
  o Paybox** (ex: `10049 PIRATININGA - E-Social`). Ele existe porque **a nota
  da F&F não diz a que serviço se refere** — eSocial, PCMSO, PGR, exame e
  treinamento saem com a mesma descrição —, e é o síndico, no condomínio, que
  precisa saber o que está pagando ao olhar o papel. É por isso que a
  predefinição da F&F tem o tipo de serviço editável por lote. Quem for
  desligar o carimbo por causa do Paybox está mexendo na coisa errada: são
  dois propósitos independentes que por acaso moram no mesmo lugar.
- **Não esconder nem remover o QR Code** para forçar a classificação `outro`.
  Funcionaria pelo mesmo mecanismo dos Correios, mas seria degradar de
  propósito a verificabilidade de um documento fiscal para enganar um
  classificador.
- **O caminho que resta é o suporte do Superlógica** (que a associação de
  `nfse` não dependa do vencimento, ou que ele seja herdado da despesa
  candidata) ou, se o endpoint de upload aceitar metadados, **enviar as notas
  já com o vencimento preenchido pela API** — a licença tem API (foi de lá que
  saiu o JSON da despesa `54927`). Essa é a única pergunta que ainda pode
  reabrir a automação.

## Planilha de Despesas do Superlógica (aba 3, v6.14.0)

Terceiro passo do fluxo dos Correios, depois de contar e carimbar: o botão
"Gerar planilha do Superlógica", no rodapé do painel de resultado, transforma
o lote em arquivo de importação de despesas — uma linha por protocolo
cobrável, `condomínio` = **ID SL** do cadastro, `valor` = valor apurado.

**O modelo é do usuário e é copiado, não reconstruído.** O arquivo de
importação (32 colunas na versão atual) é baixado do próprio Superlógica; numa
cópia dele o usuário preenche a linha 2 com o que se repete em todo lançamento
(fornecedor, favorecido, `conta_categoria`, tipo de documento, forma de
pagamento e uma chave numérica), deixando `condomínio` e `valor` em branco.
`gerar_planilha_despesas` (`logica.py`) abre esse arquivo e o preenche.
Copiar em vez de montar do zero preserva formatos de célula, validações e
colunas ocultas que o importador pode exigir e que se perderiam em silêncio —
e o layout, sendo do Superlógica, pode mudar sem aviso.

**A linha 2 é o molde e some.** Ela é substituída pela primeira linha real;
nenhuma linha de exemplo pode sobrar no arquivo final (seria despesa fantasma
na importação). Modelo salvo com várias linhas de exemplo tem o excesso
removido.

**Colunas achadas pelo nome, nunca pela posição** — `_normalizar_cabecalho`
tira acento e caixa, então `condomínio`, `Condomínio` e `CONDOMÍNIO` são a
mesma coluna. Modelo sem uma das duas colunas levanta erro com o nome da que
falta, em vez de gerar arquivo silenciosamente errado.

**Mesmo condomínio em dois protocolos gera duas linhas.** No lote de
referência, dois protocolos são do `11161 MARILIA`: saem como `496 / 138,60` e
`496 / 119,35`. Somar quebraria a correspondência entre lançamento e papel.

**O vencimento é perguntado a cada geração, não fica no modelo.** É dado do
lote, como a tarifa — e foi justamente por estar no modelo que ele quebrou a
importação: o Excel guarda data como número de série (`21/08/2026` é `46255`) e
só o **formato da célula** diz que aquilo é data. Com a célula em `General`, o
arquivo saía com o número cru, o Superlógica gravava `01/01/1970` e **recusava
os lançamentos** — de três, um entrava, e mesmo esse com a data errada.
`converter_data_digitada` (`logica.py`) aceita só o formato brasileiro
(`21/08/2026`, `21-08-2026`, `21.08.2026`, `21/08/26`); `2026-08-21` é recusado
de propósito, porque misturar as duas convenções é como uma data tipo `03/04`
acaba lançada com o mês trocado.

**A `chave` também é perguntada no início do lote (v6.17.0), pelo mesmo
motivo.** Ela muda a cada importação, e antes era editada à mão na planilha já
gerada. Abrir o arquivo no Excel só para isso é o que fez a coluna
`vencimento` perder o formato de data e voltar a ser número cru — o
Superlógica gravou 01/01/1970 outra vez, **sem acusar erro nenhum**. Com a
chave vindo do programa, não há mais razão para abrir a planilha antes de
importar. `chave=None` mantém o que estiver no modelo.

**Guarda das colunas de data, como rede.** `_data_do_molde` converte para data
de verdade o número que aparecer em `vencimento`, `competência` ou `liquidação`
quando estiver na faixa 2000–2099; qualquer outra coisa levanta erro com o nome
da coluna, em vez de virar cobrança com data errada. Perguntar o vencimento é o
conserto na origem; a guarda cobre quem mesmo assim preencher data no modelo.

**A primeira suspeita foi a errada, e só não virou conserto porque foi
testada.** Com três lançamentos e um só importado, a hipótese natural era o
campo `chave`, repetido nas três linhas e com cara de identificador único. Dois
arquivos variando só a `chave` e um variando só a data mostraram que a `chave`
não tinha nada a ver. Vale lembrar disso antes de "consertar" o próximo sintoma
parecido.

**A geração trava por completo** enquanto houver protocolo cobrável
incompleto, e `lancamentos_de_despesa` devolve os travados separados por
motivo, porque a ação é diferente: pendência sem valor resolve no próprio
painel ("Informar valor"); condomínio sem ID SL resolve na aba de Cadastro.
Arquivos que não são protocolo não travam — nunca deveriam virar despesa.
**10 dos 764 condomínios estão sem ID SL**, então essa trava vai acontecer; o
conserto é preencher o campo no cadastro.

**`lancamentos_de_despesa` recebe o `resultado` do painel, não as linhas da
planilha.** Nas linhas prontas, um arquivo que não é protocolo e uma pendência
"não foi possível ler o documento" ficam idênticos (código, condomínio e valor
vazios) — mas um deve travar e o outro deve ser ignorado. Só o `resultado`
separa os dois grupos.

**O item resolvido à mão precisa carregar `codigo`** no dict que vai para
`resultado["processados"]` (`_acao_informar_valor`). Sem ele não há como achar
o ID SL, e justamente os protocolos que exigiram atenção manual travariam a
geração.

Validado contra o modelo e o cadastro reais: os três protocolos multipágina
saem como `496 / 138,60`, `496 / 119,35` e `649 / 231,00`, com os seis campos
repetidos idênticos nas três linhas e nenhuma linha em branco.

## Divergências de lógica só no 6_0 (pós-redesign)

O redesign 6_0 nasceu como reforma **exclusiva de interface** (lógica
byte-idêntica ao 5_3), mas depois ganhou 3 correções/recursos que só existem
no 6_0 — o `identificacao_por_cnpj_5_3.py` não tem nenhuma delas:

1. **`buscar_por_nome_arquivo` ignora palavras de tipo de documento**
   ("QUITADO"/"NF"/"RECIBO"/"DEMONSTRATIVO"/"NOTA"/"FISCAL"/"BOLETO") antes do
   fuzzy match — nomes como "ARGENTINA QUITADO 05.26.pdf" caíam abaixo do
   limiar de 0.72 por causa da palavra extra.
2. **`CNPJ_FLEX` tolera `-`/`.` como separador**, além de `/`, no CNPJ do
   campo (ex: CO-ESTIPULANTE) — alguns recibos da FedCorp escrevem
   "08.578.541-0001-03" em vez de "08.578.541/0001-03".
3. **Match por código exato no nome do arquivo + desempate por cadastro em
   CNPJ ambíguo** — ver seção "Predefinições de lote" acima.

## Ideia de melhoria discutida, não implementada

- **LLM como último fallback** (não como motor principal): mandar a imagem da
  página pra API do Claude só quando as etapas acima falharem. Mais robusto a
  variação de layout, mas custa por documento e envolve enviar dados de
  clientes pra fora — decisão de não implementar por ora.

## Convenções do código

- Comentários e nomes de variáveis em português (ex: `cadastro`, `formatar_cnpj`).
- Documentos sempre normalizados para só dígitos internamente
  (`normalizar_cnpj`) — 14 para CNPJ, 11 para CPF —, formatados só na
  exibição/planilha (`formatar_documento`). O comprimento é o que distingue os
  dois; nunca inferir o tipo de outro jeito.
- Dependências: `customtkinter` (interface 6_0; traz `darkdetect`), `pypdf`,
  `reportlab`, `openpyxl`, `pymupdf`, `Pillow`, `winocr` (Windows). Versões
  fixas em `requirements.txt` (mantido em sincronia manualmente — sem
  ferramenta de lockfile automática). Instalação: `python -m pip install -r requirements.txt`.
- Cores da interface 6_0 nunca hardcoded soltas: sempre via `self.tema_atual[chave]`
  (dicts `TEMA_CLARO`/`TEMA_ESCURO`). Todo widget CustomTkinter usa `corner_radius=0`.
- Vocabulário da interface é para leigos: nunca expor "OCR" ou "DPI" em texto visível
  (usar "boletos escaneados", "qualidade de leitura").
- Gerar a distribuição: duplo clique em `gerar_exe.bat` (instala as
  dependências, roda os testes, só empacota se passarem, e monta o
  `dist/CODIFICADOR.zip`). A receita é `codificador.spec`, versionada de
  propósito — o `.gitignore` ignora `*.spec` genéricos mas abre exceção pra
  ela, porque antes o build só existia na máquina de quem o fazia.
  **O formato entregue é `--onedir` zipado, não `--onefile`**: o entregável é
  o `CODIFICADOR.zip`, contendo a pasta `Codificador/` com o executável, o
  `_internal/` e o `cadastro_condominios.xlsx` ao lado dele (é ali que
  `pasta_base()` procura o cadastro e grava `config.json` e os logs). Com
  `--onefile` o app extrairia ~40 MB no temp a cada abertura, num programa de
  uso diário, e chamaria mais atenção do antivírus/SmartScreen. O usuário
  precisa extrair a pasta inteira — o `.exe` não roda sozinho.
  **O ambiente do build precisa bater com `requirements.txt`**: o PyInstaller
  não empacota o que não está instalado, então gerar o exe numa máquina sem
  `winocr` produz um binário com a leitura de boletos escaneados morta, sem
  erro visível. `gerar_exe.bat` cobre isso instalando os requirements antes.
  Detalhes que já quebraram o exe (ícone via `sys._MEIPASS`, assets do
  CustomTkinter, `hiddenimports` dos bindings `winrt.*`) estão comentados
  dentro do `.spec`. Dependência de build fixada em `requirements-build.txt`.
- Testes: `python -m unittest discover -s tests -p "test_*.py"` (ou duplo
  clique em `rodar_testes.bat`). Usam `tests/cadastro_teste.py` (9 condomínios
  fixos) e fixtures de texto em `tests/dados/`, nunca a planilha real. Cobrem
  identificação por nome, extração de CNPJ, validação, desempate, nome de saída
  e config/migração, além da extração de dados das NFS-e
  (`test_extracao_nfse.py`, sobre as fixtures `nfse_ff.txt` — nota sem
  retenção federal — e `nfse_ff_retido.txt` — com retenção
  "3 - PIS/COFINS/CSLL Retidos"). A regra de
  desempate vive em `desempatar_por_cadastro`
  (extraída do loop justamente para ser testável). Regenerar fixtures:
  `python tests/_gerar_fixtures.py` (precisa dos PDFs-fonte, fora do repo).
