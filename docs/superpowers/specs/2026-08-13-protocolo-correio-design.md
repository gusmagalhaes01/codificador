# Protocolo de Recebimento de Documento (Correios/Imodata) — especificação

Data: 2026-08-13
Estado: aprovado, implementado

## Problema

O app só sabia identificar documentos pelo CNPJ do tomador. O "Protocolo de
Recebimento de Documento" — recibo de entrega de correspondência da Imodata,
escaneado, sem nenhum CNPJ — sempre cairia como "Não foi possível ler" sem
nenhuma ação possível: o nome do arquivo é só um número de protocolo
(ex: `20260813183654301-001.pdf`) e o texto não tem nenhum rótulo que
`sugerir_nome_condominio` reconheça, então nem a sugestão por nome (recém
implementada) ajudaria.

Testado com 4 exemplos reais (fornecidos pelo usuário): todos seguem o mesmo
formato, com o código do condomínio já pronto no texto — ex:
`W700A VILLARS (10005) Protocolo de Recebimento de Documento...`.

## Decisão

Nova função em `logica.py`, `extrair_codigo_protocolo_correio(texto)`:
reconhece o marcador `"Protocolo de Recebimento de Documento"` e extrai o
código entre parênteses logo antes dele (`\((\d+)\)\s*Protocolo de
Recebimento de Documento`, case-insensitive). Devolve o código como string,
ou `None` se o marcador não aparecer.

**Detecção automática pelo conteúdo**, não uma predefinição nova: qualquer
lote pode ter esses protocolos misturados com boletos normais.

## Onde entra no fluxo

Em `_processar_em_thread`, logo depois que `texto` é finalizado (nativo ou
via OCR, sem mudança nenhuma nessa parte) e **antes** de tentar
`extrair_cnpj_tomador`:

- Marcador encontrado + código no cadastro → `candidatos = [cnpj]` (resolvido
  via `_codigos_do_cadastro`, o mesmo índice reverso já usado por
  `buscar_por_nome_arquivo`), e cai no caminho normal de identificação
  (mesmo carimbo, mesmo registro em processados) — só muda o rótulo de
  origem para "pelo código do protocolo dos Correios".
- Marcador encontrado + código **não** cadastrado → pendente dedicado
  "Código não cadastrado" (distinto do "CNPJ não cadastrado" existente,
  porque aqui não há CNPJ nenhum pra oferecer — `cnpj: None` no registro do
  pendente). O botão "Cadastrar" ainda funciona: o campo CNPJ só fica em
  branco pro funcionário preencher na hora.
- Marcador **não** encontrado → comportamento idêntico ao de hoje (extração
  de CNPJ, desempate, retry de OCR — nada mudou nesse caminho).

Quando o marcador é encontrado, o bloco de desempate (`preferir_cadastrado_em_ambiguo`)
e o retry de DPI (`proximo_dpi_maior`) são pulados — não fazem sentido aqui:
não há CNPJ ambíguo nem checksum pra falhar.

## Testes

`tests/test_protocolo_correio.py` — 6 testes, usando trechos representativos
do formato real (só o cabeçalho, sem nomes de moradores — dado sensível que
não precisa estar no fixture): marcador presente extrai o código certo,
marcador ausente devolve `None`, texto vazio não quebra, marcador sem
parênteses não confunde, código resolve no cadastro de teste, código
desconhecido não resolve.

**Validado contra os 4 PDFs reais e a planilha real** (não só o fixture de
teste): os 4 resolveram corretamente (VILLARS/10005, ASTORIA/10010,
CARMEM/10015, DIDEROT/10017) — mesma disciplina de validação que pegou o bug
do `candidatos_por_nome` antes desta função ir pro `.exe`.

## Fora de escopo

- Outros formatos de recibo/correio que não esse (usuário confirmou que é só
  esse por ora).
- Extrair o nome do condomínio do texto do protocolo — desnecessário, o
  código já resolve sozinho no cadastro.
- Qualquer mudança na extração de CNPJ existente.

## Critério de sucesso

Os 4 arquivos de exemplo passam a identificar automaticamente pelo código,
sem precisar de OCR de nome nem intervenção manual — confirmado.
