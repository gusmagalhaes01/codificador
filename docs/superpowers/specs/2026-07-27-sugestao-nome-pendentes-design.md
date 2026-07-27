# Sugestão por nome para pendentes "Não foi possível ler" — especificação

Data: 2026-07-27
Estado: aprovado, a implementar

## Problema

Hoje, quando a extração de CNPJ do conteúdo do PDF não encontra **nenhum**
candidato válido (mesmo depois de OCR e da escalada de DPI), o arquivo vira
pendente com motivo "Não foi possível ler" e a única ação disponível é "Abrir
PDF" — nem "Cadastrar" nem "Escolher" funcionam, porque não há CNPJ nenhum
pra oferecer.

Boa parte desses casos são arquivos escaneados cujo nome já indica o
condomínio (ex: os dois "CONDE DE BONFIM", "ASSOCIACAO", "CENTRAL
COPACABANA") — o match por nome do arquivo (`buscar_por_nome_arquivo`) até
identifica candidatos prováveis, mas descarta o resultado quando fica
ambíguo (regra `LIMIAR_DIFERENCA_AMBIGUA`), porque essa trava foi feita pra
impedir decisão **automática** errada — não pra esconder a sugestão do
funcionário.

## Objetivo

Quando não sobrar nenhum candidato de CNPJ, tentar sugerir candidatos por
nome (arquivo + texto OCR) e deixar o funcionário escolher manualmente,
reaproveitando o popup "Escolher" que já existe. Nunca decidir sozinho por
nome — mantém a regra já fixada no CLAUDE.md ("a identificação sempre deve
ser feita pelo CNPJ, nunca por similaridade de nome sozinha"): aqui o CNPJ
final vem do cadastro, escolhido por uma pessoa depois de olhar o PDF, não
inferido automaticamente.

## Decisão

Nova função em `logica.py`, `candidatos_por_nome`:

```python
def candidatos_por_nome(nome_arquivo, texto, cadastro, limite=8):
    """Sugere até `limite` CNPJs candidatos comparando o nome do arquivo e o
    texto extraído do PDF (se houver) contra os nomes do cadastro — usado só
    como SUGESTÃO para escolha manual (nunca decide sozinho), diferente de
    buscar_por_nome_arquivo que bloqueia em caso de ambiguidade. Descarta
    palavras de tipo de documento do nome do arquivo (mesmo tratamento de
    buscar_por_nome_arquivo). Devolve lista de CNPJs normalizados, do mais
    provável ao menos provável; lista vazia se o melhor candidato não
    atingir LIMIAR_SCORE_NOME."""
```

Comparação:
- Alvo 1: nome do arquivo, normalizado e sem palavras de tipo de documento
  (`normalizar_texto_busca` + `remover_palavras_tipo_doc`, já existentes).
- Alvo 2 (se `texto` não for vazio): `sugerir_nome_condominio(texto)`,
  normalizado do mesmo jeito — cobre o caso de o nome do arquivo ser genérico
  mas o pouco texto que o OCR conseguiu ler ter o nome do condomínio.
- Cada alvo é comparado contra todos os nomes do cadastro
  (`SequenceMatcher`, mesmo método de `buscar_por_nome_arquivo`); o maior
  score de cada CNPJ (considerando os dois alvos) é o que conta.
- Só sugere alguma coisa se o melhor score atingir `LIMIAR_SCORE_NOME` (0.72
  — mesmo piso do match automático). Atingindo esse piso, inclui o melhor
  candidato **e** qualquer outro cujo score fique a menos de
  `LIMIAR_DIFERENCA_AMBIGUA` (0.08) de distância dele — é o mesmo cálculo que
  `buscar_por_nome_arquivo` usa pra decidir "é ambíguo", só que aqui, em vez
  de descartar tudo, devolve o grupo inteiro. É assim que os dois "CONDE DE
  BONFIM" aparecem juntos na lista em vez de sumirem (um deles sozinho não
  bateria o piso de 0.72, mas fica perto o bastante do outro pra entrar).
  `limite=8` é só um teto de segurança pro popup não ficar absurdamente
  longo — **testado contra a planilha real** (~750 condomínios): o grupo de
  nomes parecidos com "CONDE DE BONFIM" tem 4 candidatos (os dois
  verdadeiros + 2 falsos positivos por acaso), cabendo dentro do teto sem
  cortar nenhum. Uma primeira versão usava `limite=3`, que cortava
  exatamente o candidato certo quando havia falsos positivos com score mais
  alto — só apareceu testando contra o cadastro real, não contra os 9
  condomínios fixos dos testes automatizados.
- Ordena por score decrescente, corta em `limite` (padrão 3).

## Onde entra no fluxo

Em `_processar_em_thread` (`identificacao_por_cnpj_6_0.py`), no bloco que
hoje monta o pendente `"tipo": "nao_lido"` quando `len(candidatos) == 0`:
antes de decidir o motivo, chama `candidatos_por_nome(nome, texto, self.cadastro)`.

- Se devolver 1+ candidatos: pendente vira `"tipo": "ambiguo"` (reaproveita o
  botão "Escolher" e o popup existentes, sem nenhuma mudança de UI),
  `"candidatos"` recebe a lista, e o motivo exibido na tabela passa a ser
  "Nome parecido encontrado" em vez de "Não foi possível ler" — pra deixar
  claro que a sugestão veio do nome, não do CNPJ do documento.
- Se devolver lista vazia: comportamento **idêntico** ao de hoje — tipo
  `"nao_lido"`, motivo "Não foi possível ler", só "Abrir PDF" disponível.

`texto` usado na chamada é a mesma variável já populada pelo pipeline (texto
nativo ou OCR, o que tiver sido lido até ali) — nenhuma leitura extra de PDF.

## Nenhuma mudança em

- `buscar_por_nome_arquivo`, `extrair_cnpj_tomador`, `desempatar_por_cadastro`
  — a nova função é só mais uma etapa, não substitui nenhuma existente.
- Casos onde já existe 1+ candidato de CNPJ (rodapé continua sendo CNPJ
  manda, como hoje).
- Formato do `processamento.log`.
- Popup "Escolher" (`_acao_escolher_pendente`) — já trata corretamente CNPJ
  não-cadastrado entre os candidatos oferecidos, então funciona sem alteração
  mesmo quando o candidato sugerido por nome ainda não estiver no cadastro
  (raro, mas possível).

## Testes

Novo arquivo `tests/test_candidatos_por_nome.py`, cobrindo com o cadastro
fixo de testes (`tests/cadastro_teste.py`):

- Nome de arquivo batendo com um único condomínio → lista com 1 CNPJ.
- Nome de arquivo ambíguo entre os dois "CONDE DE BONFIM" → lista com os
  **dois** CNPJs (é o comportamento que faltava hoje).
- Nome de arquivo sem nenhuma correspondência (score baixo) → lista vazia.
- Texto OCR com nome do condomínio, nome do arquivo genérico (ex:
  "DIGITALIZACAO_001.pdf") → usa o alvo do texto, lista com o CNPJ certo.
- `texto` vazio (string `""`) → não quebra, usa só o nome do arquivo.
- `limite` respeitado (testado com `limite=1`).
- Distratores com score mais alto que o candidato certo não escondem esse
  candidato do resultado (regressão do bug encontrado testando contra a
  planilha real — ver acima).

## Fora de escopo

- Qualquer forma de decisão automática por nome (mantém a regra do CLAUDE.md).
- Mudar o piso `LIMIAR_SCORE_NOME` ou `LIMIAR_DIFERENCA_AMBIGUA` usados pelo
  match automático em `buscar_por_nome_arquivo` — ficam como estão.
- Interface: o popup "Escolher" já existe e já trata os casos necessários,
  não precisa de nenhuma tela nova.

## Critério de sucesso

Um pendente com nome de arquivo "CONDE DE BONFIM.pdf" e conteúdo ilegível
(sem CNPJ nenhum) passa a aparecer com "Escolher" disponível, oferecendo os
dois candidatos — em vez de só "Abrir PDF". Nenhum teste dos 34 existentes
quebra.
