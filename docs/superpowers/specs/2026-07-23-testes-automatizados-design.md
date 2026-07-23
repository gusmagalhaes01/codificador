# Testes automatizados do Codificador — especificação

Data: 2026-07-23
Estado: aprovado, a implementar
Versão alvo: **v6.3.0** (parte de v6.2.0). Não é v7 porque nada muda para quem
usa o app — é rede de segurança interna, não recurso novo.

## Problema

A lógica de identificação já foi alterada três vezes depois do redesign 6_0
(palavras de tipo de documento, separador do CNPJ, código no nome do arquivo).
Cada alteração foi validada com scripts ad-hoc que rodaram uma vez e foram
apagados. Não sobrou nada que garanta que esses casos continuam funcionando.

O risco concreto: uma mudança futura em `buscar_por_nome_arquivo` ou em
`extrair_cnpj_tomador` pode fazer o programa carimbar **o código errado** num
documento financeiro — por exemplo trocar `10490` (CENTRO COM CONDE DE BONFIM
RES) por `10491` (CENTRO COM CONDE DE BONFIM) — e ninguém percebe olhando o
PDF. O erro só apareceria na contabilidade, depois.

## Objetivo

Congelar em testes automatizados os casos reais já corrigidos, de forma que
qualquer regressão apareça em segundos, sem teste manual.

Não-objetivo: aumentar cobertura por cobertura. O alvo é o que carimba código
errado — identificação e validação. Interface, empacotamento e OCR ficam de
fora (ver "Fora de escopo").

## Decisões tomadas

**Ferramenta: `unittest` (biblioteca padrão).** Zero instalação nova na
máquina de trabalho. `pytest` seria um pouco mais confortável de escrever, mas
exigiria `pip install` — e a diferença prática, para asserções simples como as
daqui, é pequena.

**Execução por duplo clique.** Um `rodar_testes.bat` na raiz roda tudo e mostra
o resultado, para não depender de linha de comando.

**Dados de teste: texto extraído, não os PDFs.** Os fixtures são arquivos
`.txt` com o texto que o programa lê de dentro dos PDFs reais. É exatamente o
que a lógica consome, e mantém os testes rápidos e o repositório leve.

**Cadastro próprio e congelado.** Os testes NÃO leem
`cadastro_condominios.xlsx`. Usam uma lista fixa de ~9 condomínios embutida em
`tests/cadastro_teste.py`. Motivo: se lessem a planilha real, cadastrar um
condomínio novo poderia quebrar um teste sem que nada esteja errado — e teste
que falha à toa é pior que teste nenhum, porque ensina a ignorar o aviso.

## Estrutura de arquivos

```
tests/
  cadastro_teste.py          lista fixa de condomínios (dict cnpj -> {codigo, nome})
  dados/
    fedcorp_recibo_hifen.txt   recibo FedCorp, CO-ESTIPULANTE com "-" (SAN REMO)
    nfse_ff.txt                NFS-e da F&F (KLOSTERS)
    nfse_imodata.txt           nota Imodata com 2 CNPJs (MANHATTAN)
    nfse_avulsa.txt            NFS-e avulsa (LAGO MAGGIORE)
    boleto_avulso.txt          boleto avulso, campo Pagador (VILLE DE BEAUVAIS)
    cnpj_checksum_invalido.txt CO-ESTIPULANTE com CNPJ de dígito errado
  test_identificacao.py      nome do arquivo, extração de CNPJ, validação, desempate
  test_config.py             predefinições, migração, arquivo corrompido
  test_saida.py              nome do arquivo de saída
rodar_testes.bat
```

Os testes não são importados pelo app, então o PyInstaller não os inclui no
`.exe` — nenhum ajuste de build é necessário.

## Cadastro de teste (valores confirmados em uso real)

| CNPJ | Código | Nome |
|---|---|---|
| 40.338.774/0001-41 | 10695 | ARAUJO LIMA |
| 01.195.716/0001-54 | 10004 | KLOSTERS |
| 08.578.541/0001-03 | 10002 | SAN REMO |
| 05.695.194/0001-00 | 10490 | CENTRO COM CONDE DE BONFIM RES |
| 29.361.458/0001-58 | 10491 | CENTRO COM CONDE DE BONFIM |
| 07.448.975/0001-26 | 11194 | MANHATTAN |
| 10.864.886/0001-75 | 10625 | ARGENTINA |
| 29.273.778/0001-56 | 11189 | VILLE DE BEAUVAIS |
| 07.945.453/0001-30 | 10590 | LAGO MAGGIORE |

## Cobertura

### 1. Identificação pelo nome do arquivo (`buscar_por_nome_arquivo`)

| Caso | Entrada | Esperado |
|---|---|---|
| Código exato no nome | `PGR 10004 Klosters.pdf` | KLOSTERS (10004) |
| Código não confunde com data | `PGR 10004 Klosters 06.2026.pdf` | KLOSTERS (10004) |
| Fuzzy pelo nome | `ARAUJO LIMA QUITADO 05.26.pdf` | ARAUJO LIMA (10695) |
| Palavra de tipo de documento não atrapalha | `ARGENTINA QUITADO 05.26.pdf` | ARGENTINA (10625) |
| Ambiguidade protegida | `CONDE DE BONFIM.pdf` | nenhum (None) |
| Dois códigos válidos no nome | `PGR 10004 10002 Klosters.pdf` | nenhum (None) |
| Sem correspondência | `DOCUMENTO QUALQUER.pdf` | nenhum (None) |

O caso da ambiguidade é o mais importante da suíte: é a trava que impede
trocar 10490 por 10491.

### 2. Extração de CNPJ do conteúdo (`extrair_cnpj_tomador`)

| Caso | Fixture | Esperado |
|---|---|---|
| CO-ESTIPULANTE com hífen | `fedcorp_recibo_hifen.txt` | só SAN REMO |
| Intermediários ignorados | `fedcorp_recibo_hifen.txt` | nunca FedCorp (35.315.360/0001-67) nem Imodata (12.184.361/0001-14) |
| Emitente configurado ignorado | `nfse_ff.txt` com emitente = F&F | não devolve o CNPJ da F&F |
| NFS-e, campo TOMADOR (barra) | `nfse_ff.txt` | KLOSTERS |
| NFS-e de terceiro | `nfse_avulsa.txt` | LAGO MAGGIORE entre os candidatos |
| Nota com 2 CNPJs (emitente + tomador) | `nfse_imodata.txt` | devolve os dois; o desempate (seção 4) é quem resolve para MANHATTAN |
| Boleto, campo Pagador | `boleto_avulso.txt` | devolve os dois CNPJs, incluindo VILLE DE BEAUVAIS |
| Checksum inválido rejeitado | `cnpj_checksum_invalido.txt` | não devolve o CNPJ inválido |

Nos casos de dois CNPJs, o teste afirma o comportamento real: `extrair_cnpj_tomador`
devolve ambos (emitente não cadastrado + condomínio cadastrado) e quem escolhe é
o desempate, testado separadamente. Isso documenta a divisão de responsabilidade
entre as duas etapas.

### 3. Validação de CNPJ (`cnpj_valido`)

- CNPJs reais do cadastro de teste → válidos
- `00.001.208-2497-38` (caso real do NF-927) → inválido
- 14 dígitos iguais (`111...`) → inválido
- Menos de 14 dígitos → inválido

### 4. Desempate do Notas Diversas

- Dois candidatos, só um cadastrado → devolve só o cadastrado
- Dois candidatos, ambos cadastrados → continua ambíguo (não escolhe sozinho)
- Um candidato só → passa direto

### 5. Nome do arquivo de saída (`nome_saida_com_codigo`)

- `PGR 10004 Klosters.pdf` + `10004` → `10004 - PGR 10004 Klosters.pdf`
- Código com caractere inválido em nome de arquivo → sanitizado
- Código vazio → devolve o nome original

### 6. Configuração e predefinições (`carregar_config`)

- Sem `config.json` → 3 predefinições padrão, ativa = fedcorp
- `config.json` corrompido → volta aos padrões, não lança exceção
- Formato antigo (campos soltos na raiz) → migra para o perfil fedcorp,
  preservando os valores, sem afetar os outros perfis
- Salvar e recarregar → valores preservados

Estes testes usam um diretório temporário, nunca o `config.json` real.

## Alteração necessária no app

Uma só, mínima: a regra de desempate do Notas Diversas hoje está escrita dentro
de `_processar_em_thread`, e por isso não pode ser chamada de fora. Ela sai para
uma função de módulo:

```python
def desempatar_por_cadastro(candidatos, cadastro):
    """Entre vários CNPJs candidatos, prefere o único que já está cadastrado.
    Devolve a lista inalterada quando não há como desempatar."""
```

O loop passa a chamá-la quando `preferir_cadastrado_em_ambiguo` estiver ligado,
e detecta se a lista mudou para manter o sufixo "(desempate: CNPJ cadastrado)"
no log — **o formato do `processamento.log` não muda**.

Nenhuma outra função do app é alterada.

## Fora de escopo

- **Interface gráfica** (telas, botões, tema): testar GUI exige display e dá
  pouco retorno perto do custo. O risco de carimbar código errado não está lá.
- **OCR** (`extrair_texto_ocr`): depende do motor do Windows, é lento e não
  determinístico. Os fixtures de texto já cobrem o que o OCR alimenta.
- **Leitura de PDF de ponta a ponta**: os dois PDFs versionados (ARAUJO LIMA)
  são digitalizações sem texto nativo — extrair texto deles devolve vazio, então
  não servem de fixture. A leitura de PDF em si é responsabilidade do `pypdf`.
- **O `.exe` empacotado**: continua sendo verificado manualmente a cada build.

## Critério de sucesso

`rodar_testes.bat` (ou `python -m unittest discover -s tests`) passa em todos os
casos. Reverter qualquer uma das três correções feitas na v6.1/v6.2 faz pelo
menos um teste falhar — é assim que se confirma que a suíte tem valor.
