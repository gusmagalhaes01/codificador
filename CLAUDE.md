# Identificação de PDFs por CNPJ — Contexto do Projeto

## O que o programa faz

App desktop (Tkinter, Windows) que processa boletos e notas fiscais (NFS-e) em PDF,
identifica a qual condomínio cada documento pertence (via CNPJ do tomador) e escreve
o código do condomínio no PDF (rodapé ou canto superior), comparando com um cadastro
em planilha `.xlsx`.

Arquivo principal: `identificacao_por_cnpj_5_3.py`
Cadastro de condomínios: `cadastro_condominios.xlsx` (colunas: CNPJ, Código, Nome)
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

- **Fix (versão 5_3)**: adicionado `CNPJS_INTERMEDIARIOS` para excluir FedCorp e
  Imodata também no fallback genérico, não só no filtro por campo semântico.
  Corrigiu caso onde o OCR não lia o rótulo `CO-ESTIPULANTE` e o script pegava
  o CNPJ da Imodata por engano.

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

## Ideia de melhoria discutida, não implementada

- **LLM como último fallback** (não como motor principal): mandar a imagem da
  página pra API do Claude só quando as etapas acima falharem. Mais robusto a
  variação de layout, mas custa por documento e envolve enviar dados de
  clientes pra fora — decisão de não implementar por ora.

## Convenções do código

- Comentários e nomes de variáveis em português (ex: `cadastro`, `formatar_cnpj`).
- CNPJs sempre normalizados para 14 dígitos internamente (`normalizar_cnpj`),
  formatados só na exibição/planilha (`formatar_cnpj`).
- Dependências: `pypdf`, `reportlab`, `openpyxl`, `pymupdf`, `winocr` (Windows).
