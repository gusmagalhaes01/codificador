# Carimbo lateral rotacionado pro Protocolo dos Correios — especificação

Data: 2026-08-13
Estado: aprovado, a implementar

## Objetivo

Documentos identificados via `extrair_codigo_protocolo_correio` (Protocolo de
Recebimento de Documento, Correios/Imodata) recebem um carimbo diferente do
rodapé/canto superior normais: uma linha única — `"{código} {nome} -
{CNPJ formatado}"` — rotacionada 90° (sentido anti-horário, lê de baixo pra
cima), colada perto da margem **direita**, verticalmente centralizada na
página. Automático só pra esse tipo de documento — não depende de
`modo_texto` nem de nenhuma configuração que o usuário escolha.

Mockup aprovado pelo usuário (ver conversa): texto rotacionado a ~20pt da
borda direita, mesma cor/fonte já configuradas (`self.cor_texto`,
`self.tamanho_fonte`).

## Decisão técnica

`criar_overlay()` (`logica.py`) ganha um parâmetro opcional `angulo=0` —
comportamento atual preservado pra todo mundo que já chama sem esse
parâmetro. Quando `angulo == 90`:
- calcula a largura do texto (`stringWidth`) pra centralizar verticalmente:
  `y = (altura - largura_texto) / 2`;
- posiciona a `margem` pontos da borda direita: `x = largura - margem`;
- usa `c.translate(x, y)` + `c.rotate(90)` + `c.drawString(0, 0, texto)`.

`margem = 20` (pontos), fixo no código — não é uma opção de configuração
(YAGNI: só esse tipo de documento usa, sem pedido de ajuste fino).

Nova função `montar_texto_protocolo_correio(codigo, nome, cnpj)` em
`logica.py`, só formata a string — sem lógica de posicionamento.

## Onde entra no fluxo

Em `_processar_em_thread`, no ramo que já resolve `cnpj_do_protocolo` (ver
spec anterior, `2026-08-13-protocolo-correio-design.md`): quando o registro
existe no cadastro, monta `texto_pdf` com `montar_texto_protocolo_correio`
em vez do formato de `modo_texto`, e chama `processar_pdf` com
`config["angulo"] = 90` (as outras chaves de `config` — fonte, tamanho,
cor — continuam vindo da configuração normal do usuário).

## Testes

- `criar_overlay` com `angulo=90`: o texto ainda é extraível do PDF gerado
  via `extrair_texto_pdf` (pypdf lida bem com texto rotacionado) — cobre o
  carimbo em si sem precisar inspecionar pixels.
- `criar_overlay` com `angulo=0` (padrão): comportamento idêntico ao de
  antes — nenhuma regressão nos modos rodapé/canto superior existentes.
- `montar_texto_protocolo_correio`: formata corretamente código + nome + CNPJ.

## Verificação visual

Depois de implementado, gero um PDF de teste real carimbado e peço pro
usuário abrir e confirmar visualmente — texto rotacionado em PDF é sensível
a acerto fino que só a extração de texto não garante.
