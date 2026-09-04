# -*- coding: utf-8 -*-
"""
Confere o discriminador contra os PDFs reais. Uso manual, reexecutável.

Não entra na suíte automatizada (prefixo _): depende de pastas que não
estão no repositório. Rode antes de liberar qualquer mudança no
discriminador.

    python tests/_validar_discriminador.py
"""
import os
import sys
import glob

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import logica

PASTA = r"C:\Users\Desktop\Desktop\Correios II\teste"

#  Pré-existente, e NÃO é falha do discriminador: extrair_dados_protocolo_
#  correio também não acha o marcador neste arquivo, nem a 300 nem a 400
#  DPI. Ele já é pendente hoje, antes desta funcionalidade existir.
ILEGIVEIS_CONHECIDOS = {"doc16004820260821155206.pdf"}

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)


def classificar(caminho):
    """Como a aba 3 classifica: leitura normal e, se não decidir, uma
    segunda chance no recorte do topo — onde vive o cabeçalho dos dois
    formatos. Sem ela o telnet-010 não é reconhecido."""
    texto = logica.extrair_texto_pdf(caminho)
    if len(texto.strip()) < logica.LIMITE_TEXTO_MINIMO:
        texto = logica.extrair_texto_escaneado(caminho, dpi=300)
    formato = logica.formato_do_protocolo(texto)
    if formato is None:
        try:
            formato = logica.formato_do_protocolo(
                logica.extrair_texto_ocr_regiao(caminho, (0, 0, 1, 0.35), dpi=400))
        except Exception:
            pass
    return formato


erros = 0
for caminho in sorted(glob.glob(os.path.join(PASTA, "*.pdf"))):
    nome = os.path.basename(caminho)
    if os.path.getsize(caminho) == 0:
        print(f"  {nome:34} (0 byte, pulado)")
        continue

    esperado = "telnet" if "telnet" in nome.lower() else "novo"
    obtido = classificar(caminho)

    if obtido is None and nome in ILEGIVEIS_CONHECIDOS:
        print(f"  {nome:34} None     (ilegível conhecido, pré-existente)")
        continue

    #  O que NÃO pode acontecer é trocar um formato pelo outro. None é
    #  falha segura: o arquivo vira pendente em vez de receber a extração
    #  errada.
    if obtido is not None and obtido != esperado:
        erros += 1
        print(f"  {nome:34} obtido={obtido:8} esperado={esperado:8} <<< TROCOU O FORMATO")
    elif obtido is None:
        print(f"  {nome:34} None     esperado={esperado:8} <<< nao reconheceu (falha segura)")
    else:
        print(f"  {nome:34} {obtido:8} ok")

print(f"\n{'NENHUM FORMATO TROCADO' if not erros else f'{erros} FORMATOS TROCADOS'}")
sys.exit(1 if erros else 0)
