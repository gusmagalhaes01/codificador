# -*- coding: utf-8 -*-
"""
Confere a apuracao do telnet contra os 7 PDFs reais de referencia.

Uso manual, reexecutavel, fora da suite (prefixo _): depende de uma pasta
que nao esta no repositorio. A verdade abaixo foi lida a mao na folha de
cada documento.

    python tests/_validar_telnet.py
"""
import os
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

PASTA = r"C:\Users\Desktop\Desktop\Correios\telnet"

#  arquivo -> (total impresso, codigo), conferidos lendo a folha
VERDADE = {
    "procolo telnet-007.pdf": (27, "11122"),
    "procolo telnet-008.pdf": (8,  "11127"),
    "procolo telnet-009.pdf": (10, "11121"),
    "procolo telnet-010.pdf": (8,  "10151"),
    "procolo telnet-011.pdf": (3,  "11100"),
    "procolo telnet-012.pdf": (5,  "11098"),
    "procolo telnet-013.pdf": (1,  "11093"),
}

REGIOES = [(None, 300), ((0, 0, 1, 0.35), 400), ((0, 0, 1, 0.32), 500)]
cadastro = logica.carregar_cadastro(os.path.join(_RAIZ, "cadastro_condominios.xlsx"))

print(f"{'arquivo':26} {'cod':>6} {'esp':>6} {'tot':>5} {'esp':>5} {'nome':>5}  resultado")
print("-" * 82)
codigos_ok = totais_ok = errados = 0
for nome, (total_certo, codigo_certo) in VERDADE.items():
    caminho = os.path.join(PASTA, nome)
    leituras = []
    for regiao, dpi in REGIOES:
        try:
            leituras.append(logica.extrair_texto_escaneado(caminho, dpi=dpi)
                            if regiao is None
                            else logica.extrair_texto_ocr_regiao(caminho, regiao, dpi=dpi))
        except Exception:
            pass
    dados = logica.apurar_protocolo_telnet(leituras, cadastro) or {}
    aceito, observacao = logica.conferir_contagem_telnet(dados)

    cod, tot = dados.get("codigo"), dados.get("total_impresso")
    codigos_ok += (cod == codigo_certo)
    totais_ok += (tot == total_certo)
    #  O que NAO pode acontecer: preencher um valor errado.
    if aceito and tot != total_certo:
        errados += 1
    if cod is not None and cod != codigo_certo:
        errados += 1
    print(f"{nome[-12:]:26} {str(cod):>6} {codigo_certo:>6} {str(tot):>5} "
          f"{total_certo:>5} {('sim' if dados.get('nome_confere') else 'nao'):>5}  "
          f"{observacao or 'preenche'}")

print(f"\ncodigos corretos: {codigos_ok}/7   totais corretos: {totais_ok}/7")
print(f"VALORES ERRADOS PREENCHIDOS: {errados}   <- tem que ser 0")
sys.exit(1 if errados else 0)
