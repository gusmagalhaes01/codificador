# -*- coding: utf-8 -*-
"""
Valida o leitor do "Meus Correios" contra os PDFs reais.

Fora da suíte (prefixo `_`): depende de arquivos que não estão no
repositório. Reexecutável.

Rode:  python tests/_validar_meus_correios.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

PASTA = r"C:\Users\Desktop\Desktop\Correios\telnet"
#  Os 6 "Meus Correios" que convivem com os 7 telnet nessa pasta.
ESPERADO = {
    "procolo telnet-002.pdf": "10852",
    "procolo telnet-003.pdf": "10520",
    "procolo telnet-004.pdf": "11124",
    "procolo telnet-005.pdf": "10193",
    "procolo telnet-006.pdf": "10255",
    "procolo telnet.pdf": "11124",
}

cadastro = logica.carregar_cadastro("cadastro_condominios.xlsx")

print(f"{'arquivo':26} {'formato':14} {'cod':>6} {'esp':>6} {'nome':5} servico")
print("-" * 96)
certos = trocados = 0
for nome, esperado in sorted(ESPERADO.items()):
    caminho = os.path.join(PASTA, nome)
    texto = logica.extrair_texto_ocr_regiao(caminho, (0, 0, 1, 0.45), dpi=400)
    formato = logica.classificar_formato_protocolo(texto)
    if formato != "meus_correios":
        trocados += 1
        print(f"{nome[-24:]:26} {str(formato):14} FORMATO TROCADO")
        continue
    d = logica.extrair_dados_meus_correios(texto, cadastro)
    ok = d["codigo"] == esperado
    certos += ok
    print(f"{nome[-24:]:26} {formato:14} {str(d['codigo']):>6} {esperado:>6} "
          f"{('sim' if d['nome_confere'] else 'nao'):5} {d['servico']}")

print(f"\ncodigos corretos: {certos}/{len(ESPERADO)}")
print(f"FORMATOS TROCADOS: {trocados}   <- tem que ser 0")
