# -*- coding: utf-8 -*-
"""
Valida o lançamento no Superlógica contra as pastas reais.

Fora da suíte (prefixo `_`): depende de PDFs fora do repositório.
Reexecutável. Não grava nada — só monta os lançamentos e confere.

Rode:  python tests/_validar_lancamento.py
"""
import collections
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logica

if logica.FITZ_DISPONIVEL:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
CADASTRO = logica.carregar_cadastro("cadastro_condominios.xlsx")
CODIGO_PARA_DOC = logica._codigos_do_cadastro(CADASTRO)

LOTES = [
    ("Sindicato", r"C:\Users\Desktop\Music\sindicato\boletos",
     lambda nome: CODIGO_PARA_DOC.get((re.match(r"(\d{5})_", nome) or [None, None])[1])),
    ("DARF", r"C:\Users\Desktop\Music\inss",
     lambda nome: (re.search(r"GuiaPagamento_(\d{14})_", nome) or [None, None])[1]),
    ("F&F fevereiro", r"C:\Users\Desktop\Downloads\02 Fevereiro 2026-20260401T035105Z-1-001",
     lambda nome: CODIGO_PARA_DOC.get((re.search(r"\b(1\d{4})\b", nome) or [None, None])[1])),
]


def documentos_do_lote(pasta):
    documentos = []
    for raiz, _, arquivos in os.walk(pasta):
        for nome in sorted(arquivos):
            if not nome.lower().endswith(".pdf"):
                continue
            caminho = os.path.join(raiz, nome)
            try:
                texto = logica.extrair_texto_pdf(caminho)
            except Exception:
                continue
            nfse = logica.extrair_dados_nfse(texto)
            boleto = None if nfse else logica.extrair_dados_boleto(texto, CADASTRO)
            doc = logica.documento_para_lancamento(nome, nfse, boleto, texto)
            if doc:
                documentos.append(doc)
    return documentos


for rotulo, pasta, esperado_do_nome in LOTES:
    if not os.path.isdir(pasta):
        print(f"\n== {rotulo}: pasta não encontrada ({pasta})")
        continue
    docs = documentos_do_lote(pasta)
    motivos = collections.Counter(logica.motivo_fora_do_lancamento(d, CADASTRO) or "(lançado)"
                                  for d in docs)
    mistura = logica.conferir_lote_homogeneo(docs)
    vencimento_lote = datetime.date(2026, 9, 5)
    lancamentos = logica.montar_lancamentos(docs, CADASTRO, vencimento_lote) if not mistura else []
    divergentes = [d["arquivo"] for d in docs
                   if not logica.motivo_fora_do_lancamento(d, CADASTRO)
                   and esperado_do_nome(d["arquivo"])
                   and esperado_do_nome(d["arquivo"]) != d["documento"]]
    print(f"\n== {rotulo}: {len(docs)} documentos reconhecidos")
    print(f"   pasta misturada ........ {'SIM: ' + mistura.splitlines()[0] if mistura else 'não'}")
    for motivo, n in motivos.most_common():
        print(f"   {motivo:55} {n}")
    print(f"   lançamentos montados ... {len(lancamentos)}")
    print(f"   CONDOMÍNIO DIVERGE DO NOME DO ARQUIVO: {len(divergentes)}   <- tem que ser 0")
    for nome in divergentes[:5]:
        print(f"      ! {nome}")
