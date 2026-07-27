"""
Gera os fixtures de texto em tests/dados/ a partir dos PDFs reais.
Uso único: `python tests/_gerar_fixtures.py`. Os PDFs-fonte não vão para o
repositório; só os .txt resultantes. Ajuste FONTES se regenerar noutra máquina.
"""
import os
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _RAIZ)
from logica import extrair_texto_pdf

FONTES = {
    "fedcorp_recibo_hifen.txt": r"C:\Users\Dell\Downloads\BENEFICIO 07-2026 BOLETO\NF-669.pdf",
    "nfse_ff.txt": r"C:\Users\Dell\Documents\NOTAS FF\FF JUNHO 2026\06 Junho\originais\PGR\PGR 10004 Klosters.pdf",
    "nfse_imodata.txt": r"C:\Users\Dell\Downloads\nfses_pdfs_2026-05-04\NFSE_284726_11194.pdf",
    "nfse_avulsa.txt": r"C:\Users\Dell\Documents\EXEMPLO NF.pdf",
    "boleto_avulso.txt": r"C:\Users\Dell\Documents\EXEMPLO BOLETO.pdf",
    "cnpj_checksum_invalido.txt": r"C:\Users\Dell\Downloads\BENEFICIO 07-2026 BOLETO\NF-927.pdf",
}

destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados")
os.makedirs(destino, exist_ok=True)

for nome_txt, caminho_pdf in FONTES.items():
    if not os.path.isfile(caminho_pdf):
        print(f"AVISO: fonte não encontrada, pulando: {caminho_pdf}")
        continue
    texto = extrair_texto_pdf(caminho_pdf)
    with open(os.path.join(destino, nome_txt), "w", encoding="utf-8") as f:
        f.write(texto)
    print(f"gerado: dados/{nome_txt} ({len(texto)} chars)")
