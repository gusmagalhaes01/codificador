"""
Script de uso único: preenche a coluna "ID SL" do cadastro a partir de um
export de condomínios ativos do Superlógica (arquivo com extensão .xls que
na verdade é HTML — é assim que o Superlógica exporta).

O cruzamento é feito pelo CÓDIGO do condomínio (coluna "ID" do export), não
pelo CNPJ: no export de 2026-08-20, vários CNPJs vinham malformados (8 deles
como 00.000.000/0000-00, outros com 8 dígitos), e cruzar por CNPJ casava só
632 dos 757 — contra 747 pelo código, sem nenhum conflito de nome.

Uso:  python tests/_importar_id_sl.py <export.xls> [--gravar]
Sem --gravar só mostra o que faria (teste seco).
"""

import os
import re
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

from logica import carregar_cadastro, salvar_cadastro


def ler_export(caminho):
    """dict código -> id_sl, a partir do HTML exportado pelo Superlógica."""
    with open(caminho, encoding="utf-8", errors="replace") as f:
        html = f.read()

    por_codigo = {}
    for linha in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.DOTALL | re.IGNORECASE):
        celulas = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", linha, re.DOTALL | re.IGNORECASE)
        valores = [re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ").strip() for c in celulas]
        if len(valores) < 2:
            continue
        codigo, id_sl = valores[0], valores[1]
        # a linha de cabeçalho ("ID ▲", "ID SL") não é numérica
        if codigo.isdigit() and id_sl.isdigit():
            por_codigo[codigo] = id_sl
    return por_codigo


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    caminho_export = sys.argv[1]
    gravar = "--gravar" in sys.argv
    caminho_cadastro = os.path.join(_RAIZ, "cadastro_condominios.xlsx")

    por_codigo = ler_export(caminho_export)
    cadastro = carregar_cadastro(caminho_cadastro)

    preenchidos, sem_match, ja_iguais = 0, [], 0
    for dados in cadastro.values():
        codigo = str(dados["codigo"]).strip()
        id_sl = por_codigo.get(codigo)
        if id_sl is None:
            sem_match.append((codigo, dados["nome"]))
            continue
        if dados.get("id_sl") == id_sl:
            ja_iguais += 1
            continue
        dados["id_sl"] = id_sl
        preenchidos += 1

    print(f"export: {len(por_codigo)} condomínios | cadastro: {len(cadastro)}")
    print(f"ID SL preenchido: {preenchidos}")
    if ja_iguais:
        print(f"já estavam corretos: {ja_iguais}")
    print(f"sem correspondência (ficam em branco): {len(sem_match)}")
    for codigo, nome in sorted(sem_match):
        print(f"  {codigo}  {nome}")

    if gravar:
        salvar_cadastro(caminho_cadastro, cadastro)
        print(f"\nGravado em {caminho_cadastro}")
    else:
        print("\n(teste seco — rode com --gravar para aplicar)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
