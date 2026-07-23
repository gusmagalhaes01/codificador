"""
Cadastro fixo usado só pelos testes — NUNCA a planilha real
(cadastro_condominios.xlsx). Congelado de propósito: se os testes lessem a
planilha real, cadastrar um condomínio novo poderia quebrar um teste sem que
nada esteja errado. Os 9 condomínios abaixo são os que apareceram nos bugs
reais corrigidos nas versões 6.1/6.2.

Chave = CNPJ normalizado (14 dígitos). Valor = {"codigo", "nome"}.
"""

CADASTRO_TESTE = {
    "40338774000141": {"codigo": "10695", "nome": "ARAUJO LIMA"},
    "01195716000154": {"codigo": "10004", "nome": "KLOSTERS"},
    "08578541000103": {"codigo": "10002", "nome": "SAN REMO"},
    "05695194000100": {"codigo": "10490", "nome": "CENTRO COM CONDE DE BONFIM RES"},
    "29361458000158": {"codigo": "10491", "nome": "CENTRO COM CONDE DE BONFIM"},
    "07448975000126": {"codigo": "11194", "nome": "MANHATTAN"},
    "10864886000175": {"codigo": "10625", "nome": "ARGENTINA"},
    "29273778000156": {"codigo": "11189", "nome": "VILLE DE BEAUVAIS"},
    "07945453000130": {"codigo": "10590", "nome": "LAGO MAGGIORE"},
}
