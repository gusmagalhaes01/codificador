"""
Cadastro fixo usado só pelos testes — NUNCA a planilha real
(cadastro_condominios.xlsx). Congelado de propósito: se os testes lessem a
planilha real, cadastrar um condomínio novo poderia quebrar um teste sem que
nada esteja errado. Os 9 condomínios abaixo são os que apareceram nos bugs
reais corrigidos nas versões 6.1/6.2.

Chave = CNPJ normalizado (14 dígitos). Valor = {"codigo", "nome", "id_sl"}.

"id_sl" é o código do condomínio no Superlógica, usado para gerar a planilha
de despesas. LAGO MAGGIORE está sem de propósito: na planilha real 10 dos 764
condomínios não têm esse campo, e o programa precisa lidar com isso.
"""

CADASTRO_TESTE = {
    "40338774000141": {"codigo": "10695", "nome": "ARAUJO LIMA", "id_sl": "701"},
    "01195716000154": {"codigo": "10004", "nome": "KLOSTERS", "id_sl": "44"},
    "08578541000103": {"codigo": "10002", "nome": "SAN REMO", "id_sl": "42"},
    "05695194000100": {"codigo": "10490", "nome": "CENTRO COM CONDE DE BONFIM RES", "id_sl": "487"},
    "29361458000158": {"codigo": "10491", "nome": "CENTRO COM CONDE DE BONFIM", "id_sl": "488"},
    "07448975000126": {"codigo": "11194", "nome": "MANHATTAN", "id_sl": "912"},
    "10864886000175": {"codigo": "10625", "nome": "ARGENTINA", "id_sl": "633"},
    "29273778000156": {"codigo": "11189", "nome": "VILLE DE BEAUVAIS", "id_sl": "907"},
    "07945453000130": {"codigo": "10590", "nome": "LAGO MAGGIORE", "id_sl": ""},
}
