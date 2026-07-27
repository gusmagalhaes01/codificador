"""
Lógica de negócio do Codificador — extração/validação de CNPJ, identificação
por nome de arquivo, configuração persistente, leitura de PDF/OCR, carimbo no
PDF e persistência da planilha de cadastro. Nada de interface aqui (sem
tkinter/customtkinter) — é a parte coberta pelos testes em tests/.

A classe App (interface CustomTkinter) fica em identificacao_por_cnpj_6_0.py,
que importa deste módulo.
"""

import asyncio
import copy
import datetime
import io
import json
import os
import re
import sys
import unicodedata
from difflib import SequenceMatcher

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from openpyxl import Workbook, load_workbook

# --- PyMuPDF para renderizar páginas como imagem ---
try:
    import fitz
    from PIL import Image
    FITZ_DISPONIVEL = True
except ImportError:
    FITZ_DISPONIVEL = False

# --- winocr: OCR nativo do Windows, sem programas externos ---
try:
    import winocr
    WINOCR_DISPONIVEL = True
except ImportError:
    WINOCR_DISPONIVEL = False

OCR_DISPONIVEL = FITZ_DISPONIVEL and WINOCR_DISPONIVEL


# ============================================================
#  CONSTANTES
# ============================================================

CNPJ_REGEX = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
NOME_ARQUIVO_PADRAO = "cadastro_condominios.xlsx"
NOME_LOG_PADRAO = "processamento.log"
NOME_CONFIG_PADRAO = "config.json"


def pasta_base():
    """
    Pasta onde ficam config.json, os logs e a planilha de cadastro.
    Rodando como script (.py), é a pasta do próprio arquivo. Empacotado com
    PyInstaller (--onedir/--onefile), sys.frozen fica True e __file__ aponta
    pra dentro da pasta temporária/interna do pacote — nesse caso usamos a
    pasta onde está o .exe, para tudo ficar ao lado dele (portátil).
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def caminho_recurso(nome):
    """
    Caminho de um recurso EMBUTIDO só-leitura (ex: o ícone .ico), diferente de
    pasta_base() que é para arquivos graváveis ao lado do .exe. Empacotado com
    PyInstaller, os dados vão para sys._MEIPASS; rodando como script, ficam na
    pasta do próprio arquivo.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome)


NOME_ICONE = "icone.ico"


# ============================================================
#  CONFIGURAÇÃO PERSISTENTE (config.json, ao lado do script)
# ============================================================

# Campos que vivem DENTRO de cada predefinição (perfil de lote). Cada perfil é
# um "pacote" de configurações que o usuário troca de uma vez (ex: "hoje vou
# processar boletos da FedCorp" vs "hoje é a F&F"), em vez de reconfigurar tudo
# a cada lote.
CHAVES_PERFIL = [
    "cnpj_emitente", "preferir_cadastrado_em_ambiguo",
    "usar_match_nome_arquivo", "usar_ocr", "dpi_ocr", "usar_ocr_regiao",
    "regiao_x0", "regiao_y0", "regiao_x1", "regiao_y1",
    "modo_texto", "tipo_servico", "tamanho_fonte", "cor_texto",
    "renomear_com_codigo",
]

# Todos os perfis usam o mesmo pipeline de identificação
# (buscar_por_nome_arquivo -> texto/OCR/CNPJ do conteúdo, ver
# _processar_em_thread); o que muda de um perfil pro outro é só a
# configuração: CNPJ emitente, se tenta casar pelo nome do arquivo primeiro,
# como carimba o PDF, e (Notas Diversas) o desempate de CNPJ ambíguo.
PERFIS_PADRAO = {
    "fedcorp": {
        "nome": "FedCorp",
        "cnpj_emitente": "35.315.360/0001-67",
        "preferir_cadastrado_em_ambiguo": False,
        "usar_match_nome_arquivo": True,
        "usar_ocr": False,
        "dpi_ocr": 200,
        "usar_ocr_regiao": False,
        "regiao_x0": 0.0, "regiao_y0": 0.15, "regiao_x1": 1.0, "regiao_y1": 0.35,
        "modo_texto": "rodape",
        "tipo_servico": "CIPAA",
        "tamanho_fonte": "14",
        "cor_texto": "#000000",
        "renomear_com_codigo": False,
    },
    "ff": {
        "nome": "F&F",
        "cnpj_emitente": "13.736.666/0001-54",
        "preferir_cadastrado_em_ambiguo": False,
        # F&F já nomeia os arquivos com o código embutido (ex: "PGR 10004
        # Klosters.pdf"), então o match por nome do arquivo (que também
        # reconhece código exato — ver buscar_por_nome_arquivo) resolve
        # praticamente tudo sem precisar abrir o PDF.
        "usar_match_nome_arquivo": True,
        "usar_ocr": False,
        "dpi_ocr": 200,
        "usar_ocr_regiao": False,
        "regiao_x0": 0.0, "regiao_y0": 0.15, "regiao_x1": 1.0, "regiao_y1": 0.35,
        "modo_texto": "rodape",
        "tipo_servico": "PGR",
        "tamanho_fonte": "14",
        "cor_texto": "#000000",
        "renomear_com_codigo": False,
    },
    "notas_diversas": {
        "nome": "Notas Diversas",
        "cnpj_emitente": "",  # sem emitente fixo — lotes de administradoras variadas
        "preferir_cadastrado_em_ambiguo": True,
        "usar_match_nome_arquivo": False,
        "usar_ocr": False,
        "dpi_ocr": 200,
        "usar_ocr_regiao": False,
        "regiao_x0": 0.0, "regiao_y0": 0.15, "regiao_x1": 1.0, "regiao_y1": 0.35,
        "modo_texto": "topo_esquerdo",
        "tipo_servico": "",
        "tamanho_fonte": "14",
        "cor_texto": "#000000",
        "renomear_com_codigo": False,
    },
}

ORDEM_PREDEFINICOES = ["fedcorp", "ff", "notas_diversas"]

DEFAULTS_CONFIG = {
    "predefinicao_ativa": "fedcorp",
    "predefinicoes": PERFIS_PADRAO,
    "tema": "auto",  # "auto" | "claro" | "escuro"
}


def nome_saida_com_codigo(nome_original, codigo):
    """Devolve o nome do arquivo de saída prefixado com o código do condomínio,
    ex: "10002 - ARAUJO LIMA QUITADO 05.26.pdf". Remove do código caracteres
    inválidos em nome de arquivo no Windows."""
    codigo_limpo = re.sub(r'[\\/:*?"<>|]', "", str(codigo)).strip()
    if not codigo_limpo:
        return nome_original
    return f"{codigo_limpo} - {nome_original}"


def _gravar_erros_log(detalhes):
    """Faz o append de `detalhes` em erros.log (ao lado do script), com
    timestamp. Usada tanto por _registrar_erro_config quanto pelo handler
    global de exceções (App.report_callback_exception) — mesmo formato nos
    dois casos. Nunca lança exceção."""
    try:
        pasta_script = pasta_base()
        with open(os.path.join(pasta_script, "erros.log"), "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}]\n{detalhes}\n")
    except Exception:
        pass


def _registrar_erro_config(detalhes):
    """Grava uma falha de leitura/escrita de config.json em erros.log, no
    mesmo formato usado pelo handler global de exceções (App.report_callback_exception)."""
    _gravar_erros_log(detalhes)


def carregar_config():
    """
    Lê config.json (ao lado do script). Se não existir, devolve uma cópia
    dos defaults (3 predefinições). Se existir mas estiver corrompido, loga a
    exceção em erros.log e devolve uma cópia dos defaults — nunca lança
    exceção. Sempre mescla sobre os defaults, então um config parcial (de
    uma versão antiga, sem alguma chave nova) não quebra.

    Migração: versões anteriores ao recurso de predefinições salvavam os
    campos de lote soltos no nível raiz do config (ex: "cnpj_emitente" direto,
    sem "predefinicoes"). Quando isso é detectado, esses valores viram a base
    do perfil "fedcorp" — era o único perfil que existia até então — sem
    perder a configuração que o usuário já tinha.
    """
    config = copy.deepcopy(DEFAULTS_CONFIG)

    pasta_script = pasta_base()
    caminho = os.path.join(pasta_script, NOME_CONFIG_PADRAO)
    if not os.path.isfile(caminho):
        return config

    try:
        with open(caminho, "r", encoding="utf-8") as f:
            dados = json.load(f)
    except Exception:
        import traceback
        _registrar_erro_config(traceback.format_exc())
        return copy.deepcopy(DEFAULTS_CONFIG)

    if not isinstance(dados, dict):
        return config

    if "tema" in dados:
        config["tema"] = dados["tema"]

    if "predefinicoes" not in dados:
        # Formato antigo (campos de lote soltos no nível raiz) -> migra pro
        # perfil FedCorp, que era o único configurável até então.
        for chave in CHAVES_PERFIL:
            if chave in dados:
                config["predefinicoes"]["fedcorp"][chave] = dados[chave]
        return config

    # Formato novo: mescla cada predefinição salva sobre o default dela
    # (preserva chaves novas que uma versão anterior do config não tinha).
    predefinicoes_salvas = dados.get("predefinicoes")
    if isinstance(predefinicoes_salvas, dict):
        for chave_perfil, perfil_default in config["predefinicoes"].items():
            perfil_salvo = predefinicoes_salvas.get(chave_perfil)
            if isinstance(perfil_salvo, dict):
                perfil_default.update(perfil_salvo)

    ativa = dados.get("predefinicao_ativa")
    if ativa in config["predefinicoes"]:
        config["predefinicao_ativa"] = ativa

    return config


def salvar_config(config):
    """Grava `config` como JSON em config.json (ao lado do script). Falha ao
    salvar não deve travar o app — loga em erros.log e segue."""
    pasta_script = pasta_base()
    caminho = os.path.join(pasta_script, NOME_CONFIG_PADRAO)
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        import traceback
        _registrar_erro_config(traceback.format_exc())


# ============================================================
#  FUNÇÕES DE CNPJ E TEXTO
# ============================================================

def normalizar_cnpj(cnpj):
    """Remove pontuação, deixando só os dígitos."""
    return re.sub(r"\D", "", cnpj)


def formatar_cnpj(cnpj_normalizado):
    """Formata 14 dígitos como 00.000.000/0000-00 (se possível)."""
    d = re.sub(r"\D", "", cnpj_normalizado)
    if len(d) == 14:
        return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"
    return cnpj_normalizado


def extrair_texto_pdf(caminho):
    reader = PdfReader(caminho)
    return "\n".join(pagina.extract_text() or "" for pagina in reader.pages)


def cnpj_valido(cnpj_normalizado):
    """
    Valida os dígitos verificadores de um CNPJ (14 dígitos).
    Usado para descartar leituras de OCR que "parecem" um CNPJ (14 dígitos)
    mas têm algum caractere errado — o checksum não bate.
    """
    d = cnpj_normalizado
    if len(d) != 14 or d == d[0] * 14:
        return False

    def _digito(nums, pesos):
        soma = sum(int(n) * p for n, p in zip(nums, pesos))
        resto = soma % 11
        return "0" if resto < 2 else str(11 - resto)

    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

    dv1 = _digito(d[:12], pesos1)
    dv2 = _digito(d[:12] + dv1, pesos2)
    return d[12] == dv1 and d[13] == dv2


LIMITE_TEXTO_MINIMO = 30  # abaixo disso, consideramos que o PDF não tem texto legível


def extrair_texto_ocr(caminho, max_paginas=2, dpi=300):
    """
    Renderiza as primeiras páginas do PDF como imagem e roda OCR usando
    o motor nativo do Windows (winocr) — sem programas externos instalados.
    Tenta português (pt-BR) primeiro; se não disponível, usa inglês (en-US).
    """
    if not OCR_DISPONIVEL:
        raise RuntimeError("OCR não disponível. Instale: pip install pymupdf winocr")

    textos = []
    doc = fitz.open(caminho)
    try:
        for i, pagina in enumerate(doc):
            if i >= max_paginas:
                break
            pix = pagina.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # winocr usa async — rodamos dentro de um event loop síncrono
            async def _ocr(pil_img):
                try:
                    resultado = await winocr.recognize_pil(pil_img, "pt-BR")
                    return resultado.text
                except Exception:
                    # Idioma pt-BR não instalado no Windows → tenta inglês
                    try:
                        resultado = await winocr.recognize_pil(pil_img, "en-US")
                        return resultado.text
                    except Exception:
                        return ""

            textos.append(asyncio.run(_ocr(img)))
    finally:
        doc.close()
    return "\n".join(textos)


def extrair_texto_ocr_regiao(caminho, retangulo, dpi=300):
    """
    Roda OCR só num recorte retangular da 1ª página (mais rápido e evita
    interferência de carimbos como "QUITADO" fora da área de interesse).
    `retangulo` = (x0, y0, x1, y1), frações de 0.0 a 1.0 da largura/altura
    da página (ex: topo da página = y0 pequeno, y1 pequeno).
    """
    if not OCR_DISPONIVEL:
        raise RuntimeError("OCR não disponível. Instale: pip install pymupdf winocr")

    doc = fitz.open(caminho)
    try:
        pagina = doc[0]
        largura_pt, altura_pt = pagina.rect.width, pagina.rect.height
        x0, y0, x1, y1 = retangulo
        clip = fitz.Rect(x0 * largura_pt, y0 * altura_pt, x1 * largura_pt, y1 * altura_pt)
        pix = pagina.get_pixmap(dpi=dpi, clip=clip)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        async def _ocr(pil_img):
            try:
                resultado = await winocr.recognize_pil(pil_img, "pt-BR")
                return resultado.text
            except Exception:
                try:
                    resultado = await winocr.recognize_pil(pil_img, "en-US")
                    return resultado.text
                except Exception:
                    return ""

        return asyncio.run(_ocr(img))
    finally:
        doc.close()


DPI_ESCALONAMENTO = [72, 150, 200, 300, 400]  # níveis usados na re-tentativa automática de OCR


def proximo_dpi_maior(dpi_atual):
    """Retorna o próximo nível de DPI acima do atual, ou None se já está no máximo."""
    for t in DPI_ESCALONAMENTO:
        if t > dpi_atual:
            return t
    return None


#  CNPJs que aparecem nos boletos/faturas mas NUNCA são o condomínio
#  tomador — intermediários que sempre devem ser ignorados como candidatos,
#  mesmo quando o OCR não consegue ler o rótulo do campo (ex: CO-ESTIPULANTE).
CNPJS_INTERMEDIARIOS = {
    "35315360000167": "FedCorp Administradora (emitente)",
    "12184361000114": "Imodata Empreendimentos (estipulante)",
}


def extrair_cnpj_tomador(texto, cnpj_emitente_normalizado):
    """
    Retorna lista de CNPJs candidatos (normalizados) encontrados no texto,
    excluindo o CNPJ da empresa emitente (que se repete em todo documento)
    e os CNPJs de intermediários conhecidos (ex: Imodata).

    Estratégia em dois passos:
      1. Tenta extrair CNPJ de campos semânticos explícitos (CO-ESTIPULANTE,
         TOMADOR, PAGADOR, EMPREGADOR...). Se achar, retorna só esses.
      2. Se nada nos campos explícitos, faz varredura genérica — comportamento
         original para NFS-e e documentos sem rótulos padronizados. Mesmo aqui,
         CNPJs de intermediários conhecidos são descartados, para cobrir os
         casos em que o OCR não capta o rótulo do campo (ex: "CO-ESTIPULANTE"
         sai ilegível, sobrando só os nomes/CNPJs soltos no texto).
    """
    # Tolerante a variações de separador: alguns recibos FedCorp escrevem o CNPJ do
    # co-estipulante com "-" (ou ".") no lugar da "/" — ex: 08.578.541-0001-03. O
    # separador antes do bloco 0001 e antes dos 2 dígitos finais aceita /, -, . ou espaço.
    CNPJ_FLEX = r"(\d{2}[\s.]?\d{3}[\s.]?\d{3}[\s/.\-]?\d{4}[\s.\-]?\d{2})"

    cnpjs_a_ignorar = set(CNPJS_INTERMEDIARIOS.keys())
    cnpjs_a_ignorar.add(cnpj_emitente_normalizado)

    PADROES_CAMPO = [
        # Boleto/recibo FedCorp: CO-ESTIPULANTE ... CNPJ: xx.xxx.xxx/xxxx-xx
        r"CO[-\s]?ESTIPULANTE[:\s]+.{0,120}?CNPJ[:\s]*" + CNPJ_FLEX,
        # Boleto genérico: linha Pagador ... CNPJ/CPF: xxxxxxxxxxxxxxx
        r"PAGADOR[:\s]+.{0,150}?CNPJ[/\s]?CPF[:\s]*" + CNPJ_FLEX,
        # NFS-e: TOMADOR ... CNPJ: xx.xxx.xxx/xxxx-xx
        r"TOMADOR[:\s]+.{0,120}?CNPJ[:\s]*" + CNPJ_FLEX,
        # Detalhamento de faturamento: EMPREGADOR: nome (CNPJ xx...)
        r"EMPREGADOR[:\s]+.{0,120}?\(CNPJ[:\s]*" + CNPJ_FLEX + r"\)",
        # Contratos genéricos: CONTRATANTE / CLIENTE ... CNPJ
        r"(?:CONTRATANTE|CLIENTE)[:\s]+.{0,100}?CNPJ[:\s]*" + CNPJ_FLEX,
    ]

    candidatos_campo = []
    for padrao in PADROES_CAMPO:
        for m in re.finditer(padrao, texto, re.IGNORECASE | re.DOTALL):
            cnpj_norm = normalizar_cnpj(m.group(1))
            if (len(cnpj_norm) == 14 and cnpj_valido(cnpj_norm)
                    and cnpj_norm not in cnpjs_a_ignorar):
                if cnpj_norm not in candidatos_campo:
                    candidatos_campo.append(cnpj_norm)

    if candidatos_campo:
        return candidatos_campo

    # --- Fallback: varredura genérica (comportamento original) ---
    encontrados = CNPJ_REGEX.findall(texto)
    normalizados = [normalizar_cnpj(c) for c in encontrados]
    candidatos = [c for c in normalizados if cnpj_valido(c) and c not in cnpjs_a_ignorar]
    vistos = set()
    unicos = []
    for c in candidatos:
        if c not in vistos:
            vistos.add(c)
            unicos.append(c)
    return unicos


def sugerir_nome_condominio(texto):
    """Tenta extrair o nome do condomínio/tomador de vários formatos de documento."""

    # Boleto/recibo FedCorp: CO-ESTIPULANTE:  NOME DO COND    CNPJ: ...
    m = re.search(r"CO[-\s]?ESTIPULANTE[:\s]+([^\n]+?)\s+CNPJ[:\s]", texto, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Boleto genérico: linha "Pagador   CONDOMINIO DO EDIFICIO X   CNPJ/CPF: ..."
    m = re.search(r"Pagador\s+([^\n]+?)\s+CNPJ[/\s]?CPF", texto, re.IGNORECASE)
    if m:
        nome = m.group(1).strip()
        if len(nome) > 4:
            return nome

    # Detalhamento do Faturamento: EMPREGADOR: nome (CNPJ ...)
    m = re.search(r"EMPREGADOR[:\s]+(.+?)\s*\(CNPJ", texto, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # DANFSe: bloco TOMADOR DO SERVIÇO → Nome / Nome Empresarial
    if "TOMADOR DO SERVI" in texto and "Nome" in texto:
        inicio = texto.find("TOMADOR DO SERVI")
        fim = texto.find("INTERMEDIÁRIO", inicio)
        if fim == -1:
            fim = inicio + 600
        bloco = texto[inicio:fim]
        m2 = re.search(r"Nome\s*/\s*Nome Empresarial\s*\n(.+)", bloco)
        if m2:
            return m2.group(1).strip()

    return ""


# ============================================================
#  MATCH POR NOME DE ARQUIVO (tentativa antes de abrir o PDF)
# ============================================================

LIMIAR_SCORE_NOME = 0.72         # score mínimo (0-1) para considerar um match confiável
LIMIAR_DIFERENCA_AMBIGUA = 0.08  # se o 2º colocado ficar muito perto do 1º, é ambíguo

#  Palavras que indicam o TIPO do documento no nome do arquivo (não fazem parte
#  do nome do condomínio) e que diluem o fuzzy match — ex: "ARGENTINA QUITADO
#  05.26.pdf" comparado com "ARGENTINA" caía para score 0.69 por causa do
#  "QUITADO". São removidas do nome do arquivo antes de comparar com o cadastro.
PALAVRAS_TIPO_DOC = {
    "QUITADO", "NF", "RECIBO", "DEMONSTRATIVO", "NOTA", "FISCAL", "BOLETO",
}


def normalizar_texto_busca(texto):
    """Maiúsculas, sem acento, sem dígitos (datas/códigos), separadores viram espaço."""
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[_\-.]", " ", texto)
    texto = re.sub(r"\d", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    return texto


def remover_palavras_tipo_doc(texto_normalizado):
    """Remove tokens de tipo de documento (QUITADO, NF, ...) de um texto já
    normalizado. Se sobrar vazio (nome só com essas palavras), devolve o
    texto original para não perder o match por completo."""
    tokens = [t for t in texto_normalizado.split() if t not in PALAVRAS_TIPO_DOC]
    return " ".join(tokens) if tokens else texto_normalizado


def _codigos_do_cadastro(cadastro):
    """dict código (string) -> cnpj_norm, a partir do cadastro."""
    codigos = {}
    for cnpj_norm, dados in cadastro.items():
        codigo = dados.get("codigo")
        if codigo:
            codigos.setdefault(str(codigo).strip(), cnpj_norm)
    return codigos


def buscar_por_nome_arquivo(nome_arquivo, cadastro):
    """
    Tenta identificar o condomínio a partir do nome do arquivo, em duas
    etapas (a primeira que resolver, ganha — não combina as duas):

      1. Código exato: alguns lotes (ex: F&F) já nomeiam o arquivo com o
         código do condomínio embutido, ex: "PGR 10004 Klosters.pdf" -> 10004.
         Extrai os números do nome e confere se algum bate, letra por letra,
         com um código do cadastro. Mais confiável que o fuzzy — sem essa
         ambiguidade de nomes parecidos.
      2. Fuzzy pelo nome: se nenhum número do nome bater com um código
         cadastrado, cai para o comportamento original — compara o nome do
         arquivo (limpo de palavras de tipo de documento) contra os nomes
         cadastrados por similaridade de texto.

    Retorna (cnpj_norm, descricao) se achar um match único e confiável
    — `descricao` já pronta pra exibição, ex: "código no nome do arquivo" ou
    "nome do arquivo, score 0.85" — ou (None, motivo) se não achar nada ou
    ficar ambíguo; nesse caso o chamador deve cair para a extração de CNPJ do
    conteúdo do PDF.
    """
    base = os.path.splitext(nome_arquivo)[0]

    # --- 1) Código exato no nome do arquivo ---
    numeros = re.findall(r"\d+", base)
    if numeros and cadastro:
        codigos_cadastro = _codigos_do_cadastro(cadastro)
        candidatos_codigo = []
        for numero in numeros:
            cnpj_norm = codigos_cadastro.get(numero)
            if cnpj_norm and cnpj_norm not in candidatos_codigo:
                candidatos_codigo.append(cnpj_norm)
        if len(candidatos_codigo) == 1:
            return candidatos_codigo[0], "código no nome do arquivo"
        if len(candidatos_codigo) > 1:
            return None, "mais de um código possível no nome do arquivo"

    # --- 2) Fuzzy pelo nome do condomínio (comportamento original) ---
    alvo = normalizar_texto_busca(base)
    alvo = remover_palavras_tipo_doc(alvo)
    if not alvo or not cadastro:
        return None, "nome de arquivo vazio ou cadastro vazio"

    pontuacoes = []
    for cnpj_norm, dados in cadastro.items():
        nome_cad = normalizar_texto_busca(dados.get("nome", ""))
        if not nome_cad:
            continue
        score = SequenceMatcher(None, alvo, nome_cad).ratio()
        pontuacoes.append((score, cnpj_norm))

    if not pontuacoes:
        return None, "cadastro sem nomes para comparar"

    pontuacoes.sort(key=lambda x: x[0], reverse=True)
    melhor_score, melhor_cnpj = pontuacoes[0]

    if melhor_score < LIMIAR_SCORE_NOME:
        return None, f"sem match confiável pelo nome (melhor score {melhor_score:.2f})"

    if len(pontuacoes) > 1 and (melhor_score - pontuacoes[1][0]) < LIMIAR_DIFERENCA_AMBIGUA:
        return None, f"nome do arquivo ambíguo entre condomínios parecidos (scores próximos)"

    return melhor_cnpj, f"nome do arquivo, score {melhor_score:.2f}"


def desempatar_por_cadastro(candidatos, cadastro):
    """
    Entre vários CNPJs candidatos, prefere o único que já está cadastrado.
    Usado por lotes sem emitente fixo (ex: Notas Diversas), onde os candidatos
    costumam ser o emitente da nota (não cadastrado) + o condomínio tomador
    (cadastrado). Devolve [o_cadastrado] se houver exatamente um cadastrado;
    caso contrário devolve a lista inalterada.
    """
    if len(candidatos) <= 1:
        return candidatos
    cadastrados = [c for c in candidatos if c in cadastro]
    if len(cadastrados) == 1:
        return cadastrados
    return candidatos


# ============================================================
#  OVERLAY / ESCRITA NO PDF (igual às versões anteriores)
# ============================================================

def criar_overlay(largura, altura, texto, fonte, tamanho, cor, x, y, centralizado):
    """Desenha `texto` no PDF. Se tiver quebras de linha ("\n"), cada linha é
    desenhada empilhada, a primeira em cima e as seguintes abaixo dela."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(largura, altura))
    c.setFont(fonte, tamanho)
    c.setFillColor(HexColor(cor))

    altura_linha = tamanho * 1.2
    for i, linha in enumerate(texto.split("\n")):
        x_linha = x
        if centralizado:
            largura_texto = c.stringWidth(linha, fonte, tamanho)
            x_linha = (largura - largura_texto) / 2
        c.drawString(x_linha, y - i * altura_linha, linha)

    c.save()
    buffer.seek(0)
    return buffer


def processar_pdf(caminho_entrada, caminho_saida, texto, config):
    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()
    for pagina in reader.pages:
        largura = float(pagina.mediabox.width)
        altura = float(pagina.mediabox.height)
        overlay_buffer = criar_overlay(
            largura, altura, texto,
            config["fonte"], config["tamanho"], config["cor"],
            config["x"], config["y"], config["centralizado"],
        )
        overlay_page = PdfReader(overlay_buffer).pages[0]
        pagina.merge_page(overlay_page)
        writer.add_page(pagina)
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "wb") as f:
        writer.write(f)


# ============================================================
#  PERSISTÊNCIA DO CADASTRO (planilha xlsx)
# ============================================================

def carregar_cadastro(caminho):
    """Retorna dict: cnpj_normalizado -> {'codigo':..., 'nome':...}"""
    cadastro = {}
    if not os.path.isfile(caminho):
        return cadastro
    wb = load_workbook(caminho)
    sheet = wb.active
    for linha in sheet.iter_rows(min_row=2, values_only=True):
        if not linha or not linha[0]:
            continue
        cnpj, codigo, nome = (list(linha) + [None, None, None])[:3]
        cnpj_norm = normalizar_cnpj(str(cnpj))
        if cnpj_norm:
            cadastro[cnpj_norm] = {
                "codigo": str(codigo).strip() if codigo is not None else "",
                "nome": str(nome).strip() if nome is not None else "",
            }
    return cadastro


def salvar_cadastro(caminho, cadastro):
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Condominios"
    sheet.append(["CNPJ", "Código", "Nome do Condomínio"])
    for cnpj_norm, dados in sorted(cadastro.items(), key=lambda kv: kv[1]["nome"]):
        sheet.append([formatar_cnpj(cnpj_norm), dados["codigo"], dados["nome"]])
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 12
    sheet.column_dimensions["C"].width = 35
    wb.save(caminho)
