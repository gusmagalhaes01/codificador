"""
Aplicação com:
  Aba 1 - Processamento dos PDFs (lê o CNPJ do tomador no PDF, busca no
          cadastro e escreve o código correspondente no arquivo)
  Aba 2 - Cadastro de Condomínios (planilha interativa: CNPJ, Código, Nome)
  Aba 3 - Logs (histórico completo salvo em arquivo .log)

Requisitos: pip install pypdf reportlab openpyxl pymupdf winocr customtkinter darkdetect
(tkinter já vem incluído no Python padrão do Windows)
OCR usa o motor nativo do Windows — sem instalação de programas externos.
Interface em CustomTkinter, identidade visual "Swiss International Style".
"""

import asyncio
import copy
import io
import json
import os
import re
import sys
import threading
import time
import datetime
import unicodedata
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

import customtkinter as ctk
import darkdetect

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
#  TEMA VISUAL — Swiss International Style (paleta "stone")
# ============================================================

TEMA_CLARO = {
    "fundo": "#FAFAF9", "superficie": "#F5F5F4", "borda": "#E7E5E4",
    "borda_forte": "#1C1917", "texto": "#1C1917", "texto_secundario": "#57534E",
    "texto_terciario": "#A8A29E", "acento": "#003B8E", "sobre_acento": "#FFFFFF",
    "acento_hover": "#002A66",
}
TEMA_ESCURO = {
    "fundo": "#0C0A09", "superficie": "#1C1917", "borda": "#292524",
    "borda_forte": "#FAFAF9", "texto": "#FAFAF9", "texto_secundario": "#A8A29E",
    "texto_terciario": "#57534E", "acento": "#2563EB", "sobre_acento": "#FFFFFF",
    "acento_hover": "#1D4FD0",
}


def familia_fonte():
    """
    Devolve "IBM Plex Sans" se a fonte estiver instalada no sistema, senão
    "Segoe UI". Evita repetir essa checagem em vários pontos do código.
    Precisa de uma raiz Tk existente para consultar as fontes instaladas;
    se não houver nenhuma (ou a consulta falhar por qualquer motivo),
    devolve "Segoe UI" com segurança.
    """
    try:
        fontes_instaladas = set(tkfont.families())
        if "IBM Plex Sans" in fontes_instaladas:
            return "IBM Plex Sans"
    except Exception:
        pass
    return "Segoe UI"


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


# ============================================================
#  APLICAÇÃO GRÁFICA
# ============================================================

class App(ctk.CTk):
    def __init__(self):
        # --- Config persistida (config.json, ao lado do script) ---
        self.config_app = carregar_config()

        # --- Tema efetivo: resolve "auto" via darkdetect, aplica no CTk ---
        self.nome_tema = self._resolver_nome_tema(self.config_app.get("tema", "auto"))
        self.tema_atual = TEMA_CLARO if self.nome_tema == "claro" else TEMA_ESCURO
        ctk.set_appearance_mode("Dark" if self.nome_tema == "escuro" else "Light")

        super().__init__()
        self.title("Codificador v6.5.0")
        self.geometry("780x680")
        self.minsize(620, 420)
        self.resizable(True, True)
        self._aplicar_icone()

        # Caminho da planilha de cadastro e do log (ficam ao lado deste script)
        pasta_script = pasta_base()
        self.caminho_planilha = os.path.join(pasta_script, NOME_ARQUIVO_PADRAO)
        self.caminho_log = os.path.join(pasta_script, NOME_LOG_PADRAO)
        self.cadastro = carregar_cadastro(self.caminho_planilha)

        # Predefinição ativa (FedCorp / F&F / Notas Diversas) — cada uma é um
        # "pacote" de configurações de lote, trocado de uma vez pelos botões
        # da tela principal. self.predefinicao_ativa guarda só a CHAVE
        # ("fedcorp"/"ff"/"notas_diversas"); os valores em si vivem nas
        # tk.Var abaixo, inicializadas a partir do perfil ativo.
        self.predefinicao_ativa = tk.StringVar(value=self.config_app["predefinicao_ativa"])
        perfil_inicial = self.config_app["predefinicoes"][self.config_app["predefinicao_ativa"]]

        # Variáveis - aba processamento (inicializadas a partir do perfil
        # ativo da config persistida, não mais de literais hardcoded — assim
        # sobrevivem entre sessões)
        self.pasta_entrada = tk.StringVar()
        self.pasta_saida = tk.StringVar()
        self.cnpj_emitente = tk.StringVar(value=perfil_inicial["cnpj_emitente"])
        self.modo_texto = tk.StringVar(value=perfil_inicial["modo_texto"])
        self.tamanho_fonte = tk.StringVar(value=perfil_inicial["tamanho_fonte"])
        self.cor_texto = tk.StringVar(value=perfil_inicial["cor_texto"])
        self.usar_ocr = tk.BooleanVar(value=perfil_inicial["usar_ocr"])
        self.dpi_ocr = tk.IntVar(value=perfil_inicial["dpi_ocr"])

        # OCR por região (recorte) — opcional, calibrado pelo usuário
        self.usar_ocr_regiao = tk.BooleanVar(value=perfil_inicial["usar_ocr_regiao"])
        self.regiao_x0 = tk.DoubleVar(value=perfil_inicial["regiao_x0"])
        self.regiao_y0 = tk.DoubleVar(value=perfil_inicial["regiao_y0"])
        self.regiao_x1 = tk.DoubleVar(value=perfil_inicial["regiao_x1"])
        self.regiao_y1 = tk.DoubleVar(value=perfil_inicial["regiao_y1"])

        # Match por nome de arquivo (código exato ou fuzzy — ver
        # buscar_por_nome_arquivo; evita abrir o PDF na maioria dos casos)
        self.usar_match_nome_arquivo = tk.BooleanVar(value=perfil_inicial["usar_match_nome_arquivo"])

        # Renomear o PDF de saída com o código do condomínio na frente
        self.renomear_com_codigo = tk.BooleanVar(value=perfil_inicial["renomear_com_codigo"])

        # Variáveis - aba cadastro (formulário)
        self.form_cnpj = tk.StringVar()
        self.form_codigo = tk.StringVar()
        self.form_nome = tk.StringVar()

        self._montar_interface()
        self._atualizar_tabela_cadastro()
        self._carregar_log_em_tela()

    def _aplicar_icone(self):
        """
        Aplica o ícone (icone.ico) na barra de título/taskbar da janela.
        O CustomTkinter põe um ícone padrão próprio logo após a criação da
        janela (via um `after` interno), então além de setar na hora, re-aplica
        depois de um instante para sobrepor o padrão do CTk. Falha silenciosa —
        um ícone ausente nunca deve impedir o app de abrir.
        """
        caminho = caminho_recurso(NOME_ICONE)
        if not os.path.isfile(caminho):
            return

        def setar():
            try:
                self.iconbitmap(caminho)
            except Exception:
                pass

        setar()
        self.after(300, setar)

    # --------------------------------------------------------
    #  TEMA
    # --------------------------------------------------------
    @staticmethod
    def _resolver_nome_tema(tema_config):
        """Resolve o valor guardado em config["tema"] ("auto"/"claro"/"escuro")
        para "claro" ou "escuro" efetivos, consultando o SO quando for "auto"."""
        if tema_config == "claro":
            return "claro"
        if tema_config == "escuro":
            return "escuro"
        # "auto" (ou qualquer valor desconhecido) — detecta pelo SO
        try:
            return "escuro" if darkdetect.isDark() else "claro"
        except Exception:
            return "claro"

    def aplicar_tema(self, nome):
        """
        Troca o tema ativo ("claro" ou "escuro"), persiste em
        config["tema"] e salva config.json. Reaplica as cores no fundo da
        janela, no appearance mode do CustomTkinter e, via `_recolorir`,
        em todos os widgets já existentes registrados em self._widgets_tema.
        """
        nome = "escuro" if nome == "escuro" else "claro"
        self.nome_tema = nome
        self.tema_atual = TEMA_ESCURO if nome == "escuro" else TEMA_CLARO

        self.config_app["tema"] = nome
        salvar_config(self.config_app)

        ctk.set_appearance_mode("Dark" if nome == "escuro" else "Light")
        try:
            self.configure(fg_color=self.tema_atual["fundo"])
        except Exception:
            pass
        self._recolorir()
        # _recolorir não cobre a cor de borda/texto dos botões de
        # predefinição (ativo x inativo é estado, não só tema) — reaplica
        # com os valores do tema novo.
        self._atualizar_botoes_predefinicao()

    def alternar_tema(self):
        """Cicla claro <-> escuro, aplica e persiste."""
        novo = "escuro" if self.nome_tema == "claro" else "claro"
        self.aplicar_tema(novo)

    def report_callback_exception(self, exc, val, tb):
        """
        Chamado automaticamente pelo Tkinter quando qualquer callback (clique
        de botão, etc.) lança uma exceção não tratada. Sem isso, em algumas
        instalações (ex: rodando com pythonw, sem console) o erro tenta ir
        pro stderr, falha, e a janela inteira fecha sem aviso nenhum.
        Aqui a exceção é gravada no log e mostrada num popup, sem derrubar o app.
        """
        import traceback
        detalhes = "".join(traceback.format_exception(exc, val, tb))
        _gravar_erros_log(detalhes)
        messagebox.showerror(
            "Erro inesperado",
            f"{val}\n\nDetalhes salvos em erros.log (na pasta do programa).",
        )

    # --------------------------------------------------------
    #  INTERFACE GERAL
    # --------------------------------------------------------
    def _montar_interface(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.notebook = notebook

        self.aba_processar = ttk.Frame(notebook)
        self.aba_cadastro = ttk.Frame(notebook)
        self.aba_logs = ttk.Frame(notebook)
        notebook.add(self.aba_processar, text="1. Processamento")
        notebook.add(self.aba_cadastro, text="2. Cadastro de Condomínios")
        notebook.add(self.aba_logs, text="3. Logs")

        self._montar_aba_processar(self.aba_processar)
        self._montar_aba_cadastro(self.aba_cadastro)
        self._montar_aba_logs(self.aba_logs)

        # _montar_aba_processar já chama _estilizar_ttk() ao final, mas nesse
        # ponto as abas de Cadastro/Logs ainda não existem — chama de novo
        # agora que todas as três existem, para log_historico etc. também
        # receberem a cor do tema já na montagem inicial (não só ao alternar).
        self._estilizar_ttk()

    # --------------------------------------------------------
    #  ABA 2 — CADASTRO
    # --------------------------------------------------------
    def _montar_aba_cadastro(self, parent):
        pad = {"padx": 24, "pady": 6}
        fonte = familia_fonte()

        def caption(texto):
            return ttk.Label(parent, text=texto.upper(), style="Caption.TLabel", font=(fonte, 11))

        def caption_in(parent_widget, texto):
            return ttk.Label(parent_widget, text=texto.upper(), style="Caption.TLabel", font=(fonte, 11))

        titulo = ttk.Label(parent, text="Condomínios", style="Titulo.TLabel", font=(fonte, 20))
        titulo.pack(anchor="w", padx=24, pady=(20, 0))

        self.label_planilha_cadastro = ttk.Label(
            parent, text=self.caminho_planilha, style="Caption.TLabel", font=("Consolas", 10),
        )
        self.label_planilha_cadastro.pack(anchor="w", padx=24, pady=(4, 16))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=24)

        caption(" Adicionar / editar").pack(anchor="w", padx=24, pady=(16, 8))

        frame_form = ttk.Frame(parent)
        frame_form.pack(fill="x", padx=24)
        frame_form.columnconfigure(0, weight=0)
        frame_form.columnconfigure(1, weight=0)
        frame_form.columnconfigure(2, weight=1)

        caption_in(frame_form, "CNPJ").grid(row=0, column=0, sticky="w", padx=(0, 16))
        caption_in(frame_form, "Código").grid(row=0, column=1, sticky="w", padx=(0, 16))
        caption_in(frame_form, "Nome").grid(row=0, column=2, sticky="w")

        ttk.Entry(frame_form, textvariable=self.form_cnpj, width=22, font=(fonte, 12)).grid(
            row=1, column=0, sticky="w", padx=(0, 16), pady=(2, 12))
        ttk.Entry(frame_form, textvariable=self.form_codigo, width=12, font=(fonte, 12)).grid(
            row=1, column=1, sticky="w", padx=(0, 16), pady=(2, 12))
        ttk.Entry(frame_form, textvariable=self.form_nome, font=(fonte, 12)).grid(
            row=1, column=2, sticky="ew", pady=(2, 12))

        linha_botoes_form = ttk.Frame(parent)
        linha_botoes_form.pack(fill="x", padx=24, pady=(0, 16))
        ttk.Button(
            linha_botoes_form, text="Salvar", style="Cobalto.TButton",
            command=self._adicionar_ou_atualizar,
        ).pack(side="left")
        ttk.Button(linha_botoes_form, text="Limpar campos", command=self._limpar_form).pack(side="left", padx=8)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=24, pady=(0, 16))

        # --- Tabela ---
        self.label_contagem_cadastro = caption("0 cadastrados")
        self.label_contagem_cadastro.pack(anchor="w", padx=24, pady=(0, 8))

        frame_tabela = ttk.Frame(parent)
        frame_tabela.pack(fill="both", expand=True, padx=24)

        colunas = ("cnpj", "codigo", "nome")
        self.tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=14)
        self.tabela.heading("cnpj", text="CNPJ")
        self.tabela.heading("codigo", text="Código")
        self.tabela.heading("nome", text="Nome do Condomínio")
        self.tabela.column("cnpj", width=150, anchor="w")
        self.tabela.column("codigo", width=80, anchor="w")
        self.tabela.column("nome", width=380, anchor="w")
        self.tabela.pack(fill="both", expand=True, side="left")
        self.tabela.bind("<<TreeviewSelect>>", self._selecionar_linha)

        scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="left", fill="y")

        # --- Botões de baixo ---
        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=24, pady=(16, 0))
        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill="x", padx=24, pady=16)
        ttk.Button(frame_botoes, text="Remover selecionado", command=self._remover_selecionado).pack(side="left")
        ttk.Button(frame_botoes, text="Importar planilha...", command=self._importar_planilha).pack(side="left", padx=8)
        ttk.Button(frame_botoes, text="Salvar planilha agora", command=self._salvar_planilha).pack(side="left")

    def _atualizar_tabela_cadastro(self):
        self.tabela.delete(*self.tabela.get_children())
        for cnpj_norm, dados in sorted(self.cadastro.items(), key=lambda kv: kv[1]["codigo"]):
            self.tabela.insert("", "end", iid=cnpj_norm,
                                values=(formatar_cnpj(cnpj_norm), dados["codigo"], dados["nome"]))
        if hasattr(self, "label_contagem_cadastro"):
            self.label_contagem_cadastro.configure(text=f"{len(self.cadastro)} CADASTRADOS")

    def _selecionar_linha(self, event):
        selecionado = self.tabela.selection()
        if not selecionado:
            return
        cnpj_norm = selecionado[0]
        dados = self.cadastro.get(cnpj_norm, {})
        self.form_cnpj.set(formatar_cnpj(cnpj_norm))
        self.form_codigo.set(dados.get("codigo", ""))
        self.form_nome.set(dados.get("nome", ""))

    def _limpar_form(self):
        self.form_cnpj.set("")
        self.form_codigo.set("")
        self.form_nome.set("")
        self.tabela.selection_remove(self.tabela.selection())

    def _adicionar_ou_atualizar(self):
        cnpj_raw = self.form_cnpj.get().strip()
        codigo = self.form_codigo.get().strip()
        nome = self.form_nome.get().strip()

        cnpj_norm = normalizar_cnpj(cnpj_raw)
        if len(cnpj_norm) != 14:
            messagebox.showerror("Erro", "CNPJ inválido — deve ter 14 dígitos.")
            return
        if not codigo:
            messagebox.showerror("Erro", "Informe o código.")
            return

        self.cadastro[cnpj_norm] = {"codigo": codigo, "nome": nome}
        self._atualizar_tabela_cadastro()
        self._salvar_planilha(silencioso=True)
        self._limpar_form()

    def _remover_selecionado(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showinfo("Aviso", "Selecione uma linha na tabela primeiro.")
            return
        cnpj_norm = selecionado[0]
        if messagebox.askyesno("Confirmar", "Remover este condomínio do cadastro?"):
            self.cadastro.pop(cnpj_norm, None)
            self._atualizar_tabela_cadastro()
            self._salvar_planilha(silencioso=True)
            self._limpar_form()

    def _importar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha (.xlsx)",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not caminho:
            return
        novo_cadastro = carregar_cadastro(caminho)
        self.cadastro.update(novo_cadastro)
        self._atualizar_tabela_cadastro()
        self._salvar_planilha(silencioso=True)
        messagebox.showinfo("Importado", f"{len(novo_cadastro)} registro(s) importado(s).")

    def _salvar_planilha(self, silencioso=False):
        try:
            salvar_cadastro(self.caminho_planilha, self.cadastro)
            if not silencioso:
                messagebox.showinfo("Salvo", f"Planilha salva em:\n{self.caminho_planilha}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))

    # --------------------------------------------------------
    #  ABA 1 — PROCESSAMENTO DE PDFs
    # --------------------------------------------------------
    def _montar_aba_processar(self, parent_externo):
        """
        Tela Principal — Swiss International Style, sem scroll. Cabeçalho
        (título + Configurações + alternador de tema), campos de pasta de
        entrada/saída e botão primário de processamento. Os controles de
        OCR/DPI/região/estilo saíram desta tela: viram opções avançadas na
        Tarefa 3 (modal de Configurações).
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        # Lista de (widget, {propriedade: chave_do_tema}) usada por
        # self._recolorir() para reaplicar as cores do tema quando o usuário
        # alterna claro/escuro, sem precisar remontar a tela inteira.
        self._widgets_tema = []

        def registrar(widget, mapa):
            self._widgets_tema.append((widget, mapa))
            for prop, chave in mapa.items():
                try:
                    widget.configure(**{prop: tema[chave]})
                except Exception:
                    pass
            return widget

        container = registrar(
            ctk.CTkFrame(parent_externo, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        container.pack(fill="both", expand=True)

        # --- CABEÇALHO ---
        cabecalho = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        cabecalho.pack(fill="x", padx=24, pady=(24, 0))
        cabecalho.columnconfigure(0, weight=1)
        cabecalho.columnconfigure(1, weight=0)

        bloco_titulo = registrar(
            ctk.CTkFrame(cabecalho, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_titulo.grid(row=0, column=0, sticky="w")

        titulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="Codificador", font=(fonte, 20),
                         text_color=tema["texto"], anchor="w"),
            {"text_color": "texto"},
        )
        titulo.pack(anchor="w")

        subtitulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="IDENTIFICAÇÃO DE PDFS POR CNPJ", font=(fonte, 11),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        subtitulo.pack(anchor="w")

        bloco_acoes = registrar(
            ctk.CTkFrame(cabecalho, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_acoes.grid(row=0, column=1, sticky="e")

        self.botao_config = registrar(
            ctk.CTkButton(
                bloco_acoes, text="⚙ Configurações", corner_radius=0,
                fg_color="transparent", hover_color=tema["superficie"],
                border_width=1, border_color=tema["borda_forte"],
                text_color=tema["texto"], font=(fonte, 13),
                command=self._abrir_configuracoes,
            ),
            {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
        )
        self.botao_config.pack(side="left", padx=(0, 8))

        self.botao_tema = registrar(
            ctk.CTkButton(
                bloco_acoes, text=self._texto_botao_tema(), corner_radius=0,
                fg_color="transparent", hover_color=tema["superficie"],
                border_width=1, border_color=tema["borda_forte"],
                text_color=tema["texto"], font=(fonte, 13),
                command=self.alternar_tema,
            ),
            {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
        )
        self.botao_tema.pack(side="left")

        # --- hairline separando o cabeçalho do resto ---
        hairline = registrar(
            ctk.CTkFrame(container, height=1, corner_radius=0, fg_color=tema["borda"]),
            {"fg_color": "borda"},
        )
        hairline.pack(fill="x", padx=24, pady=(16, 24))

        # --- PREDEFINIÇÃO (perfil de lote ativo: FedCorp / F&F / Notas
        #     Diversas) — troca CNPJ emitente, identificação e carimbo de
        #     uma vez, sem reconfigurar tudo a cada lote. ---
        bloco_predef = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_predef.pack(fill="x", padx=24, pady=(0, 24))

        label_predef = registrar(
            ctk.CTkLabel(bloco_predef, text="PREDEFINIÇÃO", font=(fonte, 11),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        label_predef.pack(anchor="w", pady=(0, 4))

        linha_predef = registrar(
            ctk.CTkFrame(bloco_predef, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        linha_predef.pack(fill="x")

        # Cada botão é registrado no tema só pro hover (a cor de borda/texto
        # que indica qual está ativo é responsabilidade de
        # _atualizar_botoes_predefinicao, chamada logo abaixo e de novo
        # sempre que o tema alterna — ver aplicar_tema).
        self._botoes_predefinicao = {}
        for i, chave in enumerate(ORDEM_PREDEFINICOES):
            nome_perfil = self.config_app["predefinicoes"][chave]["nome"]
            botao = registrar(
                ctk.CTkButton(
                    linha_predef, text=nome_perfil, corner_radius=0,
                    fg_color="transparent", hover_color=tema["superficie"],
                    border_width=1, font=(fonte, 13),
                    command=lambda c=chave: self._trocar_predefinicao(c),
                ),
                {"hover_color": "superficie"},
            )
            botao.pack(side="left", padx=(0, 8) if i < len(ORDEM_PREDEFINICOES) - 1 else 0)
            self._botoes_predefinicao[chave] = botao

        self._atualizar_botoes_predefinicao()

        # --- CORPO ---
        corpo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        corpo.pack(fill="both", expand=True, padx=24)
        corpo.columnconfigure(0, weight=1)

        def montar_campo_pasta(linha_grid, rotulo, variavel, comando_trocar):
            label = registrar(
                ctk.CTkLabel(corpo, text=rotulo, font=(fonte, 11),
                             text_color=tema["texto_secundario"], anchor="w"),
                {"text_color": "texto_secundario"},
            )
            label.grid(row=linha_grid, column=0, sticky="w", pady=(0, 4))

            linha = registrar(
                ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"]),
                {"fg_color": "fundo"},
            )
            linha.grid(row=linha_grid + 1, column=0, sticky="ew", pady=(0, 16))
            linha.columnconfigure(0, weight=1)

            entry = registrar(
                ctk.CTkEntry(linha, textvariable=variavel, corner_radius=0,
                             fg_color=tema["superficie"], border_width=1, border_color=tema["borda"],
                             text_color=tema["texto"], font=(fonte, 13)),
                {"fg_color": "superficie", "border_color": "borda", "text_color": "texto"},
            )
            entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

            botao = registrar(
                ctk.CTkButton(
                    linha, text="Trocar", corner_radius=0, width=90,
                    fg_color="transparent", hover_color=tema["superficie"],
                    border_width=1, border_color=tema["borda_forte"], text_color=tema["texto"],
                    font=(fonte, 13), command=comando_trocar,
                ),
                {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
            )
            botao.grid(row=0, column=1)

        montar_campo_pasta(0, "PASTA COM OS PDFS", self.pasta_entrada, self._escolher_pasta_entrada)
        montar_campo_pasta(2, "SALVAR PDFS CODIFICADOS EM", self.pasta_saida, self._escolher_pasta_saida)

        # --- Botão primário ---
        self.botao_iniciar = registrar(
            ctk.CTkButton(
                corpo, text="Processar PDFs", corner_radius=0, height=44, font=(fonte, 15),
                fg_color=tema["acento"], hover_color=tema["acento_hover"],
                text_color=tema["sobre_acento"], border_width=0,
                command=self._iniciar_processamento,
            ),
            {"fg_color": "acento", "hover_color": "acento_hover", "text_color": "sobre_acento"},
        )
        self.botao_iniciar.grid(row=4, column=0, sticky="ew", pady=(8, 8))

        # --- Linha explicativa ---
        explicacao = registrar(
            ctk.CTkLabel(
                corpo,
                text=("Identificação: nome do arquivo, depois leitura do PDF. Boletos escaneados "
                      "são lidos automaticamente."),
                font=(fonte, 13), text_color=tema["texto_terciario"], justify="left", anchor="w",
            ),
            {"text_color": "texto_terciario"},
        )
        explicacao.grid(row=5, column=0, sticky="w", pady=(0, 16))

        # --- Barra de progresso (usada por _processar_em_thread) ---
        self.barra_progresso = ttk.Progressbar(corpo, mode="determinate")
        self.barra_progresso.grid(row=6, column=0, sticky="ew", pady=(0, 24))

        # --- Widgets ocultos usados pelo loop de processamento, mas sem
        # lugar na tela Swiss limpa. Nunca são pack/grid — ficam fora da
        # árvore visível, só existem como atributos para não quebrar
        # _iniciar_processamento/_processar_em_thread. ---
        self._frame_oculto = tk.Frame(parent_externo)

        # temporário: Tarefa 4/5 substitui pelo painel de Resultado
        self.log_text = tk.Text(self._frame_oculto, state="disabled", wrap="word")

        # a Tarefa 3 realoca este campo no modal de Configurações
        self.txt_tipo_servico = tk.Text(self._frame_oculto, width=30, height=3, wrap="none")
        perfil_atual = self.config_app["predefinicoes"][self.predefinicao_ativa.get()]
        self.txt_tipo_servico.insert("1.0", perfil_atual["tipo_servico"])

        self._atualizar_botao_processar()
        self._estilizar_ttk()

    def _texto_botao_tema(self):
        return "☀ Claro" if self.nome_tema == "escuro" else "🌙 Escuro"

    def _atualizar_botao_processar(self):
        """
        Conta os *.pdf (case-insensitive) na pasta de entrada e ajusta o
        texto/estado do botão primário. Chamado ao montar a tela e sempre
        que o usuário troca a pasta de entrada.
        """
        pasta = self.pasta_entrada.get().strip()
        quantidade = 0
        if pasta and os.path.isdir(pasta):
            try:
                quantidade = sum(1 for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
            except Exception:
                quantidade = 0

        if quantidade > 0:
            self.botao_iniciar.configure(text=f"Processar {quantidade} PDFs", state="normal")
        else:
            self.botao_iniciar.configure(text="Processar PDFs", state="disabled")

    def _trocar_predefinicao(self, chave):
        """
        Troca a predefinição ativa (FedCorp / F&F / Notas Diversas): aplica
        os valores do perfil nas tk.Var vivas, persiste qual ficou ativa em
        config.json, e atualiza os botões da tela principal.
        """
        if chave not in self.config_app["predefinicoes"]:
            return

        perfil = self.config_app["predefinicoes"][chave]
        self.predefinicao_ativa.set(chave)
        self.config_app["predefinicao_ativa"] = chave
        salvar_config(self.config_app)

        self.cnpj_emitente.set(perfil["cnpj_emitente"])
        self.usar_match_nome_arquivo.set(perfil["usar_match_nome_arquivo"])
        self.usar_ocr.set(perfil["usar_ocr"])
        self.dpi_ocr.set(perfil["dpi_ocr"])
        self.usar_ocr_regiao.set(perfil["usar_ocr_regiao"])
        self.regiao_x0.set(perfil["regiao_x0"])
        self.regiao_y0.set(perfil["regiao_y0"])
        self.regiao_x1.set(perfil["regiao_x1"])
        self.regiao_y1.set(perfil["regiao_y1"])
        self.modo_texto.set(perfil["modo_texto"])
        self.tamanho_fonte.set(perfil["tamanho_fonte"])
        self.cor_texto.set(perfil["cor_texto"])
        self.renomear_com_codigo.set(perfil["renomear_com_codigo"])
        self.txt_tipo_servico.delete("1.0", "end")
        self.txt_tipo_servico.insert("1.0", perfil["tipo_servico"])

        self._atualizar_botoes_predefinicao()

    def _perfil_ativo_modificado(self):
        """
        True se alguma das tk.Var "vivas" divergir do que está salvo na
        predefinição ativa — ex: usuário mexeu em algo nas Configurações mas
        ainda não clicou em Salvar. Usado só pra exibir o aviso "(modificado)"
        no botão da tela principal; não bloqueia nada.
        """
        perfil = self.config_app["predefinicoes"].get(self.predefinicao_ativa.get())
        if perfil is None:
            return False
        try:
            atual = {
                "cnpj_emitente": self.cnpj_emitente.get(),
                "usar_match_nome_arquivo": self.usar_match_nome_arquivo.get(),
                "usar_ocr": self.usar_ocr.get(),
                "dpi_ocr": self.dpi_ocr.get(),
                "usar_ocr_regiao": self.usar_ocr_regiao.get(),
                "regiao_x0": self.regiao_x0.get(),
                "regiao_y0": self.regiao_y0.get(),
                "regiao_x1": self.regiao_x1.get(),
                "regiao_y1": self.regiao_y1.get(),
                "modo_texto": self.modo_texto.get(),
                "tamanho_fonte": self.tamanho_fonte.get(),
                "cor_texto": self.cor_texto.get(),
                "renomear_com_codigo": self.renomear_com_codigo.get(),
                "tipo_servico": self.txt_tipo_servico.get("1.0", "end-1c"),
            }
        except Exception:
            return False
        return any(atual[chave] != perfil.get(chave) for chave in atual)

    def _atualizar_botoes_predefinicao(self):
        """
        Recolore os botões de predefinição da tela principal: o ativo com
        contorno forte e texto primário, os demais em texto secundário.
        Acrescenta "(modificado)" no ativo se as vars vivas divergirem do
        que está salvo (ver _perfil_ativo_modificado).
        """
        if not hasattr(self, "_botoes_predefinicao"):
            return
        tema = self.tema_atual
        ativa = self.predefinicao_ativa.get()
        modificado = self._perfil_ativo_modificado()

        for chave, botao in self._botoes_predefinicao.items():
            nome_perfil = self.config_app["predefinicoes"][chave]["nome"]
            try:
                if chave == ativa:
                    texto = f"{nome_perfil} (modificado)" if modificado else nome_perfil
                    botao.configure(
                        text=texto, border_color=tema["borda_forte"], text_color=tema["texto"],
                    )
                else:
                    botao.configure(
                        text=nome_perfil, border_color=tema["borda"], text_color=tema["texto_secundario"],
                    )
            except Exception:
                pass

    # --------------------------------------------------------
    #  MODAL DE CONFIGURAÇÕES (Tarefa 3)
    # --------------------------------------------------------

    _OPCOES_QUALIDADE = [("Rápida", 72), ("Normal", 200), ("Máxima", 300)]

    def _abrir_configuracoes(self):
        """
        Abre o modal de Configurações (ctk.CTkToplevel), que agrupa todo o
        técnico que saiu da Tela Principal: empresa, identificação, leitura
        de boletos escaneados (e leitura por região) e texto escrito no PDF.

        Padrão snapshot/cancelar/salvar: os widgets do modal ligam-se
        diretamente às self.* vars reais, então um snapshot é tirado na
        abertura para permitir descartar as edições no Cancelar.
        """
        if getattr(self, "_janela_config", None) is not None and self._janela_config.winfo_exists():
            self._janela_config.focus_force()
            return

        tema = self.tema_atual
        fonte = familia_fonte()

        # --- snapshot dos valores atuais, para o Cancelar restaurar ---
        snapshot = {
            "cnpj_emitente": self.cnpj_emitente.get(),
            "usar_match_nome_arquivo": self.usar_match_nome_arquivo.get(),
            "usar_ocr": self.usar_ocr.get(),
            "dpi_ocr": self.dpi_ocr.get(),
            "usar_ocr_regiao": self.usar_ocr_regiao.get(),
            "regiao_x0": self.regiao_x0.get(),
            "regiao_y0": self.regiao_y0.get(),
            "regiao_x1": self.regiao_x1.get(),
            "regiao_y1": self.regiao_y1.get(),
            "modo_texto": self.modo_texto.get(),
            "tamanho_fonte": self.tamanho_fonte.get(),
            "cor_texto": self.cor_texto.get(),
            "renomear_com_codigo": self.renomear_com_codigo.get(),
        }

        nome_perfil = self.config_app["predefinicoes"][self.predefinicao_ativa.get()]["nome"]

        janela = ctk.CTkToplevel(self)
        self._janela_config = janela
        janela.title(f"Configurações — Editando: {nome_perfil}")
        janela.geometry("640x680")
        janela.minsize(560, 520)
        janela.resizable(True, True)
        janela.configure(fg_color=tema["fundo"])
        janela.transient(self)
        janela.grab_set()

        def rotulo_secao(parent, texto):
            lbl = ctk.CTkLabel(
                parent, text=texto.upper(), font=(fonte, 11, "bold"),
                text_color=tema["texto_secundario"], anchor="w",
            )
            lbl.pack(fill="x", padx=24, pady=(24, 8))
            return lbl

        def hairline(parent):
            linha = ctk.CTkFrame(parent, height=1, corner_radius=0, fg_color=tema["borda"])
            linha.pack(fill="x", padx=24, pady=(16, 0))
            return linha

        # --- rodapé (Cancelar/Salvar), empacotado ANTES do conteúdo rolável
        # para nunca ser espremido pelo CTkScrollableFrame expansível ---
        rodape = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        rodape.pack(side="bottom", fill="x", padx=24, pady=16)

        # --- área rolável para o conteúdo do modal ---
        corpo = ctk.CTkScrollableFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        corpo.pack(fill="both", expand=True)

        # ===================== EMPRESA =====================
        rotulo_secao(corpo, "Empresa")
        frame_empresa = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        frame_empresa.pack(fill="x", padx=24)
        ctk.CTkLabel(
            frame_empresa, text="CNPJ da sua empresa (ignorado na identificação)",
            font=(fonte, 12), text_color=tema["texto"], anchor="w",
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkEntry(
            frame_empresa, textvariable=self.cnpj_emitente, corner_radius=0,
            fg_color=tema["superficie"], border_width=1, border_color=tema["borda"],
            text_color=tema["texto"], font=(fonte, 13),
        ).pack(fill="x", pady=(0, 8))

        hairline(corpo)

        # ===================== IDENTIFICAÇÃO =====================
        rotulo_secao(corpo, "Identificação")
        frame_ident = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        frame_ident.pack(fill="x", padx=24)
        ctk.CTkCheckBox(
            frame_ident, text="Tentar identificar pelo nome do arquivo",
            variable=self.usar_match_nome_arquivo, corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento"],
            border_color=tema["borda_forte"], text_color=tema["texto"], font=(fonte, 13),
        ).pack(anchor="w", pady=(0, 8))

        hairline(corpo)

        # ===================== LEITURA DE BOLETOS ESCANEADOS =====================
        rotulo_secao(corpo, "Leitura de boletos escaneados")
        frame_leitura = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        frame_leitura.pack(fill="x", padx=24)

        chk_ocr = ctk.CTkCheckBox(
            frame_leitura, text="Ler boletos escaneados automaticamente (mais lento)",
            variable=self.usar_ocr, corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento"],
            border_color=tema["borda_forte"], text_color=tema["texto"], font=(fonte, 13),
        )
        chk_ocr.pack(anchor="w", pady=(0, 8))
        if not OCR_DISPONIVEL:
            chk_ocr.configure(state="disabled")
            ctk.CTkLabel(
                frame_leitura,
                text="Leitura de boletos escaneados não está disponível nesta instalação.",
                font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w", justify="left",
            ).pack(fill="x", pady=(0, 8))

        linha_qualidade = ctk.CTkFrame(frame_leitura, corner_radius=0, fg_color=tema["fundo"])
        linha_qualidade.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            linha_qualidade, text="Qualidade de leitura", font=(fonte, 12),
            text_color=tema["texto"], anchor="w",
        ).pack(side="left", padx=(0, 12))

        rotulos_qualidade = [nome for nome, _ in self._OPCOES_QUALIDADE]
        mapa_qualidade = dict(self._OPCOES_QUALIDADE)
        mapa_qualidade_inverso = {v: k for k, v in self._OPCOES_QUALIDADE}
        rotulo_inicial = mapa_qualidade_inverso.get(self.dpi_ocr.get(), "Normal")
        var_qualidade = tk.StringVar(value=rotulo_inicial)

        def ao_trocar_qualidade(escolha):
            self.dpi_ocr.set(mapa_qualidade.get(escolha, 200))

        combo_qualidade = ctk.CTkOptionMenu(
            linha_qualidade, values=rotulos_qualidade, variable=var_qualidade,
            corner_radius=0, fg_color=tema["superficie"], button_color=tema["borda_forte"],
            button_hover_color=tema["acento"], text_color=tema["texto"],
            dropdown_fg_color=tema["superficie"], dropdown_text_color=tema["texto"],
            font=(fonte, 13), command=ao_trocar_qualidade,
        )
        combo_qualidade.pack(side="left")

        # --- grupo avançado recolhível: leitura por região ---
        frame_regiao_wrapper = ctk.CTkFrame(frame_leitura, corner_radius=0, fg_color=tema["fundo"])
        frame_regiao_wrapper.pack(fill="x", pady=(0, 8))

        frame_regiao_conteudo = ctk.CTkFrame(frame_regiao_wrapper, corner_radius=0, fg_color=tema["fundo"])

        estado_recolhivel = {"aberto": False}

        def alternar_regiao():
            if estado_recolhivel["aberto"]:
                frame_regiao_conteudo.pack_forget()
                botao_recolhivel.configure(text="▸ Leitura por região")
                estado_recolhivel["aberto"] = False
            else:
                frame_regiao_conteudo.pack(fill="x", pady=(8, 0))
                botao_recolhivel.configure(text="▾ Leitura por região")
                estado_recolhivel["aberto"] = True

        botao_recolhivel = ctk.CTkButton(
            frame_regiao_wrapper, text="▸ Leitura por região", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"], text_color=tema["texto"],
            font=(fonte, 13), anchor="w", command=alternar_regiao,
        )
        botao_recolhivel.pack(fill="x")

        ctk.CTkCheckBox(
            frame_regiao_conteudo, text="Usar leitura por região",
            variable=self.usar_ocr_regiao, corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento"],
            border_color=tema["borda_forte"], text_color=tema["texto"], font=(fonte, 13),
        ).pack(anchor="w", pady=(8, 8))

        label_retangulo = ctk.CTkLabel(
            frame_regiao_conteudo, text="", font=(fonte, 12),
            text_color=tema["texto_secundario"], anchor="w",
        )
        label_retangulo.pack(fill="x", pady=(0, 8))

        def atualizar_label_retangulo():
            label_retangulo.configure(
                text=(f"Região atual: x0={self.regiao_x0.get():.3f}  y0={self.regiao_y0.get():.3f}  "
                      f"x1={self.regiao_x1.get():.3f}  y1={self.regiao_y1.get():.3f}")
            )

        atualizar_label_retangulo()

        def selecionar_regiao():
            self._selecionar_regiao_visualmente()
            atualizar_label_retangulo()

        ctk.CTkButton(
            frame_regiao_conteudo, text="🖱 Selecionar região no PDF...", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"], text_color=tema["texto"],
            font=(fonte, 13), command=selecionar_regiao,
        ).pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(
            frame_regiao_conteudo,
            text="Os valores padrão são um chute inicial — calibre com um boleto real antes de usar.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w",
            justify="left", wraplength=520,
        ).pack(fill="x", pady=(0, 8))

        hairline(corpo)

        # ===================== TEXTO NO PDF =====================
        rotulo_secao(corpo, "Texto no PDF")
        frame_texto = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        frame_texto.pack(fill="x", padx=24)

        caixa_tipo_servico = ctk.CTkTextbox(
            frame_texto, height=64, corner_radius=0, fg_color=tema["superficie"],
            border_width=1, border_color=tema["borda"], text_color=tema["texto"], font=(fonte, 13),
        )
        texto_tipo_servico_atual = self.txt_tipo_servico.get("1.0", "end-1c")

        def atualizar_estado_tipo_servico(*_args):
            if self.modo_texto.get() == "rodape":
                caixa_tipo_servico.configure(state="normal")
            else:
                caixa_tipo_servico.configure(state="disabled")

        ctk.CTkRadioButton(
            frame_texto, text="Rodapé (código + nome + serviço)", value="rodape",
            variable=self.modo_texto, corner_radius=0,
            fg_color=tema["acento"], border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=atualizar_estado_tipo_servico,
        ).pack(anchor="w", pady=(0, 4))
        ctk.CTkRadioButton(
            frame_texto, text="Canto superior (só o código)", value="topo_esquerdo",
            variable=self.modo_texto, corner_radius=0,
            fg_color=tema["acento"], border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=atualizar_estado_tipo_servico,
        ).pack(anchor="w", pady=(0, 8))

        ctk.CTkCheckBox(
            frame_texto, text="Renomear o arquivo de saída com o código do condomínio",
            variable=self.renomear_com_codigo, corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento"],
            border_color=tema["borda_forte"], text_color=tema["texto"], font=(fonte, 13),
        ).pack(anchor="w", pady=(0, 4))
        ctk.CTkLabel(
            frame_texto, text="Ex: \"10002 - ARAUJO LIMA QUITADO 05.26.pdf\". Não altera o conteúdo, só o nome.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w",
            justify="left", wraplength=520,
        ).pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            frame_texto, text="Tipo de serviço", font=(fonte, 12),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", pady=(0, 4))
        caixa_tipo_servico.pack(fill="x", pady=(0, 8))
        caixa_tipo_servico.insert("1.0", texto_tipo_servico_atual)
        atualizar_estado_tipo_servico()

        linha_fonte = ctk.CTkFrame(frame_texto, corner_radius=0, fg_color=tema["fundo"])
        linha_fonte.pack(fill="x", pady=(0, 8))
        linha_fonte.columnconfigure(0, weight=1)
        linha_fonte.columnconfigure(1, weight=1)

        bloco_tamanho = ctk.CTkFrame(linha_fonte, corner_radius=0, fg_color=tema["fundo"])
        bloco_tamanho.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkLabel(
            bloco_tamanho, text="Tamanho da fonte", font=(fonte, 12),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkEntry(
            bloco_tamanho, textvariable=self.tamanho_fonte, corner_radius=0,
            fg_color=tema["superficie"], border_width=1, border_color=tema["borda"],
            text_color=tema["texto"], font=(fonte, 13),
        ).pack(fill="x")

        bloco_cor = ctk.CTkFrame(linha_fonte, corner_radius=0, fg_color=tema["fundo"])
        bloco_cor.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(
            bloco_cor, text="Cor (hex)", font=(fonte, 12),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkEntry(
            bloco_cor, textvariable=self.cor_texto, corner_radius=0,
            fg_color=tema["superficie"], border_width=1, border_color=tema["borda"],
            text_color=tema["texto"], font=(fonte, 13),
        ).pack(fill="x")

        # ===================== RODAPÉ: CANCELAR / SALVAR =====================
        # (frame já criado e empacotado com side="bottom" antes de `corpo`)

        def cancelar():
            self.cnpj_emitente.set(snapshot["cnpj_emitente"])
            self.usar_match_nome_arquivo.set(snapshot["usar_match_nome_arquivo"])
            self.usar_ocr.set(snapshot["usar_ocr"])
            self.dpi_ocr.set(snapshot["dpi_ocr"])
            self.usar_ocr_regiao.set(snapshot["usar_ocr_regiao"])
            self.regiao_x0.set(snapshot["regiao_x0"])
            self.regiao_y0.set(snapshot["regiao_y0"])
            self.regiao_x1.set(snapshot["regiao_x1"])
            self.regiao_y1.set(snapshot["regiao_y1"])
            self.modo_texto.set(snapshot["modo_texto"])
            self.tamanho_fonte.set(snapshot["tamanho_fonte"])
            self.cor_texto.set(snapshot["cor_texto"])
            self.renomear_com_codigo.set(snapshot["renomear_com_codigo"])
            self._atualizar_botoes_predefinicao()
            janela.destroy()

        def salvar():
            cnpj_raw = self.cnpj_emitente.get().strip()
            cnpj_norm = normalizar_cnpj(cnpj_raw)
            # Vazio é válido (predefinições sem emitente fixo, ex: Notas
            # Diversas); só bloqueia se foi preenchido e é inválido.
            if cnpj_raw and not cnpj_valido(cnpj_norm):
                messagebox.showerror(
                    "CNPJ inválido",
                    "O CNPJ da sua empresa não é válido. Confira os dígitos e tente novamente.",
                    parent=janela,
                )
                return

            if not re.match(r"^#[0-9a-fA-F]{6}$", self.cor_texto.get()):
                messagebox.showerror(
                    "Cor inválida",
                    "Cor inválida — use formato #RRGGBB, ex: #000000.",
                    parent=janela,
                )
                return

            novo_tipo_servico = caixa_tipo_servico.get("1.0", "end-1c")
            self.txt_tipo_servico.delete("1.0", "end")
            self.txt_tipo_servico.insert("1.0", novo_tipo_servico)

            perfil = self.config_app["predefinicoes"][self.predefinicao_ativa.get()]
            perfil["cnpj_emitente"] = cnpj_raw
            perfil["usar_match_nome_arquivo"] = self.usar_match_nome_arquivo.get()
            perfil["usar_ocr"] = self.usar_ocr.get()
            perfil["dpi_ocr"] = self.dpi_ocr.get()
            perfil["usar_ocr_regiao"] = self.usar_ocr_regiao.get()
            perfil["regiao_x0"] = self.regiao_x0.get()
            perfil["regiao_y0"] = self.regiao_y0.get()
            perfil["regiao_x1"] = self.regiao_x1.get()
            perfil["regiao_y1"] = self.regiao_y1.get()
            perfil["modo_texto"] = self.modo_texto.get()
            perfil["tamanho_fonte"] = self.tamanho_fonte.get()
            perfil["cor_texto"] = self.cor_texto.get()
            perfil["tipo_servico"] = novo_tipo_servico
            perfil["renomear_com_codigo"] = self.renomear_com_codigo.get()
            salvar_config(self.config_app)
            self._atualizar_botoes_predefinicao()

            janela.destroy()

        ctk.CTkButton(
            rodape, text="Cancelar", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"], text_color=tema["texto"],
            font=(fonte, 13), command=cancelar,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            rodape, text="Salvar", corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento_hover"],
            text_color=tema["sobre_acento"], border_width=0,
            font=(fonte, 13), command=salvar,
        ).pack(side="right")

        janela.protocol("WM_DELETE_WINDOW", cancelar)

    def _estilizar_ttk(self):
        """
        Estiliza o ttk.Notebook (abas), ttk.Progressbar e ttk.Treeview com as
        cores do tema atual. O Treeview detalhado é retrabalhado na Tarefa 4
        — aqui é só o básico para não ficar com cinza padrão do sistema.
        """
        tema = self.tema_atual
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        estilo.configure("TFrame", background=tema["fundo"])
        estilo.configure("TNotebook", background=tema["fundo"], borderwidth=0)
        estilo.configure(
            "TNotebook.Tab", background=tema["fundo"], foreground=tema["texto_secundario"],
            padding=(16, 8), borderwidth=0,
        )
        estilo.map(
            "TNotebook.Tab",
            background=[("selected", tema["superficie"])],
            foreground=[("selected", tema["texto"])],
        )
        estilo.configure(
            "TProgressbar", background=tema["acento"], troughcolor=tema["superficie"], borderwidth=0,
        )
        estilo.configure(
            "Treeview", background=tema["superficie"], fieldbackground=tema["superficie"],
            foreground=tema["texto"], borderwidth=0,
        )
        estilo.configure(
            "Treeview.Heading", background=tema["fundo"], foreground=tema["texto_secundario"],
        )

        # As abas de Cadastro e Logs ainda usam widgets ttk "clássicos" (Entry,
        # Button, Label, LabelFrame) — sem isso, o tema "clam" (necessário para
        # estilizar Notebook/Treeview acima) deixa esses widgets sempre claros,
        # destoando do resto do app no tema escuro.
        estilo.configure(
            "TLabel", background=tema["fundo"], foreground=tema["texto_secundario"],
        )
        estilo.configure(
            "TLabelframe", background=tema["fundo"], bordercolor=tema["borda"],
        )
        estilo.configure(
            "TLabelframe.Label", background=tema["fundo"], foreground=tema["texto_secundario"],
        )
        estilo.configure(
            "TEntry", fieldbackground=tema["superficie"], foreground=tema["texto"],
            bordercolor=tema["borda"], insertcolor=tema["texto"],
        )
        estilo.configure(
            "TButton", background=tema["superficie"], foreground=tema["texto"],
            bordercolor=tema["borda_forte"], focuscolor=tema["superficie"],
        )
        estilo.map(
            "TButton",
            background=[("active", tema["borda"]), ("pressed", tema["borda"])],
        )
        estilo.configure(
            "Vertical.TScrollbar", background=tema["superficie"], troughcolor=tema["fundo"],
            bordercolor=tema["borda"], arrowcolor=tema["texto_secundario"],
        )
        estilo.configure("TSeparator", background=tema["borda"])

        # Estilos Swiss usados nas abas Cadastro/Logs: legenda pequena
        # secundária, título grande e o único botão cobalto dessas abas.
        estilo.configure(
            "Caption.TLabel", background=tema["fundo"], foreground=tema["texto_secundario"],
        )
        estilo.configure(
            "Titulo.TLabel", background=tema["fundo"], foreground=tema["texto"],
        )
        estilo.configure(
            "Cobalto.TButton", background=tema["acento"], foreground=tema["sobre_acento"],
            bordercolor=tema["acento"], focuscolor=tema["acento"],
        )
        estilo.map(
            "Cobalto.TButton",
            background=[("active", tema["acento_hover"]), ("pressed", tema["acento_hover"])],
        )

        # Widgets tk "clássicos" (não-ttk) usados nas abas Cadastro/Logs — cor de
        # sistema não acompanha o tema, então fixamos explicitamente.
        for widget, cor_fundo, cor_texto in (
            (getattr(self, "log_historico", None), tema["superficie"], tema["texto"]),
        ):
            if widget is not None:
                try:
                    widget.configure(
                        bg=cor_fundo, fg=cor_texto,
                        insertbackground=cor_texto, selectbackground=tema["acento"],
                        selectforeground=tema["sobre_acento"],
                    )
                except Exception:
                    pass

    def _recolorir(self):
        """
        Percorre self._widgets_tema (registrados em _montar_aba_processar) e
        reaplica as cores do tema atual, além de re-estilizar o ttk. Chamado
        por aplicar_tema() para que o alternador de tema recolorir a tela na
        hora, sem remontar os widgets.
        """
        if not hasattr(self, "_widgets_tema"):
            return
        for widget, mapa in self._widgets_tema:
            for prop, chave in mapa.items():
                try:
                    widget.configure(**{prop: self.tema_atual[chave]})
                except Exception:
                    pass
        if hasattr(self, "botao_tema"):
            try:
                self.botao_tema.configure(text=self._texto_botao_tema())
            except Exception:
                pass
        self._estilizar_ttk()

    def _ler_retangulo_regiao(self):
        try:
            return (self.regiao_x0.get(), self.regiao_y0.get(),
                     self.regiao_x1.get(), self.regiao_y1.get())
        except tk.TclError:
            raise ValueError("Coordenadas da região devem ser números (0.0 a 1.0).")

    def _selecionar_regiao_visualmente(self):
        """
        Abre um PDF de exemplo, mostra a página inteira e deixa o usuário
        desenhar o retângulo da região com o mouse (clique e arraste).
        Ao confirmar, grava as coordenadas (frações 0.0–1.0 da página) nos
        campos x0,y0,x1,y1 e liga o OCR por região.
        """
        dono = (
            self._janela_config
            if getattr(self, "_janela_config", None) and self._janela_config.winfo_exists()
            else self
        )
        if not FITZ_DISPONIVEL:
            messagebox.showerror("Erro", "PyMuPDF não instalado. Instale: pip install pymupdf", parent=dono)
            return
        try:
            from PIL import ImageTk  # noqa: F401 — só valida disponibilidade aqui
        except ImportError:
            messagebox.showerror("Erro", "Pillow não instalado. Instale: pip install pillow", parent=dono)
            return

        caminho = filedialog.askopenfilename(
            title="Selecione um PDF de exemplo para calibrar a região",
            filetypes=[("PDF", "*.pdf")],
            parent=dono,
        )
        if not caminho:
            return

        try:
            doc = fitz.open(caminho)
            pagina = doc[0]
            pix = pagina.get_pixmap(dpi=150)
            img_original = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
        except Exception as e:
            messagebox.showerror("Erro ao abrir PDF", str(e), parent=dono)
            return

        # Imagem em tamanho real (sem espremer) — a janela tem barra de rolagem
        # pra caber em qualquer resolução de tela.
        largura_tela = img_original.width
        altura_tela = img_original.height

        janela = tk.Toplevel(self)
        try:
            self._montar_janela_selecao_regiao(janela, img_original, largura_tela, altura_tela)
        except Exception:
            janela.destroy()
            raise

    def _montar_janela_selecao_regiao(self, janela, img_original, largura_tela, altura_tela):
        from PIL import ImageTk

        janela.title("Selecione a região com o mouse")
        janela.resizable(True, True)
        janela.transient(self)
        janela.grab_set()

        ttk.Label(
            janela,
            text="Clique e arraste sobre a área com o CNPJ/nome do condomínio. Solte o mouse para ajustar. "
                 "Use as barras de rolagem se a página não couber na tela.",
            foreground="#555555",
            wraplength=780,
            justify="left",
        ).pack(padx=8, pady=(8, 4), fill="x")

        frame_canvas = ttk.Frame(janela)
        frame_canvas.pack(padx=8, pady=4, fill="both", expand=True)
        frame_canvas.rowconfigure(0, weight=1)
        frame_canvas.columnconfigure(0, weight=1)

        # Área visível da tela de seleção — limitada ao tamanho da imagem ou
        # a um teto razoável, o que for menor; o resto é acessado rolando.
        vista_largura = min(largura_tela, 820)
        vista_altura = min(altura_tela, 760)

        canvas = tk.Canvas(frame_canvas, width=vista_largura, height=vista_altura, cursor="crosshair",
                            highlightthickness=1, highlightbackground="#999999",
                            scrollregion=(0, 0, largura_tela, altura_tela))
        barra_v = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
        barra_h = ttk.Scrollbar(frame_canvas, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=barra_v.set, xscrollcommand=barra_h.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        barra_v.grid(row=0, column=1, sticky="ns")
        barra_h.grid(row=1, column=0, sticky="ew")

        img_tk = ImageTk.PhotoImage(img_original)
        canvas.create_image(0, 0, anchor="nw", image=img_tk)
        canvas.image = img_tk  # evita garbage collection

        def ao_rolar_vertical(event):
            canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        def ao_rolar_horizontal(event):
            canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")

        canvas.bind("<MouseWheel>", ao_rolar_vertical)
        canvas.bind("<Shift-MouseWheel>", ao_rolar_horizontal)

        estado = {
            "retangulo_id": None, "inicio": None,
            "x0": self.regiao_x0.get(), "y0": self.regiao_y0.get(),
            "x1": self.regiao_x1.get(), "y1": self.regiao_y1.get(),
        }

        label_coords = ttk.Label(janela, text="", font=("Consolas", 9))

        def atualizar_label():
            label_coords.configure(
                text=f"x0={estado['x0']:.3f}  y0={estado['y0']:.3f}  x1={estado['x1']:.3f}  y1={estado['y1']:.3f}"
            )

        # Se já houver um retângulo válido configurado, desenha ele de partida
        if 0 <= estado["x0"] < estado["x1"] <= 1 and 0 <= estado["y0"] < estado["y1"] <= 1:
            estado["retangulo_id"] = canvas.create_rectangle(
                estado["x0"] * largura_tela, estado["y0"] * altura_tela,
                estado["x1"] * largura_tela, estado["y1"] * altura_tela,
                outline="#e00000", width=2,
            )

        label_coords.pack(padx=8, pady=(0, 4))
        atualizar_label()

        def ao_pressionar(event):
            x, y = canvas.canvasx(event.x), canvas.canvasy(event.y)
            estado["inicio"] = (x, y)
            if estado["retangulo_id"] is not None:
                canvas.delete(estado["retangulo_id"])
            estado["retangulo_id"] = canvas.create_rectangle(
                x, y, x, y, outline="#e00000", width=2)

        def ao_arrastar(event):
            if estado["inicio"] is None:
                return
            x0, y0 = estado["inicio"]
            x, y = canvas.canvasx(event.x), canvas.canvasy(event.y)
            canvas.coords(estado["retangulo_id"], x0, y0, x, y)

        def ao_soltar(event):
            if estado["inicio"] is None:
                return
            x0_px, y0_px = estado["inicio"]
            x1_px, y1_px = canvas.canvasx(event.x), canvas.canvasy(event.y)
            x0_px, x1_px = sorted((max(0, min(x0_px, largura_tela)), max(0, min(x1_px, largura_tela))))
            y0_px, y1_px = sorted((max(0, min(y0_px, altura_tela)), max(0, min(y1_px, altura_tela))))
            estado["x0"] = x0_px / largura_tela
            estado["y0"] = y0_px / altura_tela
            estado["x1"] = x1_px / largura_tela
            estado["y1"] = y1_px / altura_tela
            estado["inicio"] = None
            atualizar_label()

        canvas.bind("<ButtonPress-1>", ao_pressionar)
        canvas.bind("<B1-Motion>", ao_arrastar)
        canvas.bind("<ButtonRelease-1>", ao_soltar)

        def confirmar():
            largura_min = 0.01
            if not (0 <= estado["x0"] and estado["x1"] <= 1 and 0 <= estado["y0"] and estado["y1"] <= 1
                    and (estado["x1"] - estado["x0"]) > largura_min
                    and (estado["y1"] - estado["y0"]) > largura_min):
                messagebox.showerror("Erro", "Desenhe um retângulo válido antes de confirmar.", parent=janela)
                return
            self.regiao_x0.set(round(estado["x0"], 3))
            self.regiao_y0.set(round(estado["y0"], 3))
            self.regiao_x1.set(round(estado["x1"], 3))
            self.regiao_y1.set(round(estado["y1"], 3))
            self.usar_ocr_regiao.set(True)
            janela.destroy()

        frame_botoes = ttk.Frame(janela)
        frame_botoes.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(frame_botoes, text="✓ Usar este recorte", command=confirmar).pack(side="left")
        ttk.Button(frame_botoes, text="Cancelar", command=janela.destroy).pack(side="left", padx=8)

    # --------------------------------------------------------
    #  PAINEL DE RESULTADO (Tela 3 — Tarefa 4)
    # --------------------------------------------------------
    @staticmethod
    def _formatar_tempo(tempo_segundos):
        """Formata segundos como "1m 04s" ou "12s" (mesma lógica do resumo do log)."""
        try:
            tempo_segundos = float(tempo_segundos)
        except (TypeError, ValueError):
            tempo_segundos = 0.0
        minutos = int(tempo_segundos // 60)
        segundos = int(tempo_segundos % 60)
        if minutos:
            return f"{minutos}m {segundos:02d}s"
        return f"{segundos}s"

    def mostrar_resultado(self, resultado):
        """
        Abre (ou reaproveita) o painel de Resultado — Tela 3. Recebe o dict
        `resultado` descrito no contrato da Tarefa 5 (ver docstring do
        arquivo/handoff): total, tempo_segundos, processados[], pendentes[].

        Esta função só monta a UI: as ações reais (cadastrar, escolher entre
        candidatos, abrir PDF) já estão implementadas em
        _acao_cadastrar_pendente/_acao_abrir_pdf_pendente/_acao_escolher_pendente,
        ligadas aos botões de cada linha das tabelas.
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        # referência ao dict renderizado, usada por ações (ex: reprocesso via
        # "Escolher") para atualizar o resultado in-place e re-renderizar
        self._resultado_atual = resultado

        processados = resultado.get("processados", []) or []
        pendentes = resultado.get("pendentes", []) or []
        total = resultado.get("total", len(processados) + len(pendentes))
        tempo_str = self._formatar_tempo(resultado.get("tempo_segundos", 0))

        if getattr(self, "_janela_resultado", None) is not None and self._janela_resultado.winfo_exists():
            for widget in self._janela_resultado.winfo_children():
                widget.destroy()
            janela = self._janela_resultado
            janela.focus_force()
        else:
            janela = ctk.CTkToplevel(self)
            self._janela_resultado = janela

        janela.title("Resultado do processamento")
        janela.geometry("820x680")
        janela.minsize(640, 480)
        janela.resizable(True, True)
        janela.configure(fg_color=tema["fundo"])
        janela.transient(self)

        # dados de cada pendente, indexados pelo iid da linha na tabela —
        # usado pelos botões de ação e pelo duplo clique para recuperar o
        # dict original (cnpj, nome_sugerido, candidatos, caminho, tipo...)
        self._pend_por_iid = {}

        # ===================== FAIXA DE CARTÕES =====================
        faixa = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        faixa.pack(fill="x", padx=24, pady=(24, 0))
        faixa.columnconfigure(0, weight=1)
        faixa.columnconfigure(2, weight=1)
        faixa.columnconfigure(4, weight=1)

        def montar_cartao(parent, coluna, caption, valor, cor_caption, cor_valor):
            bloco = ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"])
            bloco.grid(row=0, column=coluna, sticky="nsew", padx=16)
            ctk.CTkLabel(
                bloco, text=caption, font=(fonte, 11), text_color=cor_caption, anchor="w",
            ).pack(fill="x", anchor="w")
            ctk.CTkLabel(
                bloco, text=valor, font=(fonte, 40), text_color=cor_valor, anchor="w",
            ).pack(fill="x", anchor="w")

        def hairline_vertical(parent, coluna):
            linha = ctk.CTkFrame(parent, width=1, corner_radius=0, fg_color=tema["borda"])
            linha.grid(row=0, column=coluna, sticky="ns")

        montar_cartao(faixa, 0, "PROCESSADOS", str(len(processados)), tema["texto_secundario"], tema["texto"])
        hairline_vertical(faixa, 1)
        montar_cartao(faixa, 2, "PENDENTES", str(len(pendentes)), tema["acento"], tema["acento"])
        hairline_vertical(faixa, 3)
        montar_cartao(faixa, 4, "TEMPO", tempo_str, tema["texto_secundario"], tema["texto"])

        hairline = ctk.CTkFrame(janela, height=1, corner_radius=0, fg_color=tema["borda"])
        hairline.pack(fill="x", padx=24, pady=(24, 0))

        # ===================== RODAPÉ (ancorado primeiro) =====================
        rodape = ctk.CTkLabel(
            janela, text=f"{total} arquivo(s) no total.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w",
        )
        rodape.pack(side="bottom", fill="x", padx=24, pady=16)

        # ===================== ÁREA ROLÁVEL (pendentes + processados) =====================
        area = ctk.CTkScrollableFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        area.pack(side="top", fill="both", expand=True, padx=24, pady=(16, 0))

        # --- SEÇÃO PENDENTES ---
        ctk.CTkLabel(
            area, text="PENDENTES — PRECISAM DE AÇÃO", font=(fonte, 11, "bold"),
            text_color=tema["acento"], anchor="w",
        ).pack(fill="x", pady=(0, 8))

        if not pendentes:
            ctk.CTkLabel(
                area, text="Nenhum pendente — todos os arquivos foram codificados.",
                font=(fonte, 13), text_color=tema["texto_secundario"], anchor="w",
            ).pack(fill="x", pady=(0, 16))
        else:
            frame_tabela_pend = ctk.CTkFrame(area, corner_radius=0, fg_color=tema["fundo"])
            frame_tabela_pend.pack(fill="both", expand=False, pady=(0, 8))

            tabela_pend = ttk.Treeview(
                frame_tabela_pend, columns=("arquivo", "motivo"), show="headings", height=6,
            )
            tabela_pend.heading("arquivo", text="Arquivo")
            tabela_pend.heading("motivo", text="Motivo")
            tabela_pend.column("arquivo", width=280, anchor="w")
            tabela_pend.column("motivo", width=420, anchor="w")
            tabela_pend.pack(side="left", fill="both", expand=True)

            scroll_pend = ttk.Scrollbar(frame_tabela_pend, orient="vertical", command=tabela_pend.yview)
            tabela_pend.configure(yscrollcommand=scroll_pend.set)
            scroll_pend.pack(side="left", fill="y")

            for i, dados in enumerate(pendentes):
                iid = f"pend{i}"
                self._pend_por_iid[iid] = dados
                tabela_pend.insert(
                    "", "end", iid=iid,
                    values=(dados.get("arquivo", ""), dados.get("motivo", "")),
                )

            # --- botões de ação, operam sobre a linha selecionada ---
            frame_acoes = ctk.CTkFrame(area, corner_radius=0, fg_color=tema["fundo"])
            frame_acoes.pack(fill="x", pady=(0, 16))

            botao_cadastrar = ctk.CTkButton(
                frame_acoes, text="Cadastrar", corner_radius=0, state="disabled",
                fg_color=tema["borda"], hover_color=tema["acento_hover"],
                text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
                command=lambda: self._acao_cadastrar_pendente(self._pendente_selecionado(tabela_pend)),
            )
            botao_cadastrar.pack(side="left", padx=(0, 8))

            botao_escolher = ctk.CTkButton(
                frame_acoes, text="Escolher", corner_radius=0, state="disabled",
                fg_color=tema["borda"], hover_color=tema["acento_hover"],
                text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
                command=lambda: self._acao_escolher_pendente(self._pendente_selecionado(tabela_pend)),
            )
            botao_escolher.pack(side="left", padx=(0, 8))

            botao_abrir = ctk.CTkButton(
                frame_acoes, text="Abrir PDF", corner_radius=0, state="disabled",
                fg_color="transparent", hover_color=tema["superficie"],
                border_width=1, border_color=tema["borda_forte"], text_color=tema["texto"],
                font=(fonte, 13),
                command=lambda: self._acao_abrir_pdf_pendente(self._pendente_selecionado(tabela_pend)),
            )
            botao_abrir.pack(side="left")

            def atualizar_botoes_acao(_event=None):
                dados = self._pendente_selecionado(tabela_pend)
                if dados is None:
                    botao_cadastrar.configure(state="disabled", fg_color=tema["borda"])
                    botao_escolher.configure(state="disabled", fg_color=tema["borda"])
                    botao_abrir.configure(state="disabled")
                    return
                tipo = dados.get("tipo")
                cadastrar_ativo = tipo == "nao_cadastrado"
                escolher_ativo = tipo == "ambiguo"
                botao_cadastrar.configure(
                    state="normal" if cadastrar_ativo else "disabled",
                    fg_color=tema["acento"] if cadastrar_ativo else tema["borda"],
                )
                botao_escolher.configure(
                    state="normal" if escolher_ativo else "disabled",
                    fg_color=tema["acento"] if escolher_ativo else tema["borda"],
                )
                botao_abrir.configure(state="normal" if dados.get("caminho") else "disabled")

            def ao_duplo_clique(_event=None):
                dados = self._pendente_selecionado(tabela_pend)
                if dados is None:
                    return
                tipo = dados.get("tipo")
                if tipo == "nao_cadastrado":
                    self._acao_cadastrar_pendente(dados)
                elif tipo == "ambiguo":
                    self._acao_escolher_pendente(dados)
                elif dados.get("caminho"):
                    self._acao_abrir_pdf_pendente(dados)

            tabela_pend.bind("<<TreeviewSelect>>", atualizar_botoes_acao)
            tabela_pend.bind("<Double-1>", ao_duplo_clique)
            atualizar_botoes_acao()

        # --- SEÇÃO PROCESSADOS ---
        hairline2 = ctk.CTkFrame(area, height=1, corner_radius=0, fg_color=tema["borda"])
        hairline2.pack(fill="x", pady=(8, 16))

        ctk.CTkLabel(
            area, text="PROCESSADOS", font=(fonte, 11, "bold"),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", pady=(0, 8))

        if not processados:
            ctk.CTkLabel(
                area, text="Nenhum arquivo processado nesta sessão.",
                font=(fonte, 13), text_color=tema["texto_secundario"], anchor="w",
            ).pack(fill="x", pady=(0, 16))
        else:
            frame_tabela_proc = ctk.CTkFrame(area, corner_radius=0, fg_color=tema["fundo"])
            frame_tabela_proc.pack(fill="both", expand=False, pady=(0, 16))

            tabela_proc = ttk.Treeview(
                frame_tabela_proc, columns=("arquivo", "codigo", "origem"), show="headings", height=8,
            )
            tabela_proc.heading("arquivo", text="Arquivo")
            tabela_proc.heading("codigo", text="Código")
            tabela_proc.heading("origem", text="Origem")
            tabela_proc.column("arquivo", width=280, anchor="w")
            tabela_proc.column("codigo", width=90, anchor="center")
            tabela_proc.column("origem", width=280, anchor="w")
            tabela_proc.pack(side="left", fill="both", expand=True)

            scroll_proc = ttk.Scrollbar(frame_tabela_proc, orient="vertical", command=tabela_proc.yview)
            tabela_proc.configure(yscrollcommand=scroll_proc.set)
            scroll_proc.pack(side="left", fill="y")

            for i, dados in enumerate(processados):
                tabela_proc.insert(
                    "", "end", iid=f"proc{i}",
                    values=(dados.get("arquivo", ""), dados.get("codigo", ""), dados.get("origem", "")),
                )

        self._estilizar_ttk()
        janela.protocol("WM_DELETE_WINDOW", janela.destroy)

    def _pendente_selecionado(self, tabela_pend):
        """Devolve o dict do pendente correspondente à linha selecionada na
        tabela de pendentes, ou None se não houver seleção."""
        selecionado = tabela_pend.selection()
        if not selecionado:
            return None
        return self._pend_por_iid.get(selecionado[0])

    def _acao_cadastrar_pendente(self, dados):
        """Leva o usuário à aba de Cadastro com CNPJ/nome pré-preenchidos a
        partir do pendente selecionado, para ele completar o código."""
        if dados is None:
            return
        cnpj = dados.get("cnpj")
        self.form_cnpj.set(formatar_cnpj(cnpj) if cnpj else "")
        self.form_codigo.set("")
        self.form_nome.set(dados.get("nome_sugerido") or "")
        self.notebook.select(self.aba_cadastro)
        self.lift()
        self.focus_force()

    def _acao_escolher_pendente(self, dados):
        """Abre um popup para o usuário escolher, entre os CNPJs candidatos
        do pendente (caso ambíguo), qual é o condomínio de verdade."""
        if dados is None:
            return
        candidatos = dados.get("candidatos") or []
        if not candidatos:
            messagebox.showinfo(
                "Ação", "Este pendente não tem candidatos para escolher.",
                parent=self._janela_resultado,
            )
            return

        tema = self.tema_atual
        fonte = familia_fonte()

        popup = ctk.CTkToplevel(self._janela_resultado)
        popup.title("Escolher CNPJ")
        popup.geometry("480x420")
        popup.configure(fg_color=tema["fundo"])
        popup.transient(self._janela_resultado)
        popup.resizable(False, False)

        ctk.CTkLabel(
            popup, text="Qual CNPJ é o condomínio deste arquivo?",
            font=(fonte, 14, "bold"), text_color=tema["texto"], anchor="w",
        ).pack(fill="x", padx=20, pady=(20, 4))
        ctk.CTkLabel(
            popup, text=dados.get("arquivo", ""),
            font=(fonte, 11), text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 12))

        area = ctk.CTkScrollableFrame(popup, corner_radius=0, fg_color=tema["fundo"])
        area.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        var_escolha = tk.StringVar(value=candidatos[0])
        for cnpj_norm in candidatos:
            registro = self.cadastro.get(cnpj_norm)
            nome = registro.get("nome", "(não cadastrado)") if registro else "(não cadastrado)"
            texto = f"{formatar_cnpj(cnpj_norm)} — {nome}"
            ctk.CTkRadioButton(
                area, text=texto, variable=var_escolha, value=cnpj_norm,
                font=(fonte, 13), text_color=tema["texto"], fg_color=tema["acento"],
                corner_radius=0,
            ).pack(anchor="w", pady=6)

        frame_botoes = ctk.CTkFrame(popup, corner_radius=0, fg_color=tema["fundo"])
        frame_botoes.pack(fill="x", padx=20, pady=(0, 20))

        def confirmar():
            cnpj_escolhido = var_escolha.get()
            registro = self.cadastro.get(cnpj_escolhido)
            if registro is None:
                messagebox.showinfo(
                    "CNPJ não cadastrado",
                    "Esse CNPJ ainda não está cadastrado — abrindo o cadastro.",
                    parent=popup,
                )
                popup.destroy()
                dados_cad = dict(dados)
                dados_cad["cnpj"] = cnpj_escolhido
                self._acao_cadastrar_pendente(dados_cad)
                return
            popup.destroy()
            self._reprocessar_arquivo(dados, cnpj_escolhido)

        ctk.CTkButton(
            frame_botoes, text="Usar este CNPJ", corner_radius=0,
            fg_color=tema["acento"], hover_color=tema["acento_hover"],
            text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
            command=confirmar,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            frame_botoes, text="Cancelar", corner_radius=0, fg_color="transparent",
            hover_color=tema["superficie"], border_width=1, border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=popup.destroy,
        ).pack(side="left")

        popup.grab_set()

    def _reprocessar_arquivo(self, dados, cnpj_escolhido):
        """Reprocessa um único arquivo pendente com o CNPJ escolhido pelo
        usuário no popup de "Escolher", replicando o ramo de sucesso do loop
        de processamento (Tarefa 5a) para esse arquivo isolado. Ao terminar
        com sucesso, remove o pendente do painel de resultado."""
        ctx = getattr(self, "_ctx_processamento", None)
        if not ctx:
            messagebox.showerror(
                "Sem contexto de processamento",
                "Não há informações do último processamento nesta sessão. "
                "Rode o processamento novamente antes de tentar reprocessar.",
                parent=self._janela_resultado,
            )
            return

        # (validação de registro cadastrado já feita em confirmar(), no popup
        # de "Escolher", que é o único chamador desta função)
        registro = self.cadastro.get(cnpj_escolhido)
        codigo = registro["codigo"]
        condominio = registro["nome"]
        if ctx["modo"] == "rodape":
            texto_pdf = f"{codigo} {condominio} - {ctx['tipo_servico']}".strip()
        else:
            texto_pdf = codigo

        nome_saida = dados["arquivo"]
        if ctx.get("renomear"):
            nome_saida = nome_saida_com_codigo(dados["arquivo"], codigo)
        caminho_saida = os.path.join(ctx["saida"], nome_saida)

        try:
            processar_pdf(dados["caminho"], caminho_saida, texto_pdf, ctx["config"])
        except Exception as e:
            messagebox.showerror(
                "Erro ao reprocessar", f"Falha ao codificar '{dados.get('arquivo', '')}': {e}",
                parent=self._janela_resultado,
            )
            return

        messagebox.showinfo(
            "Arquivo codificado", f"Arquivo codificado com o código {codigo}.",
            parent=self._janela_resultado,
        )

        # move o pendente resolvido para "processados" no resultado
        # renderizado e re-renderiza o painel inteiro, para que cartões-
        # resumo, rodapé e as duas tabelas fiquem consistentes
        resultado = getattr(self, "_resultado_atual", None)
        if resultado is not None:
            pendentes_resultado = resultado.get("pendentes") or []
            for item in list(pendentes_resultado):
                if item is dados:
                    pendentes_resultado.remove(item)
                    break
            processados_resultado = resultado.setdefault("processados", [])
            processados_resultado.append({
                "arquivo": dados["arquivo"],
                "codigo": registro["codigo"],
                "origem": "pelo CNPJ (escolhido)",
            })
            self.mostrar_resultado(resultado)
        else:
            # fallback: sem o resultado renderizado disponível, só remove a
            # linha do pendente da tabela (comportamento antigo)
            for iid, dados_pend in list(self._pend_por_iid.items()):
                if dados_pend is dados:
                    del self._pend_por_iid[iid]
                    self._remover_linha_pendente(iid)
                    break

    def _remover_linha_pendente(self, iid):
        """Remove a linha `iid` da tabela de pendentes do painel de
        Resultado, se ela ainda existir na tela."""
        janela = getattr(self, "_janela_resultado", None)
        if janela is None or not janela.winfo_exists():
            return
        for widget in janela.winfo_children():
            self._remover_treeview_iid_recursivo(widget, iid)

    def _remover_treeview_iid_recursivo(self, widget, iid):
        if isinstance(widget, ttk.Treeview):
            if widget.exists(iid):
                widget.delete(iid)
            return
        for filho in widget.winfo_children():
            self._remover_treeview_iid_recursivo(filho, iid)

    def _acao_abrir_pdf_pendente(self, dados):
        """Abre o PDF original do pendente selecionado no visualizador padrão
        do sistema, para o funcionário conferir o conteúdo."""
        if dados is None:
            return
        caminho = dados.get("caminho")
        if not caminho:
            return
        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror(
                "Erro ao abrir PDF", f"Não foi possível abrir o arquivo:\n{e}",
                parent=self._janela_resultado,
            )

    # --------------------------------------------------------
    #  ABA 3 — LOGS
    # --------------------------------------------------------
    def _montar_aba_logs(self, parent):
        fonte = familia_fonte()

        titulo = ttk.Label(parent, text="Logs", style="Titulo.TLabel", font=(fonte, 20))
        titulo.pack(anchor="w", padx=24, pady=(20, 0))

        self.label_caminho_log = ttk.Label(
            parent, text=self.caminho_log, style="Caption.TLabel", font=("Consolas", 10),
        )
        self.label_caminho_log.pack(anchor="w", padx=24, pady=(4, 16))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=24)

        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill="x", padx=24, pady=12)
        ttk.Button(frame_botoes, text="Atualizar", command=self._carregar_log_em_tela).pack(side="left")
        ttk.Button(frame_botoes, text="Limpar log", command=self._limpar_log).pack(side="left", padx=8)
        ttk.Button(frame_botoes, text="Abrir pasta", command=self._abrir_pasta_log).pack(side="left")

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=24)

        frame_log = ttk.Frame(parent)
        frame_log.pack(fill="both", expand=True, padx=24, pady=16)

        self.log_historico = tk.Text(
            frame_log, state="disabled", wrap="word", font=("Consolas", 10),
            borderwidth=0, highlightthickness=0,
        )
        self.log_historico.pack(fill="both", expand=True, side="left")

        scroll_log = ttk.Scrollbar(frame_log, orient="vertical", command=self.log_historico.yview)
        self.log_historico.configure(yscrollcommand=scroll_log.set)
        scroll_log.pack(side="left", fill="y")

    def _carregar_log_em_tela(self):
        """Lê o arquivo .log e exibe na aba Logs."""
        self.log_historico.configure(state="normal")
        self.log_historico.delete("1.0", "end")
        if os.path.isfile(self.caminho_log):
            with open(self.caminho_log, "r", encoding="utf-8") as f:
                conteudo = f.read()
            self.log_historico.insert("end", conteudo)
            self.log_historico.see("end")
        else:
            self.log_historico.insert("end", "(nenhum log registrado ainda)")
        self.log_historico.configure(state="disabled")

    def _limpar_log(self):
        if not messagebox.askyesno("Confirmar", "Apagar todo o histórico de logs?"):
            return
        try:
            with open(self.caminho_log, "w", encoding="utf-8") as f:
                f.write("")
            self._carregar_log_em_tela()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível limpar o log:\n{e}")

    def _abrir_pasta_log(self):
        pasta = os.path.dirname(self.caminho_log)
        try:
            os.startfile(pasta)
        except Exception:
            messagebox.showinfo("Pasta do log", pasta)

    def _salvar_sessao_no_log(self, linhas_sessao: list[str]):
        """Grava as linhas da sessão de processamento no arquivo .log."""
        try:
            with open(self.caminho_log, "a", encoding="utf-8") as f:
                f.write("\n".join(linhas_sessao) + "\n\n")
        except Exception as e:
            # Falha silenciosa no log não deve travar o processamento
            print(f"[AVISO] Não foi possível salvar o log: {e}")

    def _escolher_pasta_entrada(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs originais")
        if pasta:
            self.pasta_entrada.set(pasta)
            self._atualizar_botao_processar()

    def _escolher_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
        if pasta:
            self.pasta_saida.set(pasta)

    def _log(self, mensagem):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _iniciar_processamento(self):
        entrada = self.pasta_entrada.get().strip()
        saida = self.pasta_saida.get().strip()

        if not entrada or not os.path.isdir(entrada):
            messagebox.showerror("Erro", "Selecione uma pasta de entrada válida.")
            return
        if not saida:
            messagebox.showerror("Erro", "Selecione uma pasta de saída.")
            return

        cnpj_emitente_raw = self.cnpj_emitente.get().strip()
        cnpj_emitente_norm = normalizar_cnpj(cnpj_emitente_raw)
        # Vazio é válido (predefinições sem emitente fixo, ex: Notas Diversas
        # — lotes de administradoras variadas). Só bloqueia se foi preenchido
        # mas o CNPJ digitado é inválido.
        if cnpj_emitente_raw and len(cnpj_emitente_norm) != 14:
            messagebox.showerror("Erro", "CNPJ da empresa (campo 3) inválido.")
            return

        try:
            tamanho = int(self.tamanho_fonte.get())
        except ValueError:
            messagebox.showerror("Erro", "Tamanho da fonte deve ser um número.")
            return

        if not self.cadastro:
            if not messagebox.askyesno(
                "Cadastro vazio",
                "O cadastro de condomínios está vazio. Nenhum PDF vai casar com nenhum "
                "código. Deseja continuar mesmo assim (todos ficarão pendentes)?"
            ):
                return

        usar_ocr_regiao = self.usar_ocr_regiao.get() and OCR_DISPONIVEL
        retangulo_regiao = None
        if usar_ocr_regiao:
            try:
                retangulo_regiao = self._ler_retangulo_regiao()
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
                return

        # Ajuste que não é editável pelo modal — segue sempre o que a
        # predefinição ativa define (não faz sentido "por sessão").
        perfil_ativo = self.config_app["predefinicoes"][self.predefinicao_ativa.get()]
        preferir_cadastrado_em_ambiguo = perfil_ativo["preferir_cadastrado_em_ambiguo"]

        self.botao_iniciar.configure(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(
            target=self._processar_em_thread,
            args=(entrada, saida, tamanho, cnpj_emitente_norm, self.usar_ocr.get(), int(self.dpi_ocr.get()),
                  self.usar_match_nome_arquivo.get(), usar_ocr_regiao, retangulo_regiao,
                  preferir_cadastrado_em_ambiguo),
            daemon=True,
        )
        thread.start()

    def _processar_em_thread(self, entrada, saida, tamanho, cnpj_emitente_norm, usar_ocr, dpi,
                              usar_match_nome, usar_ocr_regiao, retangulo_regiao,
                              preferir_cadastrado_em_ambiguo):
        modo = self.modo_texto.get()
        tipo_servico = self.txt_tipo_servico.get("1.0", "end-1c").strip()
        cor = self.cor_texto.get().strip() or "#000000"
        renomear = self.renomear_com_codigo.get()

        if modo == "rodape":
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 0, "y": 90, "centralizado": True}
        else:
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 120, "y": 815, "centralizado": False}

        self._ctx_processamento = {"saida": saida, "config": config, "modo": modo,
                                   "tipo_servico": tipo_servico, "renomear": renomear}

        arquivos = [f for f in os.listdir(entrada) if f.lower().endswith(".pdf")]
        if not arquivos:
            self.after(0, lambda: self._log("Nenhum arquivo PDF encontrado."))
            self.after(0, lambda: self.botao_iniciar.configure(state="normal"))
            return

        self.after(0, lambda: self.barra_progresso.configure(maximum=len(arquivos), value=0))

        # Cabeçalho da sessão
        agora = datetime.datetime.now()
        cabecalho = (
            f"{'='*60}\n"
            f"  SESSÃO: {agora.strftime('%d/%m/%Y %H:%M:%S')}\n"
            f"  Pasta:  {entrada}\n"
            f"  Match por nome de arquivo: {'Sim' if usar_match_nome else 'Não'}\n"
            f"  OCR:    {'Sim (DPI ' + str(dpi) + ')' if usar_ocr and OCR_DISPONIVEL else 'Não'}\n"
            f"  OCR por região: {'Sim' if usar_ocr_regiao else 'Não'}\n"
            f"{'='*60}"
        )
        self.after(0, lambda: self._log(cabecalho))
        self.after(0, lambda: self._log(f"{len(arquivos)} arquivos encontrados.\n"))

        linhas_log = [cabecalho, f"{len(arquivos)} arquivos encontrados.\n"]

        sucesso = 0
        total_ocr = 0
        total_match_nome = 0
        pendentes = []  # (nome_arquivo, motivo)
        res_processados = []
        res_pendentes = []
        inicio = time.time()

        for idx, nome in enumerate(arquivos, 1):
            caminho_entrada_pdf = os.path.join(entrada, nome)
            caminho_saida_pdf = os.path.join(saida, nome)

            try:
                cnpj = None
                texto = ""
                sufixo_origem = ""
                origem_humana = "pelo CNPJ do boleto"

                # --- 1) Tenta casar pelo nome do arquivo, antes de abrir o PDF
                #     (código exato no nome, ex: F&F; senão fuzzy pelo nome,
                #     ex: FedCorp — ver buscar_por_nome_arquivo) ---
                if usar_match_nome:
                    cnpj_nome, resultado = buscar_por_nome_arquivo(nome, self.cadastro)
                    if cnpj_nome is not None:
                        cnpj = cnpj_nome
                        sufixo_origem = f" (via {resultado})"
                        origem_humana = ("pelo código no nome do arquivo" if resultado == "código no nome do arquivo"
                                          else "pelo nome do arquivo")
                        total_match_nome += 1

                # --- 2) Se não casou pelo nome, extrai do conteúdo do PDF ---
                candidatos = [cnpj] if cnpj else []
                if not candidatos:
                    texto = extrair_texto_pdf(caminho_entrada_pdf)
                    usado_ocr = False
                    dpi_usado = None

                    precisa_ocr = len(texto.strip()) < LIMITE_TEXTO_MINIMO
                    if precisa_ocr and usar_ocr and OCR_DISPONIVEL:
                        # 2a) OCR por região primeiro (mais rápido, se configurado)
                        if usar_ocr_regiao:
                            texto_regiao = extrair_texto_ocr_regiao(caminho_entrada_pdf, retangulo_regiao, dpi=dpi)
                            candidatos_regiao = extrair_cnpj_tomador(texto_regiao, cnpj_emitente_norm)
                            if len(candidatos_regiao) == 1:
                                texto = texto_regiao
                                usado_ocr = True
                                dpi_usado = dpi
                                sufixo_origem = " (via OCR de região)"
                                origem_humana = "pelo CNPJ do boleto"
                                total_ocr += 1

                        # 2b) Se a região não resolveu, OCR de página inteira
                        if not usado_ocr:
                            texto_ocr = extrair_texto_ocr(caminho_entrada_pdf, dpi=dpi)
                            if len(texto_ocr.strip()) > len(texto.strip()):
                                texto = texto_ocr
                                usado_ocr = True
                                dpi_usado = dpi
                                sufixo_origem = " (via OCR)"
                                origem_humana = "pelo CNPJ do boleto"
                                total_ocr += 1

                    candidatos = extrair_cnpj_tomador(texto, cnpj_emitente_norm)

                    # 2b.1) Lote sem emitente fixo (ex: Notas Diversas) — mais de
                    #       um candidato costuma ser o emitente da nota (não
                    #       cadastrado) + o condomínio tomador (cadastrado). Se
                    #       sobrar exatamente um candidato já cadastrado, usa ele.
                    if preferir_cadastrado_em_ambiguo:
                        desempatados = desempatar_por_cadastro(candidatos, self.cadastro)
                        if desempatados != candidatos:
                            candidatos = desempatados
                            sufixo_origem += " (desempate: CNPJ cadastrado)"

                    # 2c) Nenhum CNPJ válido achado via OCR — re-tenta em DPI maior
                    #     (descarta leituras com dígito verificador errado, então um
                    #     resultado vazio pode ser fruto de OCR ruim, não de PDF sem CNPJ)
                    if not candidatos and usado_ocr and usar_ocr and OCR_DISPONIVEL:
                        proximo = proximo_dpi_maior(dpi_usado)
                        if proximo:
                            texto_retry = extrair_texto_ocr(caminho_entrada_pdf, dpi=proximo)
                            candidatos_retry = extrair_cnpj_tomador(texto_retry, cnpj_emitente_norm)
                            if candidatos_retry:
                                texto = texto_retry
                                candidatos = candidatos_retry
                                sufixo_origem = f" (via OCR, re-tentativa DPI {proximo})"
                                origem_humana = "pelo CNPJ (releitura)"

                if len(candidatos) == 0:
                    pendentes.append((nome, "CNPJ do tomador não encontrado no PDF" + sufixo_origem))
                    msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — CNPJ não encontrado{sufixo_origem}"
                    res_pendentes.append({
                        "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "nao_lido",
                        "motivo": "Não foi possível ler", "cnpj": None, "nome_sugerido": None,
                        "candidatos": None})
                elif len(candidatos) > 1:
                    lista = ", ".join(formatar_cnpj(c) for c in candidatos)
                    pendentes.append((nome, f"CNPJ ambíguo: {lista}" + sufixo_origem))
                    msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — CNPJ ambíguo ({lista}){sufixo_origem}"
                    res_pendentes.append({
                        "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "ambiguo",
                        "motivo": "Dois CNPJs possíveis", "cnpj": None, "nome_sugerido": None,
                        "candidatos": list(candidatos)})
                else:
                    cnpj = candidatos[0]
                    registro = self.cadastro.get(cnpj)
                    if registro is None:
                        nome_sugerido = sugerir_nome_condominio(texto) if texto else ""
                        detalhe = f"CNPJ {formatar_cnpj(cnpj)} não cadastrado"
                        if nome_sugerido:
                            detalhe += f" (nome sugerido: {nome_sugerido})"
                        detalhe += sufixo_origem
                        pendentes.append((nome, detalhe))
                        msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — {detalhe}"
                        res_pendentes.append({
                            "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "nao_cadastrado",
                            "motivo": "CNPJ não cadastrado", "cnpj": cnpj, "nome_sugerido": (nome_sugerido or None),
                            "candidatos": None})
                    else:
                        codigo = registro["codigo"]
                        condominio = registro["nome"]
                        if modo == "rodape":
                            texto_pdf = f"{codigo} {condominio} - {tipo_servico}".strip()
                        else:
                            texto_pdf = codigo
                        if renomear:
                            caminho_saida_pdf = os.path.join(saida, nome_saida_com_codigo(nome, codigo))
                        processar_pdf(caminho_entrada_pdf, caminho_saida_pdf, texto_pdf, config)
                        sucesso += 1
                        texto_pdf_log = texto_pdf.replace("\n", " / ")
                        msg = f"[{idx}/{len(arquivos)}] ✓ {nome} → '{texto_pdf_log}'{sufixo_origem}"
                        res_processados.append({
                            "arquivo": nome, "codigo": codigo, "origem": origem_humana})

            except Exception as e:
                pendentes.append((nome, f"Erro inesperado: {e}"))
                msg = f"[{idx}/{len(arquivos)}] ✗ {nome} — erro: {e}"
                res_pendentes.append({
                    "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "erro",
                    "motivo": "Erro ao processar", "cnpj": None, "nome_sugerido": None,
                    "candidatos": None})

            linhas_log.append(msg)
            self.after(0, lambda m=msg: self._log(m))
            self.after(0, lambda v=idx: self.barra_progresso.configure(value=v))

        # Resumo final
        tempo_total = time.time() - inicio
        minutos = int(tempo_total // 60)
        segundos = int(tempo_total % 60)
        tempo_str = f"{minutos}m {segundos}s" if minutos else f"{segundos}s"

        resumo = (
            f"\nConcluído! {sucesso}/{len(arquivos)} processados com sucesso."
            + (f" | Via nome do arquivo: {total_match_nome}" if usar_match_nome else "")
            + (f" | Via OCR: {total_ocr}" if usar_ocr and OCR_DISPONIVEL else "")
            + f" | Tempo: {tempo_str}"
        )
        if pendentes:
            resumo += f"\n{len(pendentes)} arquivo(s) pendente(s) — não foram modificados:\n"
            for nome, motivo in pendentes:
                resumo += f"  • {nome}: {motivo}\n"
            resumo += "\nCadastre os CNPJs faltantes na aba 2 e rode novamente."

        linhas_log.append(resumo)

        self.after(0, lambda: self._log(resumo))
        self.after(0, lambda: self._salvar_sessao_no_log(linhas_log))
        self.after(0, lambda: self._carregar_log_em_tela())
        self.after(0, lambda: self.botao_iniciar.configure(state="normal"))

        resultado = {
            "total": len(arquivos),
            "tempo_segundos": tempo_total,
            "processados": res_processados,
            "pendentes": res_pendentes,
        }
        self.after(0, lambda: self.mostrar_resultado(resultado))


if __name__ == "__main__":
    app = App()
    app.mainloop()
