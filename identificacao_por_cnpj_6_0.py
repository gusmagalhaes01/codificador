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
import io
import json
import os
import re
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

DEFAULTS_CONFIG = {
    "cnpj_emitente": "13.736.666/0001-54",
    "usar_match_nome_arquivo": True,
    "usar_ocr": False,
    "dpi_ocr": 200,
    "usar_ocr_regiao": False,
    "regiao_x0": 0.0,
    "regiao_y0": 0.15,
    "regiao_x1": 1.0,
    "regiao_y1": 0.35,
    "modo_texto": "topo_esquerdo",
    "tipo_servico": "CIPAA",
    "tamanho_fonte": "14",
    "cor_texto": "#000000",
    "tema": "auto",  # "auto" | "claro" | "escuro"
}


def _gravar_erros_log(detalhes):
    """Faz o append de `detalhes` em erros.log (ao lado do script), com
    timestamp. Usada tanto por _registrar_erro_config quanto pelo handler
    global de exceções (App.report_callback_exception) — mesmo formato nos
    dois casos. Nunca lança exceção."""
    try:
        pasta_script = os.path.dirname(os.path.abspath(__file__))
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
    dos defaults. Se existir mas estiver corrompido, loga a exceção em
    erros.log e devolve uma cópia dos defaults — nunca lança exceção.
    Sempre mescla sobre os defaults, então um config parcial não quebra.
    """
    pasta_script = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(pasta_script, NOME_CONFIG_PADRAO)

    config = dict(DEFAULTS_CONFIG)
    if not os.path.isfile(caminho):
        return config

    try:
        with open(caminho, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if isinstance(dados, dict):
            config.update(dados)
    except Exception:
        import traceback
        _registrar_erro_config(traceback.format_exc())
        return dict(DEFAULTS_CONFIG)

    return config


def salvar_config(config):
    """Grava `config` como JSON em config.json (ao lado do script). Falha ao
    salvar não deve travar o app — loga em erros.log e segue."""
    pasta_script = os.path.dirname(os.path.abspath(__file__))
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
    CNPJ_FLEX = r"(\d{2}[\s.]?\d{3}[\s.]?\d{3}[\s/]?\d{4}[\s-]?\d{2})"

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


def normalizar_texto_busca(texto):
    """Maiúsculas, sem acento, sem dígitos (datas/códigos), separadores viram espaço."""
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[_\-.]", " ", texto)
    texto = re.sub(r"\d", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    return texto


def buscar_por_nome_arquivo(nome_arquivo, cadastro):
    """
    Tenta casar o nome do arquivo (ex: "ARAUJO_LIMA_QUITADO_05_26.pdf") com o
    nome de algum condomínio cadastrado, via fuzzy match.

    Retorna (cnpj_norm, score) se achar um match único e confiável, ou
    (None, motivo) se não achar nada ou o resultado ficar ambíguo — nesse
    caso o chamador deve cair para a extração de CNPJ do conteúdo do PDF.
    """
    alvo = normalizar_texto_busca(os.path.splitext(nome_arquivo)[0])
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

    return melhor_cnpj, melhor_score


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
        self.title("Identificação de PDFs por CNPJ")
        self.geometry("780x680")
        self.minsize(620, 420)
        self.resizable(True, True)

        # Caminho da planilha de cadastro e do log (ficam ao lado deste script)
        pasta_script = os.path.dirname(os.path.abspath(__file__))
        self.caminho_planilha = os.path.join(pasta_script, NOME_ARQUIVO_PADRAO)
        self.caminho_log = os.path.join(pasta_script, NOME_LOG_PADRAO)
        self.cadastro = carregar_cadastro(self.caminho_planilha)

        # Variáveis - aba processamento (inicializadas a partir da config
        # persistida, não mais de literais hardcoded — assim sobrevivem
        # entre sessões)
        self.pasta_entrada = tk.StringVar()
        self.pasta_saida = tk.StringVar()
        self.cnpj_emitente = tk.StringVar(value=self.config_app["cnpj_emitente"])
        self.modo_texto = tk.StringVar(value=self.config_app["modo_texto"])
        self.tamanho_fonte = tk.StringVar(value=self.config_app["tamanho_fonte"])
        self.cor_texto = tk.StringVar(value=self.config_app["cor_texto"])
        self.usar_ocr = tk.BooleanVar(value=self.config_app["usar_ocr"])
        self.dpi_ocr = tk.IntVar(value=self.config_app["dpi_ocr"])

        # OCR por região (recorte) — opcional, calibrado pelo usuário
        self.usar_ocr_regiao = tk.BooleanVar(value=self.config_app["usar_ocr_regiao"])
        self.regiao_x0 = tk.DoubleVar(value=self.config_app["regiao_x0"])
        self.regiao_y0 = tk.DoubleVar(value=self.config_app["regiao_y0"])
        self.regiao_x1 = tk.DoubleVar(value=self.config_app["regiao_x1"])
        self.regiao_y1 = tk.DoubleVar(value=self.config_app["regiao_y1"])

        # Match por nome de arquivo (evita OCR na maioria dos casos)
        self.usar_match_nome_arquivo = tk.BooleanVar(value=self.config_app["usar_match_nome_arquivo"])

        # Variáveis - aba cadastro (formulário)
        self.form_cnpj = tk.StringVar()
        self.form_codigo = tk.StringVar()
        self.form_nome = tk.StringVar()

        self._montar_interface()
        self._atualizar_tabela_cadastro()
        self._carregar_log_em_tela()

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

    # --------------------------------------------------------
    #  ABA 2 — CADASTRO
    # --------------------------------------------------------
    def _montar_aba_cadastro(self, parent):
        pad = {"padx": 10, "pady": 6}

        info = ttk.Label(
            parent,
            text=f"Planilha: {self.caminho_planilha}",
            foreground="#555555",
        )
        info.pack(anchor="w", **pad)

        # --- Formulário ---
        frame_form = ttk.LabelFrame(parent, text="Adicionar / Editar condomínio")
        frame_form.pack(fill="x", **pad)

        linha1 = ttk.Frame(frame_form)
        linha1.pack(fill="x", padx=8, pady=4)
        ttk.Label(linha1, text="CNPJ:", width=10).pack(side="left")
        ttk.Entry(linha1, textvariable=self.form_cnpj, width=25).pack(side="left", padx=(0, 20))
        ttk.Label(linha1, text="Código:", width=10).pack(side="left")
        ttk.Entry(linha1, textvariable=self.form_codigo, width=15).pack(side="left")

        linha2 = ttk.Frame(frame_form)
        linha2.pack(fill="x", padx=8, pady=4)
        ttk.Label(linha2, text="Nome:", width=10).pack(side="left")
        ttk.Entry(linha2, textvariable=self.form_nome, width=50).pack(side="left", fill="x", expand=True)

        linha3 = ttk.Frame(frame_form)
        linha3.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(linha3, text="Adicionar / Atualizar", command=self._adicionar_ou_atualizar).pack(side="left")
        ttk.Button(linha3, text="Limpar campos", command=self._limpar_form).pack(side="left", padx=8)

        # --- Tabela ---
        frame_tabela = ttk.LabelFrame(parent, text="Condomínios cadastrados")
        frame_tabela.pack(fill="both", expand=True, **pad)

        colunas = ("cnpj", "codigo", "nome")
        self.tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=14)
        self.tabela.heading("cnpj", text="CNPJ")
        self.tabela.heading("codigo", text="Código")
        self.tabela.heading("nome", text="Nome do Condomínio")
        self.tabela.column("cnpj", width=150, anchor="center")
        self.tabela.column("codigo", width=80, anchor="center")
        self.tabela.column("nome", width=380, anchor="w")
        self.tabela.pack(fill="both", expand=True, padx=8, pady=8, side="left")
        self.tabela.bind("<<TreeviewSelect>>", self._selecionar_linha)

        scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="left", fill="y")

        # --- Botões de baixo ---
        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill="x", **pad)
        ttk.Button(frame_botoes, text="Remover selecionado", command=self._remover_selecionado).pack(side="left")
        ttk.Button(frame_botoes, text="Importar planilha existente...", command=self._importar_planilha).pack(side="left", padx=8)
        ttk.Button(frame_botoes, text="Salvar planilha agora", command=self._salvar_planilha).pack(side="left")

    def _atualizar_tabela_cadastro(self):
        self.tabela.delete(*self.tabela.get_children())
        for cnpj_norm, dados in sorted(self.cadastro.items(), key=lambda kv: kv[1]["codigo"]):
            self.tabela.insert("", "end", iid=cnpj_norm,
                                values=(formatar_cnpj(cnpj_norm), dados["codigo"], dados["nome"]))

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
        self.txt_tipo_servico.insert("1.0", self.config_app["tipo_servico"])

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
        }

        janela = ctk.CTkToplevel(self)
        self._janela_config = janela
        janela.title("Configurações")
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
            janela.destroy()

        def salvar():
            cnpj_norm = normalizar_cnpj(self.cnpj_emitente.get())
            if not cnpj_valido(cnpj_norm):
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

            self.config_app["cnpj_emitente"] = self.cnpj_emitente.get()
            self.config_app["usar_match_nome_arquivo"] = self.usar_match_nome_arquivo.get()
            self.config_app["usar_ocr"] = self.usar_ocr.get()
            self.config_app["dpi_ocr"] = self.dpi_ocr.get()
            self.config_app["usar_ocr_regiao"] = self.usar_ocr_regiao.get()
            self.config_app["regiao_x0"] = self.regiao_x0.get()
            self.config_app["regiao_y0"] = self.regiao_y0.get()
            self.config_app["regiao_x1"] = self.regiao_x1.get()
            self.config_app["regiao_y1"] = self.regiao_y1.get()
            self.config_app["modo_texto"] = self.modo_texto.get()
            self.config_app["tamanho_fonte"] = self.tamanho_fonte.get()
            self.config_app["cor_texto"] = self.cor_texto.get()
            self.config_app["tipo_servico"] = novo_tipo_servico
            salvar_config(self.config_app)

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

        caminho_saida = os.path.join(ctx["saida"], dados["arquivo"])

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
        pad = {"padx": 10, "pady": 6}

        frame_info = ttk.Frame(parent)
        frame_info.pack(fill="x", **pad)
        self.label_caminho_log = ttk.Label(
            frame_info,
            text=f"Arquivo: {self.caminho_log}",
            foreground="#555555",
        )
        self.label_caminho_log.pack(side="left", anchor="w")

        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Button(frame_botoes, text="🔄 Atualizar", command=self._carregar_log_em_tela).pack(side="left")
        ttk.Button(frame_botoes, text="🗑 Limpar log", command=self._limpar_log).pack(side="left", padx=8)
        ttk.Button(frame_botoes, text="📂 Abrir pasta", command=self._abrir_pasta_log).pack(side="left")

        frame_log = ttk.LabelFrame(parent, text="Histórico de sessões")
        frame_log.pack(fill="both", expand=True, **pad)

        self.log_historico = tk.Text(frame_log, state="disabled", wrap="word", font=("Consolas", 9))
        self.log_historico.pack(fill="both", expand=True, side="left", padx=8, pady=8)

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

        cnpj_emitente_norm = normalizar_cnpj(self.cnpj_emitente.get())
        if len(cnpj_emitente_norm) != 14:
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

        self.botao_iniciar.configure(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(
            target=self._processar_em_thread,
            args=(entrada, saida, tamanho, cnpj_emitente_norm, self.usar_ocr.get(), int(self.dpi_ocr.get()),
                  self.usar_match_nome_arquivo.get(), usar_ocr_regiao, retangulo_regiao),
            daemon=True,
        )
        thread.start()

    def _processar_em_thread(self, entrada, saida, tamanho, cnpj_emitente_norm, usar_ocr, dpi,
                              usar_match_nome, usar_ocr_regiao, retangulo_regiao):
        modo = self.modo_texto.get()
        tipo_servico = self.txt_tipo_servico.get("1.0", "end-1c").strip()
        cor = self.cor_texto.get().strip() or "#000000"

        if modo == "rodape":
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 0, "y": 90, "centralizado": True}
        else:
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 120, "y": 815, "centralizado": False}

        self._ctx_processamento = {"saida": saida, "config": config, "modo": modo, "tipo_servico": tipo_servico}

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

                # --- 1) Tenta casar pelo nome do arquivo, antes de abrir o PDF ---
                if usar_match_nome:
                    cnpj_nome, resultado = buscar_por_nome_arquivo(nome, self.cadastro)
                    if cnpj_nome is not None:
                        cnpj = cnpj_nome
                        sufixo_origem = f" (via nome do arquivo, score {resultado:.2f})"
                        origem_humana = "pelo nome do arquivo"
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
