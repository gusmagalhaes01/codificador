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
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from difflib import SequenceMatcher

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

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

#  Reserva opcional: modelos PP-OCR (do PaddleOCR) rodando em ONNX. NÃO entra
#  no requirements.txt nem no .spec — quem tiver instalado na própria máquina
#  ganha a reserva, e o CODIFICADOR.zip continua do tamanho de hoje.
try:
    from rapidocr import RapidOCR
    RAPIDOCR_DISPONIVEL = True
except Exception:
    RapidOCR = None
    RAPIDOCR_DISPONIVEL = False

_rapidocr_motor = None


# ============================================================
#  CONSTANTES
# ============================================================

CNPJ_REGEX = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
NOME_ARQUIVO_PADRAO = "cadastro_condominios.xlsx"
#  Modelo de importação de despesas do Superlógica, distribuído junto do
#  executável como o cadastro. Fica AO LADO do .exe de propósito: é ali que o
#  usuário edita fornecedor, categoria e forma de pagamento quando mudarem,
#  sem precisar de programador nem de executável novo.
NOME_MODELO_DESPESAS = "modelo_despesas.xlsx"
NOME_LOG_PADRAO = "processamento.log"
NOME_CONFIG_PADRAO = "config.json"
LIMITE_TAMANHO_LOG = 5 * 1024 * 1024  # 5 MB — acima disso, rotaciona


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


TAMANHO_LOTE_PADRAO = 20  # limite de arquivos por envio no Superlógica


def caminho_do_lote(pasta_saida, indice, tamanho_lote):
    """
    Subpasta de lote em que o arquivo de índice `indice` (0-based, contando
    só os arquivos efetivamente codificados) deve ser gravado — "Lote 01",
    "Lote 02"... Existe porque o Superlógica só aceita um punhado de
    arquivos por envio, então a saída já sai dividida no tamanho certo.

    `tamanho_lote` menor que 1 significa "não separar": devolve a própria
    pasta de saída, em vez de estourar divisão por zero no meio de um
    processamento.
    """
    if tamanho_lote < 1:
        return pasta_saida
    numero = indice // tamanho_lote + 1
    return os.path.join(pasta_saida, f"Lote {numero:02d}")


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


def rotacionar_log(caminho, limite_bytes=LIMITE_TAMANHO_LOG):
    """
    Se `processamento.log` passar de `limite_bytes`, descarta as sessões mais
    antigas (do início do arquivo) até caber no limite — sessões são blocos
    separados por linha em branco, o mesmo formato que _salvar_sessao_no_log
    já grava. Se uma única sessão sozinha já passar do limite, mantém ela
    mesmo assim (nunca apaga a mais recente). Nunca lança exceção — falha ao
    rotacionar não deve travar o processamento.
    """
    try:
        if not os.path.isfile(caminho) or os.path.getsize(caminho) <= limite_bytes:
            return
        with open(caminho, "r", encoding="utf-8") as f:
            conteudo = f.read()
        sessoes = conteudo.split("\n\n")
        while len(sessoes) > 1 and len("\n\n".join(sessoes).encode("utf-8")) > limite_bytes:
            sessoes.pop(0)
        with open(caminho, "w", encoding="utf-8") as f:
            f.write("\n\n".join(sessoes))
    except Exception:
        pass


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


def paginas_para_ocr(total_paginas, max_paginas):
    """Quantas páginas o OCR deve ler. `max_paginas=None` significa todas —
    usado pelos protocolos dos Correios, onde parar na 2ª página perderia
    unidades em silêncio."""
    if max_paginas is None:
        return total_paginas
    return min(total_paginas, max(0, max_paginas))


def extrair_texto_ocr(caminho, max_paginas=2, dpi=300):
    """
    Renderiza páginas do PDF como imagem e roda OCR usando o motor nativo do
    Windows (winocr) — sem programas externos instalados. Lê no máximo
    `max_paginas` páginas; `max_paginas=None` lê o documento inteiro.
    Tenta português (pt-BR) primeiro; se não disponível, usa inglês (en-US).
    """
    if not OCR_DISPONIVEL:
        raise RuntimeError("OCR não disponível. Instale: pip install pymupdf winocr")

    textos = []
    doc = fitz.open(caminho)
    try:
        limite = paginas_para_ocr(doc.page_count, max_paginas)
        for i, pagina in enumerate(doc):
            if i >= limite:
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


def motor_de_ocr(tem_winocr=None, tem_rapidocr=None):
    """
    Qual motor usar: "winocr", "rapidocr" ou None se nenhum existir. O winocr
    sempre ganha quando está disponível — a reserva cobre a máquina onde o
    motor nativo não existe, não a leitura que deu resultado ruim.

    Os parâmetros existem para o teste; em produção ficam None e a função
    consulta as flags do módulo.
    """
    if tem_winocr is None:
        tem_winocr = OCR_DISPONIVEL
    if tem_rapidocr is None:
        tem_rapidocr = RAPIDOCR_DISPONIVEL
    if tem_winocr:
        return "winocr"
    if tem_rapidocr:
        return "rapidocr"
    return None


def extrair_texto_rapidocr(caminho, dpi=300):
    """
    OCR pelos modelos PP-OCR em ONNX. Devolve um bloco de texto por região
    detectada, separados por quebra de linha — é a estrutura real da tabela
    do protocolo, e as regexes de contagem funcionam igual.
    """
    global _rapidocr_motor
    if not RAPIDOCR_DISPONIVEL:
        raise RuntimeError("RapidOCR não disponível.")

    import numpy as np
    if _rapidocr_motor is None:
        _rapidocr_motor = RapidOCR()   # carregar os modelos é caro; reaproveita

    pedacos = []
    doc = fitz.open(caminho)
    try:
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            resultado = _rapidocr_motor(np.array(img))
            textos = getattr(resultado, "txts", None)
            if textos is None:                      # API antiga: (lista, tempo)
                textos = [linha[1] for linha in (resultado[0] or [])]
            pedacos.extend(textos or [])
    finally:
        doc.close()
    return "\n".join(pedacos)


def extrair_texto_escaneado(caminho, dpi=300):
    """
    Texto de um PDF escaneado, com o winocr na frente e o RapidOCR de reserva.
    A reserva só entra quando o winocr não existe ou quebra — uma leitura que
    funcionou nunca é substituída.
    """
    if OCR_DISPONIVEL:
        try:
            return extrair_texto_ocr(caminho, max_paginas=None, dpi=dpi)
        except Exception:
            if not RAPIDOCR_DISPONIVEL:
                raise
    if RAPIDOCR_DISPONIVEL:
        return extrair_texto_rapidocr(caminho, dpi=dpi)
    raise RuntimeError(
        "Nenhum leitor de documentos escaneados disponível. "
        "Instale: pip install pymupdf winocr")


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
#  PROTOCOLO DE RECEBIMENTO DE DOCUMENTO (Correios/Imodata)
# ============================================================

MARCADOR_PROTOCOLO_CORREIO = "Protocolo de Recebimento de Documento"

#  Quanto ruído se tolera entre o código entre parênteses e o título do
#  documento. Manuscrito lido pelo OCR cabe; um código de outro documento,
#  mais distante, não.
LIMITE_RUIDO_ANTES_DO_MARCADOR = 40
#  Janela lida antes do título: o ruído tolerado mais o espaço do próprio
#  código entre parênteses.
JANELA_CODIGO_PROTOCOLO = LIMITE_RUIDO_ANTES_DO_MARCADOR + 12
RE_MARCADOR_PROTOCOLO = re.compile(
    re.escape(MARCADOR_PROTOCOLO_CORREIO), re.IGNORECASE)
RE_CODIGO_ENTRE_PARENTESES = re.compile(r"\((\d+)\)")


def extrair_codigo_protocolo_correio(texto):
    """
    Reconhece o "Protocolo de Recebimento de Documento" (recibo de entrega
    dos Correios, formato Imodata) — sem CNPJ nenhum no documento, mas com
    o código do condomínio já pronto no texto, ex: "W700A VILLARS (10005)
    Protocolo de Recebimento de Documento...".

    Devolve o código (string) se o marcador aparecer com um código antes
    dele; None se o marcador não aparecer (documento de outro tipo — CNPJ
    continua sendo o caminho normal) ou se aparecer sem um código
    reconhecível. Não confirma se o código está cadastrado, só extrai.

    O código nem sempre vem colado no título: gente escreve o valor à caneta
    bem naquele espaço, e o OCR lê o manuscrito no meio (caso real do
    `W700A DONATELLO (10536) 7.70 Protocolo de Recebimento...`, em que o
    protocolo entrou na planilha sem código e travou a geração da planilha de
    despesas). Por isso a busca tolera até
    `LIMITE_RUIDO_ANTES_DO_MARCADOR` caracteres entre o código e o título —
    curto de propósito, para não capturar um número entre parênteses que
    esteja longe e não seja o código deste documento.
    """
    texto = texto or ""
    marcador = RE_MARCADOR_PROTOCOLO.search(texto)
    if not marcador:
        return None

    #  Lê de trás para frente: o código deste documento é o ÚLTIMO antes
    #  do título. Pegar o primeiro da janela atribuiria o condomínio
    #  errado quando o OCR emenda o fim de um documento no começo do
    #  outro — e cobrar do condomínio errado é pior que não identificar.
    inicio = max(0, marcador.start() - JANELA_CODIGO_PROTOCOLO)
    codigos = RE_CODIGO_ENTRE_PARENTESES.findall(
        texto[inicio:marcador.start()])
    return codigos[-1] if codigos else None


def montar_texto_protocolo_correio(codigo, nome, cnpj_normalizado):
    """Formata a linha única carimbada nos protocolos dos Correios:
    "10005 VILLARS - 07.945.453/0001-30"."""
    return f"{codigo} {nome} - {formatar_cnpj(cnpj_normalizado)}"


RE_LISTANDO_PROTOCOLO = re.compile(r"Listando\s+(\d+)\s+unidade", re.IGNORECASE)
#  O winocr devolve a página numa linha só, então nada de (?m)^ aqui. O
#  (?:\s*[-–—])+ cobre o traço duplicado que o OCR produz às vezes
#  ("702 - - Enny Marins de Lima", visto no protocolo real do ASTORIA).
RE_UNIDADE_PROTOCOLO = re.compile(r"\b\d{1,4}(?:\s*[-–—])+\s*[A-Za-zÀ-ÿ]")
RE_ENTREGA_PROTOCOLO = re.compile(r"\bCorreio\b", re.IGNORECASE)
RE_CABECALHO_PROTOCOLO = re.compile(
    r"([^()\n]{0,60}?)\s*\(\d+\)\s*" + re.escape(MARCADOR_PROTOCOLO_CORREIO),
    re.IGNORECASE,
)


def extrair_dados_protocolo_correio(texto):
    """
    Lê um Protocolo de Recebimento de Documento e devolve o que é preciso
    para cobrar por ele. `None` se o documento não for um protocolo.

    Três contagens independentes porque cada uma falha de um jeito: o
    "Listando N unidades" impresso é a fonte do valor, e as outras duas
    servem para confirmá-lo (ver conferir_contagem_protocolo).
    """
    texto = texto or ""
    if MARCADOR_PROTOCOLO_CORREIO.lower() not in texto.lower():
        return None

    cabecalho = RE_CABECALHO_PROTOCOLO.search(texto)
    listando = RE_LISTANDO_PROTOCOLO.search(texto)

    #  A contagem de linhas olha só o que vem ANTES do "Listando": depois
    #  dele só há rodapé (CEP, telefone) e números que o OCR inventa lendo
    #  o valor manuscrito — nada disso é unidade.
    corpo = texto[:listando.start()] if listando else texto

    return {
        "codigo": extrair_codigo_protocolo_correio(texto),
        "condominio": cabecalho.group(1).strip() if cabecalho else "",
        "total_impresso": int(listando.group(1)) if listando else None,
        "linhas_contadas": len(RE_UNIDADE_PROTOCOLO.findall(corpo)),
        "entregas_contadas": len(RE_ENTREGA_PROTOCOLO.findall(texto)),
    }


def conferir_contagem_protocolo(dados):
    """
    Decide se dá para confiar na contagem. Devolve (aceito, motivo).

    O total impresso manda; basta que UM dos dois conferidores concorde com
    ele. Sem o total impresso não se aceita nada, mesmo que os conferidores
    concordem entre si — contar linhas por OCR sozinho é chute com cara de
    precisão, e o resultado aqui vira dinheiro cobrado.
    """
    total = dados.get("total_impresso")
    if total is None:
        return False, 'Não foi possível ler o total impresso ("Listando N unidades")'

    linhas = dados.get("linhas_contadas", 0)
    entregas = dados.get("entregas_contadas", 0)
    if total == linhas or total == entregas:
        return True, ""
    return False, f"Listando {total}, mas foram contadas {linhas} e {entregas} unidades"


# ============================================================
#  VALOR DO PROTOCOLO (unidades × tarifa) E TEXTOS DO CARIMBO
# ============================================================

def valor_protocolo(unidades, tarifa):
    """
    unidades × tarifa em Decimal, duas casas. Dinheiro não passa por float:
    a tarifa vira Decimal a partir da string para não herdar o erro de
    representação binária (3.85 float não é exatamente 3,85).
    """
    tarifa_decimal = tarifa if isinstance(tarifa, Decimal) else Decimal(str(tarifa))
    bruto = Decimal(int(unidades)) * tarifa_decimal
    return bruto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def formatar_reais(valor):
    """1234.5 -> "R$ 1.234,50" (formato brasileiro)."""
    texto = f"{Decimal(str(valor)):,.2f}"
    return "R$ " + texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


LIMITE_VALOR_DIGITADO = Decimal("1000000")  # teto de sanidade para um lote


def converter_valor_digitado(texto, exemplo="300,30"):
    """
    Lê um valor em reais digitado por uma pessoa e devolve `(valor, mensagem)`:
    `(Decimal, "")` quando válido, `(None, aviso)` quando não. O aviso vai
    direto para a tela, então é frase em português, sem jargão.

    `exemplo`: número usado nos avisos genéricos ("Digite um número, como
    ...") — a tarifa de protocolo e o valor manual de um protocolo têm
    ordens de grandeza bem diferentes (3,85 vs. 300,30), então cada tela
    passa o exemplo que faz sentido pra ela, em vez de um único exemplo
    fixo que soa estranho no outro contexto.

    Aceita "300,30", "300.30", "1.234,56" e "R$ 57,75". Tolera espaço nas
    pontas e logo depois do "R$" (sobra de copiar/colar), mas espaço no meio
    dos dígitos é recusado, não ignorado: "3 85" não pode virar 385 sem
    aviso nenhum, cem vezes o valor pretendido.

    Recusa valor com centavo fracionado (ex: "3,855"), pelo mesmo motivo que
    a tarifa recusa: um valor assim produz carimbo e planilha que não fecham
    quando alguém confere no papel. A regra é sobre o valor, não sobre a
    forma como foi digitado — zero à direita não acrescenta casa nenhuma,
    então "3,850" (que é exatamente 3,85) é aceito.

    Recusa também o ponto ambíguo pelo FORMATO, antes de qualquer teste de
    arredondamento — não dá pra confiar no quantize pra pegar esse caso:
    "1.200" sem vírgula, interpretado como decimal, é `Decimal("1.200")`,
    que arredonda pra 1,20 SEM sobrar casa nenhuma (o zero à direita some no
    quantize), então passaria calado — quem digitou mil e duzentos reais
    seria cobrado um real e vinte, o pior tipo de erro porque não dá aviso
    nenhum. Por isso a regra é sobre a forma, não sobre o resultado do
    arredondamento: ponto sem vírgula seguido de exatamente três dígitos é
    separador de milhar mal digitado, não decimal — não importa se o
    resultado arredondaria "certo" ou não. "300.30" (dois dígitos depois do
    ponto) continua valendo como decimal, e qualquer valor com vírgula
    (ex: "1.234,56") nunca é ambíguo, porque a vírgula já deixa claro qual é
    o separador decimal.
    """
    bruto = (texto or "").strip()
    if bruto.startswith("R$"):
        bruto = bruto[2:].strip()
    if not bruto:
        return None, f"Digite um número, como {exemplo}."
    if " " in bruto:
        return None, f"Digite um número, como {exemplo}."

    mensagem_ambiguo = (
        'Não ficou claro se o ponto é separador de milhar ou de centavos. '
        f'Use vírgula para os centavos — ex.: "{exemplo}".'
    )

    #  Formato brasileiro: o ponto é separador de milhar e a vírgula é o
    #  decimal. Sem vírgula, o ponto é tratado como decimal ("300.30"). Sem
    #  vírgula NENHUMA, um ponto seguido de três dígitos é ambíguo — "1.200"
    #  tanto pode ser mil e duzentos reais (milhar) quanto um real e vinte
    #  (decimal) — e é recusado pelo formato, aqui, antes de qualquer
    #  conversão pra Decimal.
    tinha_virgula = "," in bruto
    if not tinha_virgula and "." in bruto:
        ultimo_grupo = bruto.rsplit(".", 1)[-1]
        if len(ultimo_grupo) == 3 and ultimo_grupo.isdigit():
            return None, mensagem_ambiguo
    if tinha_virgula:
        bruto = bruto.replace(".", "").replace(",", ".")

    try:
        valor = Decimal(bruto)
    except (InvalidOperation, ValueError):
        return None, f"Digite um número, como {exemplo}."

    if not valor.is_finite():
        return None, f"Digite um número, como {exemplo}."
    if valor <= 0:
        return None, "O valor precisa ser maior que zero."
    if valor > LIMITE_VALOR_DIGITADO:
        return None, "Esse valor parece alto demais. Confira o que foi digitado."
    if valor != valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP):
        return None, f"Use no máximo duas casas decimais, como {exemplo}."

    return valor, ""


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


def candidatos_por_nome(nome_arquivo, texto, cadastro, limite=8):
    """
    Sugere até `limite` CNPJs candidatos comparando o nome do arquivo (e o
    texto extraído do PDF, se houver) contra os nomes do cadastro — usado só
    como SUGESTÃO para escolha manual (nunca decide sozinho), diferente de
    buscar_por_nome_arquivo, que bloqueia em caso de ambiguidade.

    Chamada quando a extração de CNPJ do conteúdo não achou nenhum candidato
    (arquivo escaneado sem CNPJ legível): em vez de deixar o pendente sem
    nenhuma ação possível, oferece os nomes parecidos pro funcionário
    escolher depois de abrir o PDF — a decisão final continua sendo por
    CNPJ (do cadastro), só que escolhida por uma pessoa, não inferida
    sozinha.

    Exige que o MELHOR score atinja LIMIAR_SCORE_NOME (0.72); a partir daí,
    inclui também qualquer outro candidato a menos de LIMIAR_DIFERENCA_AMBIGUA
    (0.08) de distância do melhor — mesmo cálculo de "é ambíguo" usado em
    buscar_por_nome_arquivo, só que aqui, em vez de descartar tudo, devolve o
    grupo inteiro. `limite` é só um teto de segurança pro popup não ficar
    absurdamente longo, calibrado contra o cadastro real (~750 condomínios):
    o grupo de nomes parecidos com "CONDE DE BONFIM" (os dois verdadeiros +
    2 falsos positivos por acaso, ex: "CONDE DE VALMONT") tem 4 candidatos
    dentro da janela — `limite=8` sobra margem sem cortar nenhum.

    Devolve lista de CNPJs normalizados, do mais provável ao menos provável;
    lista vazia se o melhor candidato não atingir LIMIAR_SCORE_NOME.
    """
    alvos = []
    base = os.path.splitext(nome_arquivo)[0]
    alvo_arquivo = remover_palavras_tipo_doc(normalizar_texto_busca(base))
    if alvo_arquivo:
        alvos.append(alvo_arquivo)
    if texto:
        nome_sugerido = sugerir_nome_condominio(texto)
        if nome_sugerido:
            alvo_texto = normalizar_texto_busca(nome_sugerido)
            if alvo_texto:
                alvos.append(alvo_texto)

    if not alvos or not cadastro:
        return []

    melhor_score_por_cnpj = {}
    for cnpj_norm, dados in cadastro.items():
        nome_cad = normalizar_texto_busca(dados.get("nome", ""))
        if not nome_cad:
            continue
        melhor = max(SequenceMatcher(None, alvo, nome_cad).ratio() for alvo in alvos)
        melhor_score_por_cnpj[cnpj_norm] = melhor

    candidatos = sorted(melhor_score_por_cnpj.items(), key=lambda kv: kv[1], reverse=True)
    if not candidatos or candidatos[0][1] < LIMIAR_SCORE_NOME:
        return []

    # Inclui o melhor e qualquer outro perto o bastante dele (mesmo cálculo
    # de "ambíguo" de buscar_por_nome_arquivo) — é assim que os dois "CONDE
    # DE BONFIM" aparecem juntos, mesmo o "RES" tendo score individual abaixo
    # de LIMIAR_SCORE_NOME. `limite` é só um teto de segurança pro popup não
    # ficar absurdamente longo — contra o cadastro real (~750 condomínios),
    # o grupo de nomes parecidos com "CONDE DE BONFIM" tem 4 candidatos
    # dentro da janela, então o padrão (8) sobra margem sem cortar nenhum.
    melhor_score = candidatos[0][1]
    proximos = [
        cnpj_norm for cnpj_norm, score in candidatos
        if melhor_score - score < LIMIAR_DIFERENCA_AMBIGUA
    ]
    return proximos[:limite]


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

MARGEM_LATERAL_ROTACIONADO = 20  # pontos da borda direita, carimbo do protocolo dos Correios


def criar_overlay(largura, altura, texto, fonte, tamanho, cor, x, y, centralizado,
                  angulo=0, alinhamento="esquerda"):
    """Desenha `texto` no PDF. Se tiver quebras de linha ("\n"), cada linha é
    desenhada empilhada, a primeira em cima e as seguintes abaixo dela.

    `angulo=90` é um modo especial (protocolo dos Correios, ver
    montar_texto_protocolo_correio): ignora x/y/centralizado/alinhamento e
    desenha uma linha única rotacionada 90° (sentido anti-horário — lê de
    baixo pra cima), colada perto da borda direita e verticalmente
    centralizada.

    `alinhamento` ("esquerda", "centro", "direita") vale para o modo normal:
    "direita" faz o texto TERMINAR em x, usado pelo carimbo do valor no topo
    direito do protocolo. `centralizado=True` continua equivalendo a
    "centro", para não quebrar quem já chamava a função."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(largura, altura))
    c.setFont(fonte, tamanho)
    c.setFillColor(HexColor(cor))

    if angulo == 90:
        largura_texto = c.stringWidth(texto, fonte, tamanho)
        x_rotacionado = largura - MARGEM_LATERAL_ROTACIONADO
        y_rotacionado = (altura - largura_texto) / 2
        c.saveState()
        c.translate(x_rotacionado, y_rotacionado)
        c.rotate(90)
        c.drawString(0, 0, texto)
        c.restoreState()
    else:
        altura_linha = tamanho * 1.2
        for i, linha in enumerate(texto.split("\n")):
            x_linha = x
            if centralizado or alinhamento == "centro":
                largura_texto = c.stringWidth(linha, fonte, tamanho)
                x_linha = (largura - largura_texto) / 2
            elif alinhamento == "direita":
                largura_texto = c.stringWidth(linha, fonte, tamanho)
                x_linha = x - largura_texto
            c.drawString(x_linha, y - i * altura_linha, linha)

    c.save()
    buffer.seek(0)
    return buffer


def processar_pdf(caminho_entrada, caminho_saida, texto, config, carimbos_extras=None):
    """
    Carimba o PDF e grava a saída. `carimbos_extras` permite mais de um
    carimbo por página num único passe de escrita — usado pelo protocolo dos
    Correios, que leva o código na lateral e o valor no topo direito. Sem
    ele, o comportamento é o de sempre: um carimbo só, vindo de `config`.
    """
    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()
    for pagina in reader.pages:
        largura = float(pagina.mediabox.width)
        altura = float(pagina.mediabox.height)

        carimbos = [{
            "texto": texto,
            "fonte": config["fonte"],
            "tamanho": config["tamanho"],
            "cor": config["cor"],
            "x": config["x"],
            "y": config["y"],
            "centralizado": config["centralizado"],
            "angulo": config.get("angulo", 0),
            "alinhamento": config.get("alinhamento", "esquerda"),
        }]
        carimbos.extend(carimbos_extras or [])

        for carimbo in carimbos:
            overlay_buffer = criar_overlay(
                largura, altura, carimbo["texto"],
                carimbo["fonte"], carimbo["tamanho"], carimbo["cor"],
                carimbo["x"], carimbo["y"], carimbo["centralizado"],
                carimbo.get("angulo", 0), carimbo.get("alinhamento", "esquerda"),
            )
            pagina.merge_page(PdfReader(overlay_buffer).pages[0])

        writer.add_page(pagina)

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "wb") as f:
        writer.write(f)


# ============================================================
#  PERSISTÊNCIA DO CADASTRO (planilha xlsx)
# ============================================================

def carregar_cadastro(caminho):
    """
    Retorna dict: cnpj_normalizado -> {'codigo':..., 'nome':..., 'id_sl':...}.

    "ID SL" (4ª coluna) é o código do condomínio no Superlógica — outro
    número, sem relação com o código interno (ex: KLOSTERS é 10004 aqui e 44
    lá). Nada da identificação nem dos carimbos usa esse campo; quem usa é a
    geração da planilha de despesas do Superlógica (`lancamentos_de_despesa`). Planilha antiga de 3 colunas carrega normalmente, com o
    campo vazio.
    """
    cadastro = {}
    if not os.path.isfile(caminho):
        return cadastro
    wb = load_workbook(caminho)
    sheet = wb.active
    for linha in sheet.iter_rows(min_row=2, values_only=True):
        if not linha or not linha[0]:
            continue
        cnpj, codigo, nome, id_sl = (list(linha) + [None, None, None, None])[:4]
        cnpj_norm = normalizar_cnpj(str(cnpj))
        if cnpj_norm:
            cadastro[cnpj_norm] = {
                "codigo": str(codigo).strip() if codigo is not None else "",
                "nome": str(nome).strip() if nome is not None else "",
                "id_sl": str(id_sl).strip() if id_sl is not None else "",
            }
    return cadastro


def salvar_cadastro(caminho, cadastro):
    """
    Reescreve a planilha inteira a partir do dict. A coluna "ID SL" é gravada
    junto — sem isso, qualquer edição pela aba de Cadastro apagaria o ID SL de
    todos os condomínios de uma vez, em silêncio. Registro sem a chave (vindo
    de um formulário antigo) grava a coluna em branco.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Condominios"
    sheet.append(["CNPJ", "Código", "Nome do Condomínio", "ID SL"])
    for cnpj_norm, dados in sorted(cadastro.items(), key=lambda kv: kv[1]["nome"]):
        sheet.append([formatar_cnpj(cnpj_norm), dados["codigo"], dados["nome"],
                       dados.get("id_sl", "")])
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 12
    sheet.column_dimensions["C"].width = 35
    sheet.column_dimensions["D"].width = 10
    wb.save(caminho)


# ============================================================
#  EXTRAÇÃO DE DADOS DA NFS-e (DANFSe) PARA PLANILHA
# ============================================================
#
#  Contrapartida do carimbo: em vez de escrever o código no PDF, lê os dados
#  da nota e joga numa planilha. Vale só para NFS-e com texto nativo (DANFSe
#  da prefeitura do Rio, como as da predefinição F&F), onde cada campo vem
#  rotulado ("Rótulo\n \nValor").
#
#  Decisão importante: aqui NÃO se usa OCR, de propósito. CNPJ tem dígito
#  verificador, então um erro de leitura é detectável (é o que sustenta a
#  escada de DPI em extrair_cnpj_tomador); valor e data não têm nada disso —
#  "1.234,56" lido como "1.234,58" passaria direto para uma planilha
#  financeira sem ninguém perceber. Documento sem texto nativo é reportado
#  como não lido, nunca "chutado".

#  Títulos das seções do DANFSe. Servem para recortar o documento antes de
#  procurar um rótulo: vários rótulos se repetem (ex: "Valor do Serviço"
#  aparece em TRIBUTAÇÃO MUNICIPAL e em VALOR TOTAL DA NFS-E), e sem o
#  recorte a leitura pegaria a ocorrência da seção errada.
SECOES_DANFSE = [
    "EMITENTE DA NFS-",  # sem a letra final: "...NFS-e" na v1.0, "...NFS-E" na v2.0
    "PRESTADOR / FORNECEDOR",  # subseção só na v2.0
    "TOMADOR DO SERVI",
    "TOMADOR / ADQUIRENTE",  # nome da seção do tomador na DANFSe v2.0
    "INTERMEDI",
    "SERVIÇO PRESTADO",
    "TRIBUTAÇÃO MUNICIPAL",
    "TRIBUTAÇÃO FEDERAL",
    "TRIBUTAÇÃO IBS",  # tributo novo da reforma tributária, só na v2.0 —
                        # marcador de seção pra não vazar pro bloco vizinho;
                        # nenhum campo de dentro dela é extraído (não pedido)
    "VALOR TOTAL DA NFS",
    "TOTAIS APROXIMADOS",
    "INFORMAÇÕES COMPLEMENTARES",
]


def bloco_secao(texto, titulo):
    """Recorta o trecho do DANFSe que vai de `titulo` até o início da próxima
    seção. Sensível a maiúsculas/minúsculas de propósito: os títulos de seção
    saem sempre em CAIXA ALTA (nas DANFSe v1.0 e v2.0), mas o mesmo texto em
    Título Normal aparece como rótulo de campo dentro de outra seção (ex:
    "Código de Tributação Municipal", dentro de SERVIÇO PRESTADO) — buscar
    sem diferenciar maiúsculas pegaria esse rótulo por engano, antes da
    seção de verdade."""
    inicio = texto.find(titulo)
    if inicio == -1:
        return ""
    inicio += len(titulo)
    fim = len(texto)
    for outra in SECOES_DANFSE:
        pos = texto.find(outra, inicio)
        if pos != -1 and pos < fim:
            fim = pos
    return texto[inicio:fim]


def campo_danfse(bloco, rotulo):
    """
    Valor que vem logo abaixo de um rótulo. O DANFSe usa "Rótulo\n \nValor";
    no bloco de tributação federal vem sem a linha em branco ("Rótulo\nValor").
    Só essas duas formas são aceitas — de propósito. Com um `\\s*` solto, um
    campo vazio (ex: "Benefício Municipal", que às vezes não tem valor)
    engoliria as linhas em branco e devolveria o RÓTULO seguinte como se
    fosse o seu valor.

    Sem diferenciar maiúsculas/minúsculas — mesmo motivo do bloco_secao.
    """
    m = re.search(re.escape(rotulo) + r"[ \t]*\n[ \t]*\n?[ \t]*(.+)", bloco, re.IGNORECASE)
    if not m:
        return None
    return m.group(1).strip() or None


def converter_valor_br(texto_valor):
    """'R$ 1.234,56' -> 1234.56. Campo vazio, '-' ou não-numérico -> None."""
    if not texto_valor:
        return None
    limpo = texto_valor.replace("R$", "").strip()
    if not limpo or limpo == "-":
        return None
    limpo = limpo.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", limpo):
        return None  # não é número — provavelmente o campo estava vazio
    return float(limpo)


def converter_percentual(texto_valor):
    """'5,00 %' -> 5.0"""
    if not texto_valor:
        return None
    limpo = texto_valor.replace("%", "").strip().replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", limpo):
        return None
    return float(limpo)


def converter_data_br(texto_data):
    """'21/07/2026' -> date; '21/07/2026 20:35:22' -> datetime."""
    if not texto_data:
        return None
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})(?:\s+(\d{2}):(\d{2}):(\d{2}))?", texto_data)
    if not m:
        return None
    dia, mes, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        if m.group(4):
            return datetime.datetime(ano, mes, dia, int(m.group(4)),
                                      int(m.group(5)), int(m.group(6)))
        return datetime.date(ano, mes, dia)
    except ValueError:
        return None


def extrair_dados_nfse(texto):
    """
    Lê os campos de uma NFS-e (DANFSe) a partir do texto nativo do PDF.
    Retorna dict com os campos, ou None se o documento não for um DANFSe —
    caso do "Detalhamento do Faturamento", que vem no mesmo lote mas tem
    outro layout.
    """
    pos_emitente = texto.find("EMITENTE DA NFS-")
    cabecalho = texto[:pos_emitente] if pos_emitente != -1 else texto

    numero = campo_danfse(cabecalho, "Número da NFS-e")
    if not numero:
        return None

    # "TOMADOR / ADQUIRENTE" é o nome da seção na DANFSe v2.0 — a
    # Prefeitura renomeou (não é só maiúscula), então tenta os dois.
    bloco_tomador = bloco_secao(texto, "TOMADOR DO SERVI") or bloco_secao(texto, "TOMADOR / ADQUIRENTE")
    bloco_municipal = bloco_secao(texto, "TRIBUTAÇÃO MUNICIPAL")
    bloco_federal = bloco_secao(texto, "TRIBUTAÇÃO FEDERAL")
    bloco_total = bloco_secao(texto, "VALOR TOTAL DA NFS")

    cnpj_tomador = ""
    m_cnpj = CNPJ_REGEX.search(bloco_tomador)
    if m_cnpj:
        candidato = normalizar_cnpj(m_cnpj.group(0))
        if cnpj_valido(candidato):
            cnpj_tomador = candidato

    # "Valor da Operação / Serviço" é o rótulo na DANFSe v2.0, no lugar de
    # "Valor do Serviço" — mesma renomeação de seção, tenta os dois.
    valor_servico_bruto = (campo_danfse(bloco_total, "Valor do Serviço")
                            or campo_danfse(bloco_total, "Valor da Operação / Serviço"))

    return {
        "numero": numero,
        "competencia": converter_data_br(campo_danfse(cabecalho, "Competência da NFS-e")),
        "emissao": converter_data_br(campo_danfse(cabecalho, "Data e Hora da emissão da NFS-e")),
        "cnpj_tomador": cnpj_tomador,
        "nome_tomador": campo_danfse(bloco_tomador, "Nome / Nome Empresarial") or "",
        "valor_servico": converter_valor_br(valor_servico_bruto),
        "valor_liquido": converter_valor_br(campo_danfse(bloco_total, "Valor Líquido da NFS-e")),
        "bc_issqn": converter_valor_br(campo_danfse(bloco_municipal, "BC ISSQN")),
        "aliquota": converter_percentual(campo_danfse(bloco_municipal, "Alíquota Aplicada")),
        "issqn": converter_valor_br(campo_danfse(bloco_municipal, "ISSQN Apurado")),
        "retencao_issqn": campo_danfse(bloco_municipal, "Retenção do ISSQN") or "",
        # Retenções federais — o que de fato é descontado da nota (somadas,
        # batem com "Total das Retenções Federais" e explicam a diferença
        # entre valor do serviço e valor líquido). "Contribuições Sociais -
        # Retidas" já é o agregado de PIS+COFINS+CSLL retidos (os 4,65%), por
        # isso não se guarda PIS e COFINS separados: aqueles dois campos do
        # DANFSe são "Débito Apuração Própria", débito da própria empresa, que
        # não desconta nada da nota e confundiria a conferência do líquido.
        "previdencia_retida": converter_valor_br(
            campo_danfse(bloco_federal, "Contribuição Previdenciária - Retida")),
        "contrib_sociais_retidas": converter_valor_br(
            campo_danfse(bloco_federal, "Contribuições Sociais - Retidas")),
    }


#  (rótulo da coluna, largura, formato numérico do Excel)
COLUNAS_NFSE = [
    ("Arquivo", 38, None),
    ("Nº da NFS-e", 12, None),
    ("Competência", 13, "DD/MM/YYYY"),
    ("Data de emissão", 19, "DD/MM/YYYY HH:MM"),
    ("CNPJ do tomador", 20, None),
    ("Nome do tomador", 42, None),
    ("Código", 10, None),
    ("Valor do serviço", 16, "R$ #,##0.00"),
    ("Valor líquido", 15, "R$ #,##0.00"),
    ("BC ISSQN", 13, "R$ #,##0.00"),
    ("Alíquota (%)", 12, "0.00"),
    ("ISSQN apurado", 15, "R$ #,##0.00"),
    ("Retenção do ISSQN", 18, None),
    ("Prev. retida", 14, "R$ #,##0.00"),
    ("Contrib. sociais retidas", 24, "R$ #,##0.00"),
    ("Observação", 34, None),
]


def linha_planilha_nfse(nome_arquivo, dados, cadastro, observacao=""):
    """
    Monta a linha da planilha a partir dos dados extraídos. O código do
    condomínio vem SEMPRE do cadastro pelo CNPJ do tomador — nunca por
    semelhança de nome (ver "Cuidado: condomínios com nomes parecidos").
    Se o CNPJ não estiver cadastrado, o código sai vazio e a observação avisa.
    """
    if dados is None:
        return [nome_arquivo] + [None] * (len(COLUNAS_NFSE) - 2) + [observacao]

    registro = cadastro.get(dados["cnpj_tomador"])
    codigo = registro["codigo"] if registro else ""
    if not registro and not observacao:
        observacao = "CNPJ do tomador não está no cadastro"

    return [
        nome_arquivo,
        dados["numero"],
        dados["competencia"],
        dados["emissao"],
        formatar_cnpj(dados["cnpj_tomador"]) if dados["cnpj_tomador"] else "",
        dados["nome_tomador"],
        codigo,
        dados["valor_servico"],
        dados["valor_liquido"],
        dados["bc_issqn"],
        dados["aliquota"],
        dados["issqn"],
        dados["retencao_issqn"],
        dados["previdencia_retida"],
        dados["contrib_sociais_retidas"],
        observacao,
    ]


def salvar_planilha_nfse(caminho, linhas):
    """
    Grava a planilha de extração. Valores monetários e datas vão como
    números/datas de verdade (não texto), para poder somar e filtrar no Excel.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Notas fiscais"

    sheet.append([c[0] for c in COLUNAS_NFSE])
    for celula in sheet[1]:
        celula.font = Font(bold=True)

    for linha in linhas:
        sheet.append(linha)

    for indice, (_, largura, formato) in enumerate(COLUNAS_NFSE, start=1):
        letra = sheet.cell(row=1, column=indice).column_letter
        sheet.column_dimensions[letra].width = largura
        if formato:
            for numero_linha in range(2, sheet.max_row + 1):
                sheet.cell(row=numero_linha, column=indice).number_format = formato

    # Cabeçalho fixo + autofiltro, para conferência no Excel
    sheet.freeze_panes = "A2"
    ultima_coluna = sheet.cell(row=1, column=len(COLUNAS_NFSE)).column_letter
    sheet.auto_filter.ref = f"A1:{ultima_coluna}{sheet.max_row}"

    wb.save(caminho)


# ============================================================
#  PLANILHA DOS PROTOCOLOS DOS CORREIOS
# ============================================================

COLUNAS_PROTOCOLO = [
    ("Arquivo", 38, None),
    ("Condomínio", 30, None),
    ("Código", 10, None),
    ("Unidades", 10, "0"),
    ("Tarifa", 12, "R$ #,##0.00"),
    ("Valor", 14, "R$ #,##0.00"),
    ("Observação", 44, None),
]


def linha_planilha_protocolo(nome_arquivo, dados, cadastro, tarifa=None,
                             unidades=None, observacao="", valor_manual=None):
    """
    Monta a linha da planilha. O nome do condomínio vem do cadastro quando o
    código está lá; senão fica o que o próprio documento traz no cabeçalho.

    `unidades=None` é o caso pendente (contagem recusada) ou o de um arquivo
    que nem é protocolo: Unidades, Tarifa e Valor saem VAZIOS, nunca 0 — 0
    significaria "entregou zero unidades".

    `valor_manual`: valor em reais digitado por uma pessoa no painel de
    resultado (pendência resolvida à mão, sem contagem nem tarifa). Vale
    mesmo quando `dados` é None — "não foi possível ler o documento" é uma
    pendência de verdade (aparece pra pessoa resolver na tela), e é
    justamente o que esse valor existe pra resolver. Só "documento não é um
    protocolo" (também `dados=None`, mas sem `valor_manual`) não tem o que
    informar.
    """
    if dados is None:
        if valor_manual is None:
            return [nome_arquivo, "", "", None, None, None, observacao]
        #  Sem `dados` não há de onde tirar condomínio nem código — mas o
        #  valor digitado à mão não pode ser descartado em silêncio só
        #  porque a leitura automática falhou.
        aviso = "Valor informado manualmente"
        observacao = f"{observacao}; {aviso}" if observacao else aviso
        return [nome_arquivo, "", "", None, None, float(valor_manual), observacao]

    codigo = dados.get("codigo") or ""
    registro = None
    if codigo:
        cnpj = _codigos_do_cadastro(cadastro).get(codigo)
        registro = cadastro.get(cnpj) if cnpj else None

    condominio = registro["nome"] if registro else dados.get("condominio", "")
    if not registro:
        #  Duas causas distintas: sem código não tem o que procurar no
        #  cadastro (o documento não trouxe/o leitor não achou); com código
        #  e sem bater no cadastro, o código foi lido mas não está
        #  cadastrado. Concatena com uma observação já existente (ex.: motivo
        #  de contagem recusada) em vez de sobrescrevê-la — um protocolo pode
        #  estar pendente E sem código cadastrado ao mesmo tempo.
        motivo_codigo = ("Código não identificado no documento" if not codigo
                         else "Código não cadastrado")
        observacao = f"{observacao}; {motivo_codigo}" if observacao else motivo_codigo

    if valor_manual is not None:
        #  Valor digitado por uma pessoa no painel de resultado: Unidades e
        #  Tarifa ficam VAZIAS, porque não houve contagem nem multiplicação —
        #  0 ali significaria "entregou zero unidades". A observação é o único
        #  rastro de que o número não foi calculado pelo programa; o carimbo no
        #  PDF mostra só o valor.
        aviso = "Valor informado manualmente"
        observacao = f"{observacao}; {aviso}" if observacao else aviso
        return [nome_arquivo, condominio, codigo, None, None,
                float(valor_manual), observacao]

    if unidades is None:
        return [nome_arquivo, condominio, codigo, None, None, None, observacao]

    valor = valor_protocolo(unidades, tarifa)
    return [nome_arquivo, condominio, codigo, int(unidades),
            float(tarifa), float(valor), observacao]


def resolver_protocolo_manual(ctx, dados, valor, cadastro, carimbar):
    """
    Resolve à mão uma pendência da aba 3 (contagem dos Correios) —
    lógica de `_acao_informar_valor` extraída pra cá pra ficar testável sem
    depender de widget. Decide a pasta de saída (raiz ou a subpasta "Lote NN"
    certa, continuando de onde o processamento automático parou — mesma
    contagem `ctx["carimbados"]` que o laço automático usa, ver
    `caminho_do_lote`), chama `carimbar` pra gravar o PDF de fato e, só se
    isso der certo, avança `ctx["carimbados"]` e atualiza a linha da
    planilha em memória (`ctx["linhas"]`).

    `carimbar` é `(pasta_destino, registro_dados) -> None`, injetada pra
    manter esta função sem I/O de PDF de verdade — se levantar exceção, ela
    sobe sem `carimbados` avançar nem a linha da planilha mudar (mesma regra
    do laço automático: `carimbados` conta só quem foi de fato carimbado).

    Devolve `(linha_atualizada, motivo_painel, pasta_destino)` pro chamador
    montar o registro do painel de resultado.
    """
    registro_dados = {"codigo": dados.get("codigo"),
                      "condominio": dados.get("condominio", "")}

    pasta_destino = (
        caminho_do_lote(ctx["pasta_saida"], ctx.get("carimbados", 0),
                        ctx.get("tamanho_lote", 0))
        if ctx.get("separar_em_lotes") else ctx["pasta_saida"]
    )

    carimbar(pasta_destino, registro_dados)  # deixa exceção subir sem tocar ctx

    ctx["carimbados"] = ctx.get("carimbados", 0) + 1

    linha_atualizada = linha_planilha_protocolo(
        dados["arquivo"], registro_dados, cadastro,
        observacao=dados.get("motivo_original", ""), valor_manual=valor)
    ctx["linhas"][dados["indice_linha"]] = linha_atualizada

    motivo_painel = linha_atualizada[-1]
    return linha_atualizada, motivo_painel, pasta_destino


def salvar_planilha_protocolo(caminho, linhas):
    """
    Grava a planilha dos protocolos com linha de TOTAL no rodapé. Os totais
    são calculados aqui em Python (não como fórmula do Excel) para que o
    arquivo já chegue com o número pronto, sem depender de o Excel abrir e
    recalcular.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Protocolos"

    sheet.append([c[0] for c in COLUNAS_PROTOCOLO])
    for celula in sheet[1]:
        celula.font = Font(bold=True)

    for linha in linhas:
        sheet.append(linha)

    ultima_dados = sheet.max_row

    for indice, (_, largura, formato) in enumerate(COLUNAS_PROTOCOLO, start=1):
        letra = sheet.cell(row=1, column=indice).column_letter
        sheet.column_dimensions[letra].width = largura
        if formato:
            for numero_linha in range(2, ultima_dados + 1):
                sheet.cell(row=numero_linha, column=indice).number_format = formato

    sheet.freeze_panes = "A2"
    ultima_coluna = sheet.cell(row=1, column=len(COLUNAS_PROTOCOLO)).column_letter
    sheet.auto_filter.ref = f"A1:{ultima_coluna}{ultima_dados}"

    #  Total depois do autofiltro, para não virar uma linha filtrável
    linha_total = ultima_dados + 1
    celula_rotulo = sheet.cell(row=linha_total, column=1, value="TOTAL")
    celula_rotulo.font = Font(bold=True)
    for coluna in (4, 6):   # Unidades e Valor
        total = sum(
            linha[coluna - 1] for linha in linhas
            if isinstance(linha[coluna - 1], (int, float))
        )
        celula = sheet.cell(row=linha_total, column=coluna, value=round(total, 2))
        celula.font = Font(bold=True)
        celula.number_format = COLUNAS_PROTOCOLO[coluna - 1][2]

    wb.save(caminho)


# ============================================================
#  PLANILHA DE DESPESAS DO SUPERLÓGICA
# ============================================================

#  Nomes das duas colunas que o programa preenche, já normalizados. O resto do
#  layout (32 colunas na versão atual) pertence ao Superlógica e é copiado do
#  modelo do usuário sem interpretação.
COLUNA_DESPESA_CONDOMINIO = "condominio"
COLUNA_DESPESA_VALOR = "valor"
COLUNA_DESPESA_VENCIMENTO = "vencimento"
LINHA_MOLDE_DESPESAS = 2

#  Colunas do modelo que o Superlógica lê como data. O Excel guarda data como
#  número de série (21/08/2026 é 46255) e só o FORMATO da célula diz que
#  aquilo é data — um modelo com a célula em "General" gera uma planilha em
#  que o importador lê o número cru e grava 01/01/1970. Aconteceu de verdade,
#  e as linhas foram recusadas na importação.
COLUNAS_DATA_DESPESAS = ("vencimento", "competencia", "liquidacao")
EPOCA_EXCEL = datetime.datetime(1899, 12, 30)
SERIAL_EXCEL_MINIMO = 36526   # 2000-01-01
SERIAL_EXCEL_MAXIMO = 73050   # 2099-12-31
FORMATO_DATA_DESPESAS = "DD/MM/YYYY"


def _data_do_molde(nome_coluna, valor):
    """
    Converte para data de verdade o que o modelo trouxer numa coluna de data.
    Devolve `(valor, formato)` — `formato` é `None` quando não há o que mudar.

    Número dentro da faixa de datas plausíveis é série do Excel e vira data.
    Qualquer outra coisa levanta erro: melhor recusar do que gerar cobrança
    com data errada, que é o que acontecia antes desta checagem.
    """
    if valor is None:
        return None, None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if SERIAL_EXCEL_MINIMO <= valor <= SERIAL_EXCEL_MAXIMO:
            return (EPOCA_EXCEL + datetime.timedelta(days=float(valor)),
                    FORMATO_DATA_DESPESAS)
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor, FORMATO_DATA_DESPESAS
    raise ValueError(
        f'A coluna "{nome_coluna}" do modelo tem {valor!r}, que não é uma '
        "data. Abra o modelo e digite a data na célula (ex.: 21/08/2026) — "
        "o Superlógica recusa o lançamento quando a data não vem como data.")


def _normalizar_cabecalho(texto):
    """Cabeçalho sem acento, minúsculo e sem espaços nas pontas. Serve para
    achar a coluna pelo NOME em vez da posição: o modelo é do Superlógica e
    pode ser reordenado ou reacentuado sem aviso."""
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return texto.strip().lower()


FORMATOS_DATA_ACEITOS = ("%d/%m/%Y", "%d/%m/%y")


def converter_data_digitada(texto):
    """
    Lê uma data digitada por uma pessoa e devolve `(data, mensagem)`:
    `(datetime, "")` quando válida, `(None, aviso)` quando não. O aviso vai
    direto para a tela, então é frase em português, sem jargão.

    Aceita só o formato brasileiro (`21/08/2026`, `21-08-2026`, `21.08.2026`,
    `21/08/26`). Não aceita `2026-08-21` de propósito: misturar as duas
    convenções é como uma data tipo `03/04` acaba lançada com o mês trocado.
    """
    bruto = (texto or "").strip().replace("-", "/").replace(".", "/")
    if not bruto:
        return None, "Digite a data, como 21/08/2026."

    for formato in FORMATOS_DATA_ACEITOS:
        try:
            return datetime.datetime.strptime(bruto, formato), ""
        except ValueError:
            continue
    return None, "Data inválida. Use o formato 21/08/2026."


def gerar_planilha_despesas(caminho_modelo, caminho_saida, lancamentos,
                            vencimento=None):
    """
    Gera a planilha de importação de despesas do Superlógica a partir do
    modelo do usuário. `lancamentos` é [(id_sl, valor), ...] na ordem de saída.

    O modelo é ABERTO E PREENCHIDO, nunca reconstruído: copiar preserva
    formatos de célula, validações e colunas ocultas que o importador do
    Superlógica pode exigir e que uma planilha montada do zero perderia sem
    aviso.

    A linha 2 do modelo é o molde — os campos que se repetem em todo
    lançamento (fornecedor, categoria, forma de pagamento...). Ela é
    SUBSTITUÍDA pela primeira linha real; nenhuma linha de exemplo pode
    sobrar no arquivo final.

    `vencimento` é data do LOTE, não do modelo: quando informada, vence o que
    estiver na coluna `vencimento` e é gravada como data de verdade em todas
    as linhas. Fica fora do modelo de propósito — ela muda a cada geração, e
    era editando o modelo à mão que a data virava número e o Superlógica
    recusava os lançamentos.
    """
    if not lancamentos:
        raise ValueError(
            "Nenhum lançamento para gerar — a planilha não foi criada.")

    wb = load_workbook(caminho_modelo)
    sheet = wb.active

    colunas = {}
    for celula in sheet[1]:
        nome = _normalizar_cabecalho(celula.value)
        if nome:
            colunas.setdefault(nome, celula.column)

    obrigatorias = [COLUNA_DESPESA_CONDOMINIO, COLUNA_DESPESA_VALOR]
    if vencimento is not None:
        obrigatorias.append(COLUNA_DESPESA_VENCIMENTO)
    faltando = [nome for nome in obrigatorias if nome not in colunas]
    if faltando:
        raise ValueError(
            "O modelo não tem a(s) coluna(s): " + ", ".join(faltando) +
            ". Confira se o arquivo é o modelo de despesas do Superlógica.")

    coluna_condominio = colunas[COLUNA_DESPESA_CONDOMINIO]
    coluna_valor = colunas[COLUNA_DESPESA_VALOR]

    nomes_por_coluna = {celula.column: _normalizar_cabecalho(celula.value)
                        for celula in sheet[1]}

    #  Valor e estilo de cada célula do molde, lidos ANTES de escrever — a
    #  primeira linha gerada sobrescreve a própria linha-molde. Colunas de
    #  data que não estejam como data de verdade (`is_date`) são convertidas
    #  aqui; ver `_data_do_molde`.
    molde = []
    for coluna in range(1, sheet.max_column + 1):
        celula = sheet.cell(row=LINHA_MOLDE_DESPESAS, column=coluna)
        valor, formato = celula.value, None
        if nomes_por_coluna.get(coluna) in COLUNAS_DATA_DESPESAS and not celula.is_date:
            valor, formato = _data_do_molde(nomes_por_coluna[coluna], celula.value)
        molde.append((valor, copy.copy(celula._style), formato))

    for indice, (id_sl, valor) in enumerate(lancamentos):
        numero_linha = LINHA_MOLDE_DESPESAS + indice
        for coluna, (valor_molde, estilo, formato) in enumerate(molde, start=1):
            celula = sheet.cell(row=numero_linha, column=coluna)
            celula.value = valor_molde
            celula._style = copy.copy(estilo)
            if formato:
                celula.number_format = formato
        sheet.cell(row=numero_linha, column=coluna_condominio).value = id_sl
        sheet.cell(row=numero_linha, column=coluna_valor).value = float(valor)
        if vencimento is not None:
            celula_venc = sheet.cell(row=numero_linha,
                                     column=colunas[COLUNA_DESPESA_VENCIMENTO])
            celula_venc.value = vencimento
            celula_venc.number_format = FORMATO_DATA_DESPESAS

    #  Modelo salvo com mais de uma linha de exemplo não pode deixar resto
    #  depois do último lançamento — seria despesa fantasma na importação.
    primeira_sobra = LINHA_MOLDE_DESPESAS + len(lancamentos)
    if sheet.max_row >= primeira_sobra:
        sheet.delete_rows(primeira_sobra, sheet.max_row - primeira_sobra + 1)

    wb.save(caminho_saida)


def lancamentos_de_despesa(resultado, cadastro):
    """
    Monta os lançamentos de despesa a partir do resultado de um lote de
    protocolos. Devolve `(lancamentos, travas)`:

    - `lancamentos`: [(id_sl, valor)] na ordem do painel, pronto para
      `gerar_planilha_despesas`;
    - `travas`: {"sem_valor": [...], "sem_id_sl": [...]} com os nomes dos
      arquivos que impedem a geração — a interface não gera nada enquanto
      houver qualquer um, e mostra os dois grupos separados porque a ação é
      diferente (pendência resolve no painel; ID SL resolve no cadastro).

    Recebe o `resultado` do painel, e não as linhas da planilha, porque nas
    linhas um arquivo que não é protocolo e uma pendência "não foi possível
    ler o documento" ficam idênticos (código, condomínio e valor vazios) — e
    um deve travar enquanto o outro deve ser ignorado.

    Mesmo condomínio em dois protocolos gera dois lançamentos: cada um
    continua rastreável até o papel que o originou.
    """
    codigos = _codigos_do_cadastro(cadastro)
    lancamentos = []
    travas = {"sem_valor": [], "sem_id_sl": []}

    #  Ignorados ficam de fora sem travar: não são protocolo, nunca deveriam
    #  virar despesa.
    for item in resultado.get("pendentes", []) or []:
        travas["sem_valor"].append(item.get("arquivo", ""))

    for item in resultado.get("processados", []) or []:
        codigo = item.get("codigo") or ""
        registro = cadastro.get(codigos.get(codigo, "")) if codigo else None
        id_sl = (registro or {}).get("id_sl", "")
        if not id_sl:
            travas["sem_id_sl"].append(item.get("arquivo", ""))
            continue
        lancamentos.append((id_sl, item["valor"]))

    return lancamentos, travas


def buscar_condominios(termo, cadastro, limite=200):
    """
    Procura condomínios no cadastro por pedaço do nome ou do código, para o
    seletor do painel de resultado. Devolve lista de dicts com `codigo`,
    `nome`, `id_sl` e `cnpj`, ordenada por nome.

    Existe porque um protocolo pode chegar sem código legível (alguém escreve
    o valor à caneta por cima) ou com um código que não está no cadastro — e
    nesses casos não havia como destravar a geração da planilha de despesas
    pela tela.

    Quem está sem `id_sl` aparece na lista do mesmo jeito: escondê-lo faria a
    pessoa escolher e nada acontecer, sem entender por quê. O painel avisa.
    """
    procurado = normalizar_texto_busca(termo or "")
    digitos = re.sub(r"\D", "", termo or "")

    achados = []
    for cnpj_norm, dados in cadastro.items():
        nome = dados.get("nome", "") or ""
        codigo = str(dados.get("codigo", "") or "")
        casa_nome = procurado and procurado in normalizar_texto_busca(nome)
        casa_codigo = digitos and digitos in codigo
        if procurado or digitos:
            if not (casa_nome or casa_codigo):
                continue
        achados.append({
            "codigo": codigo,
            "nome": nome,
            "id_sl": dados.get("id_sl", "") or "",
            "cnpj": cnpj_norm,
        })

    achados.sort(key=lambda c: c["nome"])
    return achados[:limite]
