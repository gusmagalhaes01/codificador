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
from decimal import Decimal
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


from logica import (
    CNPJ_REGEX, NOME_ARQUIVO_PADRAO, NOME_LOG_PADRAO, NOME_CONFIG_PADRAO,
    NOME_ICONE, CHAVES_PERFIL, PERFIS_PADRAO, ORDEM_PREDEFINICOES,
    DEFAULTS_CONFIG, LIMITE_TEXTO_MINIMO, DPI_ESCALONAMENTO,
    NOME_MODELO_DESPESAS,
    CNPJS_INTERMEDIARIOS, LIMIAR_SCORE_NOME, LIMIAR_DIFERENCA_AMBIGUA,
    PALAVRAS_TIPO_DOC, OCR_DISPONIVEL, FITZ_DISPONIVEL, WINOCR_DISPONIVEL,
    TAMANHO_LOTE_PADRAO,
    pasta_base, caminho_recurso, nome_saida_com_codigo, caminho_do_lote,
    _gravar_erros_log, _registrar_erro_config,
    carregar_config, salvar_config, rotacionar_log,
    normalizar_cnpj, formatar_documento, extrair_texto_pdf, cnpj_valido, cpf_valido,
    extrair_texto_ocr, extrair_texto_ocr_regiao, proximo_dpi_maior,
    extrair_cnpj_tomador, sugerir_nome_condominio, extrair_codigo_protocolo_correio,
    montar_texto_protocolo_correio, montar_bloco_paybox,
    normalizar_texto_busca, remover_palavras_tipo_doc, _codigos_do_cadastro,
    buscar_por_nome_arquivo, desempatar_por_cadastro, candidatos_por_nome,
    criar_overlay, processar_pdf, carregar_cadastro, salvar_cadastro,
    extrair_dados_nfse, linha_planilha_nfse, salvar_planilha_nfse,
    extrair_dados_protocolo_correio, conferir_contagem_protocolo,
    valor_protocolo, formatar_reais,
    converter_valor_digitado,
    linha_planilha_protocolo, salvar_planilha_protocolo,
    extrair_texto_escaneado, COLUNAS_PROTOCOLO, resolver_protocolo_manual,
    registro_painel_resolvido,
    renderizar_previa_pdf, LARGURA_PREVIA,
    COL_CODIGO, COL_UNIDADES, COL_VALOR, COL_OBSERVACAO,
    COLUNAS_EDITAVEIS, COLUNAS_QUE_RECARIMBAM,
    validar_edicao_protocolo, aplicar_edicao_na_linha,
    lancamentos_de_despesa, gerar_planilha_despesas,
    buscar_condominios,
    converter_data_digitada,
)


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


#  Posição do carimbo de valor no protocolo dos Correios: canto superior
#  direito, no espaço em branco do documento. O rodapé foi descartado porque
#  protocolos de duas páginas têm conteúdo lá embaixo. A margem é fixa em
#  pontos, mas a posição final (x/y) é calculada em cima do tamanho real da
#  1ª página de cada PDF (ver _carimbar_protocolo) — páginas menores que A4
#  não teriam esse canto disponível se x/y ficassem fixos em 567/814.
MARGEM_VALOR_PROTOCOLO = 28

#  Bloco do Paybox (fornecedor / vencimento / valor) nos protocolos: canto
#  INFERIOR DIREITO, escolhido pelo usuário sobre um protocolo real. Fica
#  acima do rodapé da Imodata e à esquerda do carimbo lateral rotacionado.
#
#  A margem é medida da borda direita (e o texto é alinhado à direita) pelo
#  mesmo motivo do carimbo de valor: x fixo cairia fora da folha em páginas
#  menores que A4. O corpo 10 é o maior que faz a linha do fornecedor (a mais
#  larga, ~265pt) caber sem alcançar o carimbo lateral.
MARGEM_DIREITA_BLOCO_PAYBOX = 30
#  `Y` é a linha de CIMA do bloco; as outras três descem a partir dela. O
#  valor foi medido, não estimado: num protocolo de página cheia a tabela de
#  unidades desce até ~81 e o rodapé da Imodata começa logo abaixo, então a
#  janela livre é estreita. Varrendo as 3 páginas de um protocolo cheio
#  (10380 ROXY) e um curto (10005 VILLARS), a posição livre mais baixa ficou
#  em 74–77. Descer mais encosta no rodapé e no "1 de N"; subir encosta nas
#  linhas de assinatura.
Y_BLOCO_PAYBOX = 77
TAMANHO_BLOCO_PAYBOX = 10

#  Qualidade de leitura dos protocolos dos Correios: fixa na melhor, de
#  propósito, ignorando a que estiver escolhida na predefinição ativa.
#  Motivo: o protocolo é SEMPRE uma fotocópia escaneada e apagada (nunca tem
#  texto nativo) e o que sai daqui é dinheiro cobrado. Medido nos quatro
#  protocolos reais de referência: na "Rápida" (72) nenhum dos quatro é lido
#  — dois nem são reconhecidos como protocolo; a 150 dois ainda não fecham;
#  a 200 os quatro fecham, mas o do CARMEM passa apoiado num único
#  conferidor (contagem de linhas dá 1, só o "Correio" confirma os 3); a 300
#  os três sinais concordam nos quatro. Custa cerca de meio segundo por
#  documento, o que é imperceptível num lote.
#  A aba 1 continua usando a qualidade da predefinição — lá o CNPJ tem
#  dígito verificador e muitos boletos têm texto nativo.
QUALIDADE_LEITURA_PROTOCOLO = 300


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
        self.title("Codificador v6.17.1")
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

        # Separar a saída em subpastas "Lote 01", "Lote 02"... — o Superlógica
        # só aceita um punhado de arquivos por envio. Fica fora das
        # predefinições de lote de propósito: é preferência de quem opera
        # (depende do sistema de destino), não característica do documento.
        self.separar_em_lotes = tk.BooleanVar(
            value=bool(self.config_app.get("separar_em_lotes", False)))
        self.tamanho_lote = tk.StringVar(
            value=str(self.config_app.get("tamanho_lote", TAMANHO_LOTE_PADRAO)))

        # Variáveis - aba cadastro (formulário)
        self.form_cnpj = tk.StringVar()
        self.form_codigo = tk.StringVar()
        # Código do condomínio no Superlógica — guardado só como referência
        self.form_id_sl = tk.StringVar()
        self.form_nome = tk.StringVar()

        # Variáveis - aba de extração de notas para planilha
        self.pasta_notas = tk.StringVar()
        self.arquivo_planilha_saida = tk.StringVar()

        # Variáveis - aba de protocolos dos Correios
        self.pasta_protocolos = tk.StringVar()
        self.pasta_protocolos_saida = tk.StringVar()
        self.arquivo_planilha_protocolos = tk.StringVar()

        # Contexto do último lote de protocolos processado — alimenta o
        # painel de resultado (resolução manual, botão "Abrir planilha").
        # None até o 1º processamento da sessão; sempre lido com getattr/
        # `.get()` por segurança, nunca indexado direto.
        self._ctx_protocolos = None
        #  Painel de resultado, criado sob demanda. Nasce como None para que
        #  `_janela_para_dialogo` possa consultá-lo antes de qualquer
        #  processamento sem cair no __getattr__ do Tk.
        self._janela_resultado = None

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
        self.aba_extracao = ttk.Frame(notebook)
        self.aba_protocolos = ttk.Frame(notebook)
        self.aba_cadastro = ttk.Frame(notebook)
        self.aba_logs = ttk.Frame(notebook)
        notebook.add(self.aba_processar, text="1. Processamento")
        notebook.add(self.aba_extracao, text="2. Extrair dados")
        notebook.add(self.aba_protocolos, text="3. Protocolos dos Correios")
        notebook.add(self.aba_cadastro, text="4. Cadastro de Condomínios")
        notebook.add(self.aba_logs, text="5. Logs")

        # A ordem importa: _montar_aba_processar cria self._widgets_tema, que
        # as abas seguintes usam para registrar os widgets delas no tema.
        self._montar_aba_processar(self.aba_processar)
        self._montar_aba_extracao(self.aba_extracao)
        self._montar_aba_protocolos(self.aba_protocolos)
        self._montar_aba_cadastro(self.aba_cadastro)
        self._montar_aba_logs(self.aba_logs)

        # _montar_aba_processar já chama _estilizar_ttk() ao final, mas nesse
        # ponto as abas de Cadastro/Logs ainda não existem — chama de novo
        # agora que todas as três existem, para log_historico etc. também
        # receberem a cor do tema já na montagem inicial (não só ao alternar).
        self._estilizar_ttk()

    # --------------------------------------------------------
    #  ABA 4 — CADASTRO
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
        frame_form.columnconfigure(2, weight=0)
        frame_form.columnconfigure(3, weight=1)

        caption_in(frame_form, "CNPJ / CPF").grid(row=0, column=0, sticky="w", padx=(0, 16))
        caption_in(frame_form, "Código").grid(row=0, column=1, sticky="w", padx=(0, 16))
        caption_in(frame_form, "ID SL").grid(row=0, column=2, sticky="w", padx=(0, 16))
        caption_in(frame_form, "Nome").grid(row=0, column=3, sticky="w")

        ttk.Entry(frame_form, textvariable=self.form_cnpj, width=22, font=(fonte, 12)).grid(
            row=1, column=0, sticky="w", padx=(0, 16), pady=(2, 12))
        ttk.Entry(frame_form, textvariable=self.form_codigo, width=12, font=(fonte, 12)).grid(
            row=1, column=1, sticky="w", padx=(0, 16), pady=(2, 12))
        ttk.Entry(frame_form, textvariable=self.form_id_sl, width=10, font=(fonte, 12)).grid(
            row=1, column=2, sticky="w", padx=(0, 16), pady=(2, 12))
        ttk.Entry(frame_form, textvariable=self.form_nome, font=(fonte, 12)).grid(
            row=1, column=3, sticky="ew", pady=(2, 12))

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

        colunas = ("cnpj", "codigo", "id_sl", "nome")
        self.tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=14)
        self.tabela.heading("cnpj", text="CNPJ / CPF")
        self.tabela.heading("codigo", text="Código")
        self.tabela.heading("id_sl", text="ID SL")
        self.tabela.heading("nome", text="Nome do Condomínio")
        self.tabela.column("cnpj", width=150, anchor="w")
        self.tabela.column("codigo", width=80, anchor="w")
        self.tabela.column("id_sl", width=70, anchor="w")
        self.tabela.column("nome", width=330, anchor="w")
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
                                values=(formatar_documento(cnpj_norm), dados["codigo"],
                                        dados.get("id_sl", ""), dados["nome"]))
        if hasattr(self, "label_contagem_cadastro"):
            self.label_contagem_cadastro.configure(text=f"{len(self.cadastro)} CADASTRADOS")

    def _selecionar_linha(self, event):
        selecionado = self.tabela.selection()
        if not selecionado:
            return
        cnpj_norm = selecionado[0]
        dados = self.cadastro.get(cnpj_norm, {})
        self.form_cnpj.set(formatar_documento(cnpj_norm))
        self.form_codigo.set(dados.get("codigo", ""))
        self.form_id_sl.set(dados.get("id_sl", ""))
        self.form_nome.set(dados.get("nome", ""))

    def _limpar_form(self):
        self.form_cnpj.set("")
        self.form_codigo.set("")
        self.form_id_sl.set("")
        self.form_nome.set("")
        self.tabela.selection_remove(self.tabela.selection())

    def _adicionar_ou_atualizar(self):
        cnpj_raw = self.form_cnpj.get().strip()
        codigo = self.form_codigo.get().strip()
        id_sl = self.form_id_sl.get().strip()
        nome = self.form_nome.get().strip()

        documento = normalizar_cnpj(cnpj_raw)
        if len(documento) not in (11, 14):
            messagebox.showerror(
                "Erro",
                "Documento inválido — informe um CNPJ (14 dígitos) ou um CPF "
                "(11 dígitos).",
            )
            return
        if not codigo:
            messagebox.showerror("Erro", "Informe o código.")
            return

        #  O documento é a chave do cadastro: trocá-lo num registro que já
        #  existe precisa REMOVER a chave antiga, senão o condomínio passa a
        #  existir duas vezes. É o caso de quem estava cadastrado por CPF e
        #  depois obteve CNPJ próprio.
        selecionado = self.tabela.selection()
        if selecionado and selecionado[0] != documento:
            self.cadastro.pop(selecionado[0], None)

        self.cadastro[documento] = {"codigo": codigo, "nome": nome, "id_sl": id_sl}
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

        # --- Separar a saída em lotes (logo abaixo do botão primário) ---
        self._montar_controle_lotes(corpo, registrar).grid(
            row=5, column=0, sticky="w", pady=(0, 8))

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
        explicacao.grid(row=6, column=0, sticky="w", pady=(0, 16))

        # --- Barra de progresso (usada por _processar_em_thread) ---
        self.barra_progresso = ttk.Progressbar(corpo, mode="determinate")
        self.barra_progresso.grid(row=7, column=0, sticky="ew", pady=(0, 24))

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

    def _montar_controle_lotes(self, parent, registrar):
        """
        Checkbox "Separar a saída em lotes" + campo com a quantidade por lote,
        para colocar logo abaixo do botão primário. Usado nas duas abas que
        geram PDFs (Processamento e Protocolos) — as duas compartilham as
        mesmas tk.Var, então marcar numa vale para a outra.

        Devolve o frame; quem chama decide onde posicionar (grid/pack).
        """
        tema = self.tema_atual
        fonte = familia_fonte()

        linha = registrar(
            ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )

        chk = registrar(
            ctk.CTkCheckBox(
                linha, text="Separar a saída em lotes de", variable=self.separar_em_lotes,
                corner_radius=0, fg_color=tema["acento"], hover_color=tema["acento"],
                border_color=tema["borda_forte"], text_color=tema["texto"], font=(fonte, 13),
                command=self._salvar_preferencia_lotes,
            ),
            {"fg_color": "acento", "hover_color": "acento", "border_color": "borda_forte",
             "text_color": "texto"},
        )
        chk.pack(side="left")

        campo = registrar(
            ctk.CTkEntry(
                linha, textvariable=self.tamanho_lote, width=56, corner_radius=0,
                fg_color=tema["superficie"], border_width=1, border_color=tema["borda"],
                text_color=tema["texto"], font=(fonte, 13), justify="center",
            ),
            {"fg_color": "superficie", "border_color": "borda", "text_color": "texto"},
        )
        campo.pack(side="left", padx=8)
        # Grava ao sair do campo: digitar não dispara command em CTkEntry.
        campo.bind("<FocusOut>", lambda _e: self._salvar_preferencia_lotes())

        rotulo = registrar(
            ctk.CTkLabel(linha, text="arquivos", font=(fonte, 13),
                         text_color=tema["texto"], anchor="w"),
            {"text_color": "texto"},
        )
        rotulo.pack(side="left")

        return linha

    def _tamanho_lote_valido(self):
        """Quantidade por lote como inteiro ≥ 1. Campo em branco ou com texto
        inválido cai no padrão — nunca deixa um valor ruim chegar até o
        processamento (0 dividiria por zero, negativo não faz sentido)."""
        try:
            valor = int(self.tamanho_lote.get().strip())
        except (ValueError, AttributeError):
            return TAMANHO_LOTE_PADRAO
        return valor if valor >= 1 else TAMANHO_LOTE_PADRAO

    def _salvar_preferencia_lotes(self):
        """Persiste a preferência de separar em lotes no config.json. Fica no
        nível raiz, não dentro da predefinição: depende do sistema pra onde os
        arquivos vão (Superlógica), não do tipo de documento do lote."""
        self.config_app["separar_em_lotes"] = bool(self.separar_em_lotes.get())
        self.config_app["tamanho_lote"] = self._tamanho_lote_valido()
        salvar_config(self.config_app)

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

    def _montar_painel_resultado(self, titulo, geometria, minimo):
        """
        Cria (ou reaproveita, limpando o conteúdo) a janela de um painel de
        resultado. Compartilhado pelas abas 1 e 3 — as duas têm o mesmo
        formato de janela, só muda o que vai dentro.
        """
        existente = getattr(self, "_janela_resultado", None)
        reaproveitada = existente is not None and existente.winfo_exists()
        if reaproveitada:
            for widget in existente.winfo_children():
                widget.destroy()
            janela = existente
            janela.focus_force()
        else:
            janela = ctk.CTkToplevel(self)
            self._janela_resultado = janela

        janela.title(titulo)
        #  `geometry` só na primeira vez: reaplicá-la ao remontar jogava a
        #  janela de volta ao tamanho padrão, desfazendo o redimensionamento
        #  que a pessoa tinha feito para ver a tabela inteira.
        if not reaproveitada:
            janela.geometry(geometria)
        janela.minsize(*minimo)
        janela.resizable(True, True)
        janela.configure(fg_color=self.tema_atual["fundo"])
        janela.transient(self)
        return janela

    def _montar_faixa_cartoes(self, parent, cartoes, fonte=None):
        """
        Linha de caixinhas de resumo no topo do painel: rótulo pequeno em cima,
        número grande embaixo, separados por hairlines verticais. `cartoes` é
        uma lista de (valor, rotulo, destaque); o destaque sai em cobalto,
        reservado ao que exige ação do usuário. `fonte` é opcional — quem
        chama pode passar a `familia_fonte()` já calculada (evita consultar
        de novo as fontes do sistema); se omitido, o helper calcula sozinho,
        pra continuar funcionando também fora de `mostrar_resultado`.
        """
        tema = self.tema_atual
        if fonte is None:
            fonte = familia_fonte()

        faixa = ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"])
        faixa.pack(fill="x", padx=24, pady=(24, 0))
        #  Os rótulos de valor ficam acessíveis em `faixa.labels_valor` (na
        #  ordem de `cartoes`) para que um painel possa atualizar o número sem
        #  remontar a faixa inteira. Ver `_atualizar_painel_protocolos`.
        faixa.labels_valor = []

        for indice, (valor, rotulo, destaque) in enumerate(cartoes):
            coluna = indice * 2
            faixa.columnconfigure(coluna, weight=1)

            cor_rotulo = tema["acento"] if destaque else tema["texto_secundario"]
            cor_valor = tema["acento"] if destaque else tema["texto"]

            bloco = ctk.CTkFrame(faixa, corner_radius=0, fg_color=tema["fundo"])
            bloco.grid(row=0, column=coluna, sticky="nsew", padx=16)
            ctk.CTkLabel(bloco, text=rotulo, font=(fonte, 11),
                         text_color=cor_rotulo, anchor="w").pack(fill="x", anchor="w")
            label_valor = ctk.CTkLabel(bloco, text=valor, font=(fonte, 40),
                                        text_color=cor_valor, anchor="w")
            label_valor.pack(fill="x", anchor="w")
            faixa.labels_valor.append(label_valor)

            if indice < len(cartoes) - 1:
                linha = ctk.CTkFrame(faixa, width=1, corner_radius=0,
                                     fg_color=tema["borda"])
                linha.grid(row=0, column=coluna + 1, sticky="ns")

        return faixa

    def _montar_tabela_resultado(self, parent, colunas, larguras, altura=6,
                                  ancoras=None, pady=(0, 8)):
        """
        Treeview estilizada com a aparência do projeto, com barra de rolagem.
        `colunas` é uma lista de (chave, titulo); `larguras` casa por posição.
        `ancoras`, se informado, também casa por posição com `colunas` e
        define o alinhamento de cada coluna (padrão "w" em todas quando
        omitido). `pady` controla o espaçamento vertical do frame que envolve
        a tabela (padrão (0, 8), igual ao da tabela de pendentes). Devolve a
        tabela; quem chama insere as linhas.
        """
        frame = ctk.CTkFrame(parent, corner_radius=0, fg_color=self.tema_atual["fundo"])
        frame.pack(fill="both", expand=False, pady=pady)

        chaves = [c[0] for c in colunas]
        if ancoras is None:
            ancoras = ["w"] * len(colunas)
        tabela = ttk.Treeview(frame, columns=chaves, show="headings", height=altura)
        for (chave, titulo), largura, ancora in zip(colunas, larguras, ancoras):
            tabela.heading(chave, text=titulo)
            tabela.column(chave, width=largura, anchor=ancora)
        tabela.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
        tabela.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")
        return tabela

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

        janela = self._montar_painel_resultado(
            "Resultado do processamento", "820x680", (640, 480))

        # dados de cada pendente, indexados pelo iid da linha na tabela —
        # usado pelos botões de ação e pelo duplo clique para recuperar o
        # dict original (cnpj, nome_sugerido, candidatos, caminho, tipo...)
        self._pend_por_iid = {}

        # ===================== FAIXA DE CARTÕES =====================
        self._montar_faixa_cartoes(janela, [
            (str(len(processados)), "PROCESSADOS", False),
            (str(len(pendentes)), "PENDENTES", True),
            (tempo_str, "TEMPO", False),
        ], fonte=fonte)

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
            tabela_pend = self._montar_tabela_resultado(
                area, [("arquivo", "Arquivo"), ("motivo", "Motivo")], [280, 420])

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
            tabela_proc = self._montar_tabela_resultado(
                area,
                [("arquivo", "Arquivo"), ("codigo", "Código"), ("origem", "Origem")],
                [280, 90, 280], altura=8,
                ancoras=["w", "center", "w"], pady=(0, 16))

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
        self.form_cnpj.set(formatar_documento(cnpj) if cnpj else "")
        self.form_codigo.set("")
        # limpa: senão herdaria o ID SL da linha que estivesse selecionada
        self.form_id_sl.set("")
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
            texto = f"{formatar_documento(cnpj_norm)} — {nome}"
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
    # --------------------------------------------------------
    #  ABA 2 — EXTRAIR DADOS DAS NOTAS PARA PLANILHA
    # --------------------------------------------------------
    def _montar_aba_extracao(self, parent_externo):
        """
        Contrapartida do carimbo: em vez de escrever o código no PDF, lê os
        dados das NFS-e da pasta e gera uma planilha. Mesma linguagem visual
        da tela Principal (Swiss, cantos retos, cobalto só no botão primário).
        """
        tema = self.tema_atual
        fonte = familia_fonte()

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
        bloco_titulo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_titulo.pack(fill="x", padx=24, pady=(24, 0))

        titulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="Extrair dados", font=(fonte, 20),
                         text_color=tema["texto"], anchor="w"),
            {"text_color": "texto"},
        )
        titulo.pack(anchor="w")

        subtitulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="NOTAS FISCAIS PARA PLANILHA", font=(fonte, 11),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        subtitulo.pack(anchor="w")

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

        def montar_campo(linha_grid, rotulo, variavel, comando_trocar):
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
                             fg_color=tema["superficie"], border_width=1,
                             border_color=tema["borda"], text_color=tema["texto"],
                             font=(fonte, 13)),
                {"fg_color": "superficie", "border_color": "borda", "text_color": "texto"},
            )
            entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

            botao = registrar(
                ctk.CTkButton(
                    linha, text="Trocar", corner_radius=0, width=90,
                    fg_color="transparent", hover_color=tema["superficie"],
                    border_width=1, border_color=tema["borda_forte"],
                    text_color=tema["texto"], font=(fonte, 13), command=comando_trocar,
                ),
                {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
            )
            botao.grid(row=0, column=1)

        montar_campo(0, "PASTA COM AS NOTAS", self.pasta_notas, self._escolher_pasta_notas)
        montar_campo(2, "SALVAR PLANILHA EM", self.arquivo_planilha_saida,
                     self._escolher_planilha_saida)

        # --- Botão primário (mesmo papel do "Processar" na tela Principal) ---
        self.botao_extrair = registrar(
            ctk.CTkButton(
                corpo, text="Extrair dados", corner_radius=0, height=44, font=(fonte, 15),
                fg_color=tema["acento"], hover_color=tema["acento_hover"],
                text_color=tema["sobre_acento"], border_width=0,
                command=self._iniciar_extracao,
            ),
            {"fg_color": "acento", "hover_color": "acento_hover", "text_color": "sobre_acento"},
        )
        self.botao_extrair.grid(row=4, column=0, sticky="ew", pady=(8, 8))

        explicacao = registrar(
            ctk.CTkLabel(
                corpo,
                text=("Lê as notas fiscais da pasta e gera uma planilha com os dados de cada "
                      "uma. Não altera os PDFs. Notas escaneadas não são lidas — só as que "
                      "têm texto."),
                font=(fonte, 13), text_color=tema["texto_terciario"],
                justify="left", anchor="w", wraplength=640,
            ),
            {"text_color": "texto_terciario"},
        )
        explicacao.grid(row=5, column=0, sticky="w", pady=(0, 16))

        self.barra_extracao = ttk.Progressbar(corpo, mode="determinate")
        self.barra_extracao.grid(row=6, column=0, sticky="ew", pady=(0, 8))

        self.label_status_extracao = registrar(
            ctk.CTkLabel(corpo, text="", font=(fonte, 13),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        self.label_status_extracao.grid(row=7, column=0, sticky="w", pady=(0, 24))

        self._atualizar_botao_extrair()

    def _escolher_pasta_notas(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com as notas fiscais")
        if pasta:
            self.pasta_notas.set(pasta)
            # Sugere um destino junto da pasta escolhida, para o usuário não
            # precisar decidir nada quando só quer a planilha rapidamente.
            nome_pasta = os.path.basename(os.path.normpath(pasta)) or "notas"
            self.arquivo_planilha_saida.set(
                os.path.join(pasta, f"dados_notas_{nome_pasta}.xlsx"))
        self._atualizar_botao_extrair()

    def _escolher_planilha_saida(self):
        atual = self.arquivo_planilha_saida.get().strip()
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=os.path.basename(atual) if atual else "dados_notas.xlsx",
            initialdir=os.path.dirname(atual) if atual else None,
        )
        if caminho:
            self.arquivo_planilha_saida.set(caminho)
        self._atualizar_botao_extrair()

    def _atualizar_botao_extrair(self):
        """Conta os *.pdf da pasta escolhida e ajusta texto/estado do botão."""
        pasta = self.pasta_notas.get().strip()
        quantidade = 0
        if pasta and os.path.isdir(pasta):
            try:
                quantidade = sum(1 for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
            except Exception:
                quantidade = 0

        if quantidade > 0 and self.arquivo_planilha_saida.get().strip():
            plural = "nota" if quantidade == 1 else "notas"
            self.botao_extrair.configure(
                text=f"Extrair dados de {quantidade} {plural}", state="normal")
        else:
            self.botao_extrair.configure(text="Extrair dados", state="disabled")

    def _iniciar_extracao(self):
        pasta = self.pasta_notas.get().strip()
        destino = self.arquivo_planilha_saida.get().strip()

        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione a pasta com as notas fiscais.")
            return
        if not destino:
            messagebox.showerror("Erro", "Escolha onde salvar a planilha.")
            return
        if os.path.isfile(destino) and not messagebox.askyesno(
            "Substituir planilha",
            f"Este arquivo já existe e será substituído:\n\n{destino}\n\nContinuar?"
        ):
            return

        self.botao_extrair.configure(state="disabled")
        self.label_status_extracao.configure(text="Lendo as notas...")

        thread = threading.Thread(
            target=self._extrair_em_thread, args=(pasta, destino), daemon=True)
        thread.start()

    def _extrair_em_thread(self, pasta, destino):
        """
        Lê cada PDF da pasta escolhida (sem entrar em subpastas) e monta as
        linhas da planilha. Sem OCR de propósito — ver a nota em logica.py,
        seção de extração: valor e data não têm dígito verificador, então uma
        leitura errada entraria na planilha sem ninguém perceber.
        """
        arquivos = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
        total = len(arquivos)
        self.after(0, lambda: self.barra_extracao.configure(maximum=total, value=0))

        linhas = []
        extraidas = 0
        ignoradas = 0
        sem_cadastro = 0

        for indice, nome in enumerate(arquivos, 1):
            caminho = os.path.join(pasta, nome)
            self.after(0, lambda i=indice, n=nome:
                       self.label_status_extracao.configure(text=f"Lendo {i} de {total}: {n}"))

            dados = None
            observacao = ""
            try:
                texto = extrair_texto_pdf(caminho)
                if len(texto.strip()) < LIMITE_TEXTO_MINIMO:
                    observacao = "Nota escaneada — não foi possível ler"
                else:
                    dados = extrair_dados_nfse(texto)
                    if dados is None:
                        observacao = "Não é uma nota fiscal (outro tipo de documento)"
            except Exception as e:
                observacao = f"Erro ao ler: {e}"

            linha = linha_planilha_nfse(nome, dados, self.cadastro, observacao)
            linhas.append(linha)

            if dados is None:
                ignoradas += 1
            else:
                extraidas += 1
                if not linha[6]:           # coluna "Código" vazia
                    sem_cadastro += 1

            self.after(0, lambda v=indice: self.barra_extracao.configure(value=v))

        try:
            salvar_planilha_nfse(destino, linhas)
        except Exception as e:
            #  `erro=e` como argumento padrão: o `except ... as e` desvincula
            #  `e` ao sair do bloco, e o `self.after` só executa a lambda bem
            #  depois — sem o argumento padrão, ela veria "cannot access free
            #  variable 'e'" em vez da mensagem amigável.
            self.after(0, lambda: self.label_status_extracao.configure(
                text="Não foi possível salvar a planilha."))
            self.after(0, self._atualizar_botao_extrair)
            self.after(0, lambda erro=e: messagebox.showerror(
                "Erro ao salvar",
                f"Não foi possível gravar a planilha:\n{erro}\n\n"
                "Se ela estiver aberta no Excel, feche e tente de novo."))
            return

        resumo = f"{extraidas} nota(s) extraída(s)"
        if sem_cadastro:
            resumo += f" · {sem_cadastro} com CNPJ fora do cadastro"
        if ignoradas:
            resumo += f" · {ignoradas} arquivo(s) ignorado(s)"

        self.after(0, lambda: self.label_status_extracao.configure(text=resumo))
        self.after(0, self._atualizar_botao_extrair)
        self.after(0, lambda: self._concluir_extracao(destino, resumo))

    def _concluir_extracao(self, destino, resumo):
        if messagebox.askyesno(
            "Planilha pronta",
            f"{resumo}.\n\nSalva em:\n{destino}\n\nAbrir a planilha agora?"
        ):
            try:
                os.startfile(destino)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir a planilha:\n{e}")

    # --------------------------------------------------------
    #  ABA 3 — PROTOCOLOS DOS CORREIOS
    # --------------------------------------------------------
    def _montar_aba_protocolos(self, parent_externo):
        """
        Conta as unidades de cada protocolo, calcula o valor pela tarifa do
        lote, carimba o PDF e gera a planilha. Mesma linguagem visual das
        outras abas (Swiss, cantos retos, cobalto só no botão primário).
        """
        tema = self.tema_atual
        fonte = familia_fonte()

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

        bloco_titulo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        bloco_titulo.pack(fill="x", padx=24, pady=(24, 0))

        titulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="Protocolos dos Correios", font=(fonte, 20),
                         text_color=tema["texto"], anchor="w"),
            {"text_color": "texto"},
        )
        titulo.pack(anchor="w")

        subtitulo = registrar(
            ctk.CTkLabel(bloco_titulo, text="CONTAGEM DE UNIDADES E VALOR", font=(fonte, 11),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        subtitulo.pack(anchor="w")

        hairline = registrar(
            ctk.CTkFrame(container, height=1, corner_radius=0, fg_color=tema["borda"]),
            {"fg_color": "borda"},
        )
        hairline.pack(fill="x", padx=24, pady=(16, 24))

        corpo = registrar(
            ctk.CTkFrame(container, corner_radius=0, fg_color=tema["fundo"]),
            {"fg_color": "fundo"},
        )
        corpo.pack(fill="both", expand=True, padx=24)
        corpo.columnconfigure(0, weight=1)

        def montar_campo(linha_grid, rotulo, variavel, comando_trocar):
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
                             fg_color=tema["superficie"], border_width=1,
                             border_color=tema["borda"], text_color=tema["texto"],
                             font=(fonte, 13)),
                {"fg_color": "superficie", "border_color": "borda", "text_color": "texto"},
            )
            entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

            botao = registrar(
                ctk.CTkButton(
                    linha, text="Trocar", corner_radius=0, width=90,
                    fg_color="transparent", hover_color=tema["superficie"],
                    border_width=1, border_color=tema["borda_forte"],
                    text_color=tema["texto"], font=(fonte, 13), command=comando_trocar,
                ),
                {"hover_color": "superficie", "border_color": "borda_forte", "text_color": "texto"},
            )
            botao.grid(row=0, column=1)

        montar_campo(0, "PASTA COM OS PROTOCOLOS", self.pasta_protocolos,
                     self._escolher_pasta_protocolos)
        montar_campo(2, "SALVAR OS PDFS CARIMBADOS EM", self.pasta_protocolos_saida,
                     self._escolher_pasta_protocolos_saida)
        montar_campo(4, "SALVAR PLANILHA EM", self.arquivo_planilha_protocolos,
                     self._escolher_planilha_protocolos)

        self.botao_protocolos = registrar(
            ctk.CTkButton(
                corpo, text="Calcular protocolos", corner_radius=0, height=44,
                font=(fonte, 15), fg_color=tema["acento"],
                hover_color=tema["acento_hover"], text_color=tema["sobre_acento"],
                border_width=0, command=self._iniciar_protocolos,
            ),
            {"fg_color": "acento", "hover_color": "acento_hover", "text_color": "sobre_acento"},
        )
        self.botao_protocolos.grid(row=6, column=0, sticky="ew", pady=(8, 8))

        # --- Separar a saída em lotes (mesmas tk.Var da aba 1) ---
        self._montar_controle_lotes(corpo, registrar).grid(
            row=7, column=0, sticky="w", pady=(0, 8))

        explicacao = registrar(
            ctk.CTkLabel(
                corpo,
                text=("Conta quantas unidades cada protocolo entregou, multiplica pelo "
                      "valor por linha que você informar e escreve o total no canto "
                      "superior direito do PDF, junto do código do condomínio. "
                      "Protocolos em que a contagem não confere ficam sem carimbo e "
                      "aparecem na planilha com o motivo."),
                font=(fonte, 13), text_color=tema["texto_terciario"],
                justify="left", anchor="w", wraplength=640,
            ),
            {"text_color": "texto_terciario"},
        )
        explicacao.grid(row=8, column=0, sticky="w", pady=(0, 16))

        self.barra_protocolos = ttk.Progressbar(corpo, mode="determinate")
        self.barra_protocolos.grid(row=9, column=0, sticky="ew", pady=(0, 8))

        self.label_status_protocolos = registrar(
            ctk.CTkLabel(corpo, text="", font=(fonte, 13),
                         text_color=tema["texto_secundario"], anchor="w"),
            {"text_color": "texto_secundario"},
        )
        self.label_status_protocolos.grid(row=10, column=0, sticky="w", pady=(0, 24))

        self._atualizar_botao_protocolos()

    def _escolher_pasta_protocolos(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os protocolos")
        if pasta:
            self.pasta_protocolos.set(pasta)
            nome_pasta = os.path.basename(os.path.normpath(pasta)) or "protocolos"
            self.pasta_protocolos_saida.set(os.path.join(pasta, "carimbados"))
            self.arquivo_planilha_protocolos.set(
                os.path.join(pasta, f"protocolos_{nome_pasta}.xlsx"))
        self._atualizar_botao_protocolos()

    def _escolher_pasta_protocolos_saida(self):
        pasta = filedialog.askdirectory(title="Onde salvar os PDFs carimbados")
        if pasta:
            self.pasta_protocolos_saida.set(pasta)
        self._atualizar_botao_protocolos()

    def _escolher_planilha_protocolos(self):
        atual = self.arquivo_planilha_protocolos.get().strip()
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=os.path.basename(atual) if atual else "protocolos.xlsx",
            initialdir=os.path.dirname(atual) if atual else None,
        )
        if caminho:
            self.arquivo_planilha_protocolos.set(caminho)
        self._atualizar_botao_protocolos()

    def _atualizar_botao_protocolos(self):
        """Conta os *.pdf da pasta escolhida e ajusta texto/estado do botão."""
        pasta = self.pasta_protocolos.get().strip()
        quantidade = 0
        if pasta and os.path.isdir(pasta):
            try:
                quantidade = sum(1 for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
            except Exception:
                quantidade = 0

        pronto = (quantidade > 0
                  and self.pasta_protocolos_saida.get().strip()
                  and self.arquivo_planilha_protocolos.get().strip())
        if pronto:
            plural = "protocolo" if quantidade == 1 else "protocolos"
            self.botao_protocolos.configure(
                text=f"Calcular {quantidade} {plural}", state="normal")
        else:
            self.botao_protocolos.configure(text="Calcular protocolos", state="disabled")

    def _pedir_tarifa(self):
        """
        Pede o valor por linha do lote. Campo vazio de propósito, sem valor
        padrão: a tarifa muda com o tempo (nos protocolos de referência
        aparecem 3,45 e 3,85) e um padrão herdado passaria batido.
        Devolve Decimal, ou None se o usuário cancelar.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        janela = ctk.CTkToplevel(self)
        janela.title("Valor por linha")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(self)
        janela.grab_set()

        resultado = {"valor": None}

        ctk.CTkLabel(janela, text="Quanto custa cada linha entregue?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text="Ex.: 3,85", font=(fonte, 12),
                     text_color=tema["texto_terciario"]).pack(padx=24, anchor="w")

        entrada = ctk.CTkEntry(janela, corner_radius=0, width=200,
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(padx=24, pady=(12, 4), anchor="w")
        entrada.focus_set()

        aviso = ctk.CTkLabel(janela, text="", font=(fonte, 12),
                             text_color=tema["acento"])
        aviso.pack(padx=24, pady=(0, 8), anchor="w")

        def confirmar():
            #  Exemplo "3,85": a tarifa é por linha entregue, valor pequeno —
            #  o mesmo exemplo do valor manual (300,30) pareceria sugerir uma
            #  ordem de grandeza errada aqui.
            valor, mensagem = converter_valor_digitado(entrada.get(), exemplo="3,85")
            if valor is None:
                aviso.configure(text=mensagem)
                return
            resultado["valor"] = valor
            janela.destroy()

        linha_botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        linha_botoes.pack(padx=24, pady=(0, 24), anchor="e")

        ctk.CTkButton(linha_botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(linha_botoes, text="Continuar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return resultado["valor"]

    def montar_config_atual(self):
        """
        Monta o dict de configuração do carimbo (fonte, tamanho, cor e
        posição) a partir dos campos da aba 1 — a mesma
        predefinição em uso na sessão. Usado por telas que carimbam PDFs fora
        do fluxo de identificação por CNPJ, como a aba de protocolos dos
        Correios (posição/cor do texto principal seguem o que já está
        configurado; o carimbo lateral e o do valor da aba 3 sobrepõem x/y
        conforme o próprio layout do protocolo pede).

        Só aparência: qualidade de leitura NÃO entra aqui. A aba 3 lê sempre
        na melhor qualidade (QUALIDADE_LEITURA_PROTOCOLO) e um valor de
        leitura dentro de um dict de carimbo misturava duas coisas.
        """
        try:
            tamanho = int(self.tamanho_fonte.get())
        except ValueError:
            perfil_ativo = self.predefinicao_ativa.get()
            tamanho = int(DEFAULTS_CONFIG["predefinicoes"][perfil_ativo]["tamanho_fonte"])

        cor = self.cor_texto.get().strip() or "#000000"
        modo = self.modo_texto.get()
        if modo == "rodape":
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 0, "y": 90, "centralizado": True}
        else:
            config = {"fonte": "Helvetica-Bold", "tamanho": tamanho, "cor": cor,
                      "x": 120, "y": 815, "centralizado": False}
        return config

    def _iniciar_protocolos(self):
        pasta = self.pasta_protocolos.get().strip()
        pasta_saida = self.pasta_protocolos_saida.get().strip()
        destino = self.arquivo_planilha_protocolos.get().strip()

        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione a pasta com os protocolos.")
            return
        if not pasta_saida:
            messagebox.showerror("Erro", "Escolha onde salvar os PDFs carimbados.")
            return
        if os.path.normcase(os.path.abspath(os.path.normpath(pasta_saida))) == \
                os.path.normcase(os.path.abspath(os.path.normpath(pasta))):
            messagebox.showerror(
                "Erro", "A pasta de saída precisa ser diferente da pasta de origem, "
                        "para não sobrescrever os protocolos originais.")
            return
        if not destino:
            messagebox.showerror("Erro", "Escolha onde salvar a planilha.")
            return
        if os.path.isfile(destino) and not messagebox.askyesno(
            "Substituir planilha",
            f"Este arquivo já existe e será substituído:\n\n{destino}\n\nContinuar?"
        ):
            return

        tarifa = self._pedir_tarifa()
        if tarifa is None:
            return

        #  O vencimento é perguntado AQUI, junto da tarifa, e não mais só na
        #  hora de gerar a planilha: ele vai carimbado no PDF (é um dos três
        #  campos que o Paybox usa para anexar o documento à despesa sozinho)
        #  e o carimbo acontece agora. Perguntar uma vez só também garante que
        #  o vencimento do papel e o da planilha nunca divirjam — se
        #  divergissem, o Paybox não acharia o lançamento e ninguém saberia.
        vencimento = self._pedir_vencimento_despesas()
        if vencimento is None:
            return

        #  A chave também é dado do lote e muda a cada importação. Perguntar
        #  aqui evita que alguém tenha de abrir a planilha gerada no Excel só
        #  para trocá-la — foi assim que o vencimento perdeu o formato de data
        #  e o Superlógica gravou 01/01/1970 sem acusar erro.
        chave = self._pedir_chave_despesas()
        if chave is None:
            return

        self.botao_protocolos.configure(state="disabled")
        self.label_status_protocolos.configure(text="Lendo os protocolos...")

        thread = threading.Thread(
            target=self._processar_protocolos_em_thread,
            args=(pasta, pasta_saida, destino, tarifa, vencimento, chave),
            daemon=True)
        thread.start()

    def _ler_texto_protocolo(self, caminho, dpi):
        """Texto nativo se houver; senão o documento inteiro pelo leitor de
        escaneados (winocr, com RapidOCR de reserva se estiver instalado).
        Todas as páginas — parar na 2ª perderia unidades.

        Devolve (texto, usou_leitor_escaneado): o segundo valor diz se foi
        preciso recorrer ao leitor de escaneados, para o chamador só tentar
        de novo em qualidade maior quando a leitura já não era confiável
        (texto nativo não deveria virar uma segunda tentativa por OCR)."""
        texto = extrair_texto_pdf(caminho)
        if len(texto.strip()) >= LIMITE_TEXTO_MINIMO:
            return texto, False
        return extrair_texto_escaneado(caminho, dpi=dpi), True

    def _processar_protocolos_em_thread(self, pasta, pasta_saida, destino, tarifa,
                                         vencimento, chave):
        arquivos = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
        total = len(arquivos)
        self.after(0, lambda: self.barra_protocolos.configure(maximum=total, value=0))

        config = self.montar_config_atual()
        codigos = _codigos_do_cadastro(self.cadastro)
        separar_em_lotes = self.separar_em_lotes.get()
        tamanho_lote = self._tamanho_lote_valido()

        linhas = []
        carimbados = 0
        pendentes = 0
        ignorados = 0
        falhas_carimbo = 0
        processados_painel = []
        pendentes_painel = []
        ignorados_painel = []

        for indice, nome in enumerate(arquivos, 1):
            caminho = os.path.join(pasta, nome)
            self.after(0, lambda i=indice, n=nome:
                       self.label_status_protocolos.configure(
                           text=f"Lendo {i} de {total}: {n}"))

            dados = None
            unidades = None
            observacao = ""
            #  Só fica True quando dá pra concluir, com segurança, que o
            #  arquivo de fato NÃO é um protocolo dos Correios (texto nativo,
            #  sem o marcador) — é o único caso que deve contar como
            #  "ignorado" no resumo. "Não foi possível ler" e erro de leitura
            #  ficam False aqui e caem no bucket de pendência lá embaixo: são
            #  cobranças que ficaram de fora do lote, não documentos alheios.
            nao_e_protocolo = False
            try:
                dpi = QUALIDADE_LEITURA_PROTOCOLO
                texto, usou_leitor_escaneado = self._ler_texto_protocolo(caminho, dpi)
                dados = extrair_dados_protocolo_correio(texto)

                #  Teto de UMA re-tentativa por arquivo, some com a que já
                #  existia para a contagem divergente — nunca duas re-leituras
                #  no mesmo arquivo. `ja_tentou_de_novo` é o que garante isso:
                #  se o marcador do protocolo já não foi achado na 1ª leitura
                #  e a 2ª (qualidade maior) também não achar, a checagem de
                #  contagem abaixo não tenta uma 3ª leitura.
                ja_tentou_de_novo = False
                if dados is None and usou_leitor_escaneado:
                    #  Sem o marcador do protocolo não dá para saber se é
                    #  porque o documento não é um protocolo dos Correios ou
                    #  porque a leitura na qualidade "Rápida" simplesmente não
                    #  pegou o texto — só texto nativo (não veio do leitor de
                    #  escaneados) sustenta a conclusão "não é um protocolo"
                    #  sem re-tentar.
                    dpi_maior = proximo_dpi_maior(dpi)
                    if dpi_maior and dpi_maior != dpi:
                        ja_tentou_de_novo = True
                        try:
                            texto_maior = extrair_texto_escaneado(caminho, dpi=dpi_maior)
                            dados = extrair_dados_protocolo_correio(texto_maior)
                        except Exception:
                            #  A re-tentativa falhou — dados continua None,
                            #  e a observação abaixo reflete "não foi possível
                            #  ler", não "não é um protocolo".
                            pass

                if dados is None:
                    if usou_leitor_escaneado:
                        observacao = "Não foi possível ler o documento"
                    else:
                        observacao = "Não é um protocolo dos Correios"
                        nao_e_protocolo = True
                else:
                    aceito, motivo = conferir_contagem_protocolo(dados)
                    if not aceito and usou_leitor_escaneado and not ja_tentou_de_novo:
                        #  Uma re-tentativa em qualidade maior, como já se faz
                        #  quando o CNPJ sai com checksum inválido — só faz
                        #  sentido quando a leitura original já veio do leitor
                        #  de escaneados; texto nativo não confere de novo por
                        #  esse caminho, e uma falha aqui não pode apagar o
                        #  motivo original (é a informação que o funcionário
                        #  usa pra conferir o papel).
                        dpi_maior = proximo_dpi_maior(dpi)
                        if dpi_maior and dpi_maior != dpi:
                            try:
                                texto_maior = extrair_texto_escaneado(caminho, dpi=dpi_maior)
                                novos = extrair_dados_protocolo_correio(texto_maior)
                                if novos:
                                    aceito_novo, motivo_novo = conferir_contagem_protocolo(novos)
                                    #  Só adota a 2ª leitura quando ela ACEITA a
                                    #  contagem — se a 2ª leitura vier pior (ex:
                                    #  embaralhou o cabeçalho e perdeu o
                                    #  "Listando"), manter a 1ª leitura preserva
                                    #  o código e o motivo que o funcionário usa
                                    #  pra conferir o papel.
                                    if aceito_novo:
                                        dados, aceito, motivo = novos, aceito_novo, motivo_novo
                            except Exception:
                                #  A re-tentativa falhou (ex.: leitor de
                                #  escaneados indisponível) — mantém o motivo
                                #  já apurado pela 1ª conferência.
                                pass
                    if aceito:
                        unidades = dados["total_impresso"]
                    else:
                        observacao = motivo
            except Exception as e:
                if isinstance(e, RuntimeError) and "leitor de documentos escaneados" in str(e):
                    #  Mensagem técnica de logica.py, não tocada — aqui só se
                    #  traduz pra linguagem de leigo antes de ir pra planilha.
                    observacao = "Não foi possível ler protocolos escaneados nesta máquina."
                else:
                    observacao = f"Erro ao ler: {e}"

            if unidades is not None:
                try:
                    #  `carimbados` ainda não foi incrementado: é o índice
                    #  0-based deste arquivo entre os efetivamente carimbados.
                    pasta_destino = (caminho_do_lote(pasta_saida, carimbados, tamanho_lote)
                                      if separar_em_lotes else pasta_saida)
                    self._carimbar_protocolo(caminho, pasta_destino, nome, dados,
                                             unidades, tarifa, config, codigos,
                                             vencimento=vencimento)
                    carimbados += 1
                except Exception as e:
                    #  Erro ao carimbar não desfaz a cobrança (o valor é
                    #  devido pela entrega, não pelo carimbo ter dado certo),
                    #  mas `observacao` fica marcada aqui pra alimentar o
                    #  motivo tanto na planilha quanto no painel — as duas
                    #  telas precisam contar a mesma história sobre este
                    #  arquivo.
                    observacao = f"Erro ao carimbar: {e}"
                    falhas_carimbo += 1
            elif nao_e_protocolo:
                ignorados += 1
            else:
                #  Cobre tanto "contagem a conferir" (dados veio, mas a
                #  conferência recusou) quanto "não foi possível ler"/erro de
                #  leitura (dados ficou None sem ser por não-ser-protocolo) —
                #  as duas são cobranças que ficaram de fora do lote, mesma
                #  pendência sob a ótica de quem vai olhar o resumo.
                pendentes += 1

            linha_protocolo = linha_planilha_protocolo(
                nome, dados, self.cadastro, tarifa, unidades, observacao)
            linhas.append(linha_protocolo)
            #  Observação final, já com o que `linha_planilha_protocolo`
            #  acrescenta por conta própria (ex.: "Código não cadastrado") —
            #  é o texto que vai pra planilha, então é esse que o painel
            #  precisa mostrar, não a `observacao` local de antes da chamada
            #  (que fica sem esse acréscimo). Guardada à parte de `observacao`
            #  porque a resolução manual (`_acao_informar_valor`) reusa a
            #  RAW `observacao` como semente — repassar a já mesclada faria
            #  "Código não cadastrado" aparecer em dobro na 2ª mesclagem.
            observacao_planilha = linha_protocolo[-1]

            registro_cadastro = self.cadastro.get(
                codigos.get((dados or {}).get("codigo") or "", "")) or None
            if registro_cadastro:
                condominio_painel = f"{dados['codigo']} {registro_cadastro['nome']}"
            elif dados:
                condominio_painel = dados.get("condominio", "")
            else:
                condominio_painel = ""

            registro_painel = {
                "arquivo": nome,
                "condominio": condominio_painel,
                "codigo": (dados or {}).get("codigo"),
                "caminho": caminho,
                "indice_linha": len(linhas) - 1,
            }
            if unidades is not None:
                registro_painel["unidades"] = unidades
                #  Decimal até a hora de exibir — só a escrita na planilha
                #  converte pra float (é o Excel que soma float, não o
                #  painel).
                registro_painel["valor"] = valor_protocolo(unidades, tarifa)
                #  Mostra o MESMO motivo que foi pra planilha — inclui tanto
                #  "Erro ao carimbar: ..." (quando `_carimbar_protocolo`
                #  levantou exceção) quanto "Código não cadastrado" (quando o
                #  carimbo lateral saiu sem código por falta de cadastro):
                #  antes só o primeiro caso aparecia aqui, e uma linha
                #  CALCULADOS podia parecer limpa quando na verdade saiu sem
                #  o carimbo de identificação.
                if observacao_planilha:
                    registro_painel["motivo"] = observacao_planilha
                #  `motivo_original` (sem o acréscimo de código) é o que
                #  `_acao_informar_valor` usa como semente ao resolver essa
                #  linha à mão mais tarde — repassar a já mesclada duplicaria
                #  o "Código não cadastrado".
                registro_painel["motivo_original"] = observacao
                processados_painel.append(registro_painel)
            elif nao_e_protocolo:
                #  Reusa `registro_painel` (que já traz `caminho` e
                #  `indice_linha`) em vez de um dict de dois campos: a grade
                #  do painel mostra TODAS as linhas da planilha, ignorados
                #  inclusive, e a prévia do documento precisa do caminho para
                #  funcionar em qualquer linha selecionada.
                registro_painel["motivo"] = observacao_planilha
                ignorados_painel.append(registro_painel)
            else:
                registro_painel["motivo"] = observacao_planilha
                registro_painel["motivo_original"] = observacao
                pendentes_painel.append(registro_painel)

            self.after(0, lambda v=indice: self.barra_protocolos.configure(value=v))

        planilha_salva = True
        try:
            salvar_planilha_protocolo(destino, linhas)
        except Exception as e:
            #  Ao contrário da aba 2 (onde essa falha aborta o lote), aqui o
            #  painel precisa abrir mesmo assim: as pendências já foram
            #  apuradas e reprocessar o lote inteiro é caro (cada protocolo é
            #  lido na melhor qualidade). O usuário resolve um pendente à
            #  mão — o que já regrava a planilha inteira — ou reabre o
            #  painel depois de fechar o Excel.
            #  `erro=e` como argumento padrão: o `except ... as e` desvincula
            #  `e` ao sair do bloco, e o `self.after` só executa a lambda bem
            #  depois — sem o argumento padrão, ela mostraria "cannot access
            #  free variable 'e'" em vez da mensagem amigável.
            planilha_salva = False
            #  A frase final muda conforme existe ou não pendente pra
            #  resolver — sem pendente nenhum (lote inteiro calculado com
            #  sucesso), "resolva um pendente" seria uma instrução
            #  impossível de seguir; o botão "Abrir planilha" do painel
            #  tenta gravar de novo nesse caso (ver
            #  `_acao_abrir_planilha_protocolos`).
            if pendentes:
                como_regravar = ("A lista de protocolos não foi perdida — ela é "
                                  "regravada assim que você resolver um pendente "
                                  "no painel a seguir.")
            else:
                como_regravar = ("A lista de protocolos não foi perdida — no "
                                  "painel a seguir, use o botão \"Abrir planilha\" "
                                  "para tentar gravar de novo.")
            self.after(0, lambda erro=e, msg=como_regravar: messagebox.showwarning(
                "Planilha não salva",
                f"Não foi possível gravar a planilha:\n{erro}\n\n"
                f"Se ela estiver aberta no Excel, feche o arquivo. {msg}"))

        #  A planilha guarda float (é o que o Excel soma), mas o total do
        #  resumo volta para Decimal antes de somar: em lote grande, somar
        #  float acumula erro de centavo, e aqui o número é dinheiro.
        indice_coluna_valor = [c[0] for c in COLUNAS_PROTOCOLO].index("Valor")
        total_valor = sum(
            (Decimal(str(l[indice_coluna_valor])) for l in linhas
             if isinstance(l[indice_coluna_valor], float)),
            Decimal("0.00"),
        )
        #  O valor é devido pela entrega, não pelo carimbo ter dado certo —
        #  então o total cobra também o(s) protocolo(s) que falharam ao
        #  carimbar (ver falhas_carimbo abaixo). O texto do resumo precisa
        #  deixar isso explícito, sem ambiguidade, ou parece que o total só
        #  cobre os "carimbados".
        resumo = f"{carimbados} protocolo(s) carimbado(s) · {formatar_reais(total_valor)} no total"
        if falhas_carimbo:
            plural = "não carimbado" if falhas_carimbo == 1 else "não carimbados"
            resumo += (f" (inclui {falhas_carimbo} {plural} — "
                       "a cobrança vale mesmo sem o carimbo)")
        if pendentes:
            resumo += f" · {pendentes} pendente(s) (não lido(s) ou contagem a conferir)"
        if ignorados:
            resumo += f" · {ignorados} ignorado(s) (não é protocolo dos Correios)"
        if not planilha_salva:
            if pendentes:
                resumo += " · planilha não salva — resolva um pendente para regravar"
            else:
                resumo += (" · planilha não salva — use o botão \"Abrir planilha\" "
                            "no painel para tentar gravar de novo")

        self.after(0, lambda: self.label_status_protocolos.configure(text=resumo))
        self.after(0, self._atualizar_botao_protocolos)

        self._ctx_protocolos = {
            "pasta_saida": pasta_saida,
            "config": config,
            "destino": destino,
            "codigos": codigos,
            "linhas": linhas,
            #  Preferência de separar em lotes + o contador de arquivos já
            #  carimbados neste processamento: a resolução manual de um
            #  pendente (`_acao_informar_valor`) precisa continuar essa
            #  mesma contagem para cair na subpasta "Lote NN" certa, em vez
            #  de ir parar solta na raiz — ver `caminho_do_lote`.
            "separar_em_lotes": separar_em_lotes,
            "tamanho_lote": tamanho_lote,
            "carimbados": carimbados,
            #  Tarifa do lote. Necessária para recalcular o valor quando as
            #  unidades são editadas na grade do painel. Já era lida por
            #  `ctx.get("tarifa")` em _acao_escolher_condominio, mas nunca
            #  chegava a ser gravada aqui — sempre vinha None.
            "tarifa": tarifa,
            #  Perguntado no início do lote: vai carimbado no PDF e depois
            #  para a coluna `vencimento` da planilha de despesas. Tem que ser
            #  o MESMO nos dois lugares, senão o Paybox não anexa o documento.
            "vencimento": vencimento,
            #  Mesma ideia do vencimento: dado do lote, gravado direto na
            #  planilha para ninguém precisar abrir o arquivo no Excel.
            "chave": chave,
        }
        resultado = {
            "total": total,
            "total_valor": total_valor,
            "processados": processados_painel,
            "pendentes": pendentes_painel,
            "ignorados": ignorados_painel,
        }
        self.after(0, lambda: self.mostrar_resultado_protocolos(resultado))

    def _carimbar_protocolo(self, caminho, pasta_saida, nome, dados, unidades,
                            tarifa, config, codigos, valor_manual=None,
                            vencimento=None):
        """
        Três carimbos num passe só: o lateral rotacionado com o código (igual
        ao da aba 1, para a IA do Superlógica continuar lendo), o valor no
        topo direito, e o bloco do Paybox (fornecedor, vencimento e valor —
        ver `montar_bloco_paybox`), que é o que faz o documento se anexar
        sozinho à despesa. Código fora do cadastro carimba só o valor — o
        valor não depende do cadastro, e perder a cobrança por isso seria pior.

        `vencimento` ausente simplesmente omite o bloco do Paybox: o resto do
        carimbo continua igual, e o documento volta a exigir anexo manual.
        """
        valor = valor_manual if valor_manual is not None else valor_protocolo(unidades, tarifa)

        #  O carimbo traz SÓ o valor, nunca a conta que o gerou. O Superlógica
        #  lê esse carimbo, e uma linha como "15 un × R$ 3,85 = R$ 57,75" tem
        #  dois valores em reais: ele pode capturar a tarifa no lugar do total.
        #  Com um número só não há o que confundir. A conta continua na
        #  planilha, nas colunas Unidades e Tarifa.
        texto_valor = formatar_reais(valor)

        cnpj = codigos.get(dados.get("codigo") or "")
        registro = self.cadastro.get(cnpj) if cnpj else None

        #  Posição do carimbo de valor a partir do tamanho real da 1ª página
        #  do PDF — uma margem fixa pra A4 cairia fora da folha em páginas
        #  menores.
        leitor_medida = PdfReader(caminho)
        primeira_pagina = leitor_medida.pages[0]
        largura_pagina = float(primeira_pagina.mediabox.width)
        altura_pagina = float(primeira_pagina.mediabox.height)
        leitor_medida.close()
        x_valor = largura_pagina - MARGEM_VALOR_PROTOCOLO
        y_valor = altura_pagina - MARGEM_VALOR_PROTOCOLO

        config_carimbo = dict(config)
        config_carimbo["angulo"] = 90
        if registro:
            texto_lateral = montar_texto_protocolo_correio(
                dados["codigo"], registro["nome"], cnpj)
        else:
            #  Sem cadastro não há nome nem CNPJ para a linha lateral; o
            #  carimbo principal vira o próprio valor, no topo direito.
            texto_lateral = texto_valor
            config_carimbo["angulo"] = 0
            config_carimbo["centralizado"] = False
            config_carimbo["alinhamento"] = "direita"
            config_carimbo["x"] = x_valor
            config_carimbo["y"] = y_valor

        extras = []
        if registro:
            extras.append({
                "texto": texto_valor,
                "fonte": config["fonte"],
                "tamanho": config["tamanho"],
                "cor": config["cor"],
                "x": x_valor,
                "y": y_valor,
                "centralizado": False,
                "angulo": 0,
                "alinhamento": "direita",
            })

        #  Bloco que o Paybox lê para anexar o protocolo à despesa sozinho.
        #  Negrito e corpo maior que o resto de propósito: o Superlógica faz
        #  OCR da imagem da página, e letra pequena é o que mais atrapalha.
        if vencimento is not None:
            extras.append({
                "texto": montar_bloco_paybox(vencimento, valor),
                "fonte": "Helvetica-Bold",
                "tamanho": TAMANHO_BLOCO_PAYBOX,
                "cor": config["cor"],
                "x": largura_pagina - MARGEM_DIREITA_BLOCO_PAYBOX,
                "y": Y_BLOCO_PAYBOX,
                "centralizado": False,
                "angulo": 0,
                "alinhamento": "direita",
            })

        caminho_saida = os.path.join(pasta_saida, nome)
        processar_pdf(caminho, caminho_saida, texto_lateral, config_carimbo, extras)

    def mostrar_resultado_protocolos(self, resultado):
        """
        Painel de encerramento da aba 3 — a PLANILHA em si, editável, com a
        prévia do documento ao lado.

        Uma grade só, em vez das três tabelas anteriores (pendentes,
        calculados, ignorados): a planilha já contém tudo que elas mostravam,
        e mantê-las juntas fazia a mesma linha aparecer duas vezes na tela. A
        distinção não some — vira ESTADO da linha (cor + filtro).

        Ver docs/superpowers/specs/2026-08-29-planilha-embutida-previa-design.md.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        self._resultado_protocolos = resultado

        processados, pendentes, total_valor = self._listas_do_painel_protocolos(resultado)

        janela = self._montar_painel_resultado(
            "Resultado dos protocolos", "1180x720", (900, 560))

        self._cartoes_protocolos = self._montar_faixa_cartoes(janela, [
            (str(len(processados)), "PROTOCOLOS", False),
            (str(len(pendentes)), "PENDENTES", True),
            (formatar_reais(total_valor), "TOTAL", False),
        ], fonte=fonte)

        ctk.CTkFrame(janela, height=1, corner_radius=0,
                     fg_color=tema["borda"]).pack(fill="x", padx=24, pady=(24, 0))

        # --- RODAPÉ (fixo embaixo, montado antes para não ser empurrado) ---
        rodape = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        rodape.pack(side="bottom", fill="x", padx=24, pady=16)

        self._label_total_protocolos = ctk.CTkLabel(
            rodape, text=f"{resultado.get('total', 0)} arquivo(s) no total.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="w")
        self._label_total_protocolos.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            rodape, text="Gerar planilha do Superlógica", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=self._acao_gerar_despesas,
        ).pack(side="right", padx=(0, 8))

        ctk.CTkButton(
            rodape, text="Abrir planilha", corner_radius=0,
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=self._acao_abrir_planilha_protocolos,
        ).pack(side="right")

        # --- CORPO: grade à esquerda, prévia à direita ---
        corpo = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        corpo.pack(side="top", fill="both", expand=True, padx=24, pady=(16, 0))
        corpo.columnconfigure(0, weight=1)
        corpo.rowconfigure(1, weight=1)

        cabecalho = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        cabecalho.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(
            cabecalho, text="PLANILHA DO LOTE", font=(fonte, 11, "bold"),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(side="left")

        #  Filtro por estado. Substitui a separação em três tabelas: quem quer
        #  ver só o que precisa de ação clica em "Pendentes".
        self._filtro_protocolos = ctk.CTkSegmentedButton(
            cabecalho, values=["Tudo", "Pendentes", "Calculados", "Ignorados"],
            corner_radius=0, font=(fonte, 12),
            fg_color=tema["superficie"], selected_color=tema["acento"],
            selected_hover_color=tema["acento_hover"],
            unselected_color=tema["superficie"],
            unselected_hover_color=tema["borda"],
            text_color=tema["texto"],
            command=lambda _v: self._preencher_grade_protocolos())
        self._filtro_protocolos.set("Tudo")
        self._filtro_protocolos.pack(side="right")

        self._montar_grade_protocolos(corpo, tema, fonte)
        self._montar_previa_protocolos(corpo, tema, fonte)

        # --- Ações da linha selecionada ---
        acoes = ctk.CTkFrame(corpo, corner_radius=0, fg_color=tema["fundo"])
        acoes.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self._botao_valor_protocolo = ctk.CTkButton(
            acoes, text="Informar valor", corner_radius=0, state="disabled",
            fg_color=tema["borda"], hover_color=tema["acento_hover"],
            text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
            command=lambda: self._acao_informar_valor(self._linha_selecionada_protocolo()))
        self._botao_valor_protocolo.pack(side="left", padx=(0, 8))

        self._botao_cond_protocolo = ctk.CTkButton(
            acoes, text="Escolher condomínio", corner_radius=0, state="disabled",
            fg_color=tema["borda"], hover_color=tema["acento_hover"],
            text_color=tema["sobre_acento"], border_width=0, font=(fonte, 13),
            command=lambda: self._acao_escolher_condominio(self._linha_selecionada_protocolo()))
        self._botao_cond_protocolo.pack(side="left", padx=(0, 8))

        self._botao_abrir_protocolo = ctk.CTkButton(
            acoes, text="Abrir PDF", corner_radius=0, state="disabled",
            fg_color="transparent", hover_color=tema["superficie"],
            border_width=1, border_color=tema["borda_forte"],
            text_color=tema["texto"], font=(fonte, 13),
            command=lambda: self._acao_abrir_pdf_protocolo(self._linha_selecionada_protocolo()))
        self._botao_abrir_protocolo.pack(side="left")

        ctk.CTkLabel(
            acoes,
            text="Duplo clique numa célula para editar: código, unidades, valor, observação.",
            font=(fonte, 12), text_color=tema["texto_terciario"], anchor="e",
        ).pack(side="right")

        self._preencher_grade_protocolos()

    def _montar_grade_protocolos(self, parent, tema, fonte):
        """A planilha como Treeview: uma linha por linha do arquivo gerado."""
        frame = ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["fundo"])
        frame.grid(row=1, column=0, sticky="nsew", padx=(0, 16))

        colunas = [c[0] for c in COLUNAS_PROTOCOLO]
        larguras = [180, 165, 60, 70, 70, 90, 200]
        ancoras = ["w", "w", "center", "center", "e", "e", "w"]

        tabela = ttk.Treeview(frame, columns=colunas, show="headings", height=16)
        for chave, largura, ancora in zip(colunas, larguras, ancoras):
            tabela.heading(chave, text=chave)
            tabela.column(chave, width=largura, anchor=ancora)
        tabela.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
        tabela.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")

        #  Estado da linha por cor — o que antes eram três tabelas. Cobalto
        #  para o que precisa de ação, tom apagado para o ignorado (não é
        #  protocolo, não vira cobrança).
        tabela.tag_configure("pendente", background=tema["superficie"],
                             foreground=tema["acento"])
        tabela.tag_configure("ignorado", foreground=tema["texto_terciario"])

        tabela.bind("<<TreeviewSelect>>", self._ao_selecionar_linha_protocolo)
        tabela.bind("<Double-1>", self._ao_duplo_clique_grade)
        self._grade_protocolos = tabela
        self._linha_grade_por_iid = {}

    def _montar_previa_protocolos(self, parent, tema, fonte):
        """Coluna da direita: primeira página do documento da linha selecionada."""
        largura_painel = LARGURA_PREVIA // 2 + 48
        painel = ctk.CTkFrame(parent, corner_radius=0, fg_color=tema["superficie"],
                              width=largura_painel)
        painel.grid(row=1, column=1, sticky="ns")
        painel.grid_propagate(False)

        ctk.CTkLabel(
            painel, text="DOCUMENTO", font=(fonte, 11, "bold"),
            text_color=tema["texto_secundario"], anchor="w",
        ).pack(fill="x", padx=12, pady=(12, 8))

        self._label_previa = ctk.CTkLabel(
            painel, text="Selecione uma linha para ver o documento.",
            font=(fonte, 12), text_color=tema["texto_terciario"],
            wraplength=largura_painel - 24, justify="left", anchor="n")
        self._label_previa.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        #  Cache por caminho: reabrir a mesma linha não re-renderiza. Não
        #  pré-renderiza o lote — 15 arquivos seria barato, 500 não.
        self._cache_previa = {}

    def _estado_da_linha_protocolo(self, indice):
        """Devolve "pendente", "calculado" ou "ignorado" para a linha `indice`."""
        resultado = self._resultado_protocolos or {}
        for chave, estado in (("pendentes", "pendente"),
                              ("ignorados", "ignorado"),
                              ("processados", "calculado")):
            for registro in resultado.get(chave, []) or []:
                if registro.get("indice_linha") == indice:
                    #  Processado sem ID SL ainda precisa de ação: conta como
                    #  pendente, mesma regra de _listas_do_painel_protocolos.
                    if estado == "calculado" and self._falta_id_sl(registro):
                        return "pendente"
                    return estado
        return "calculado"

    def _registro_da_linha_protocolo(self, indice):
        """Registro do painel (com `caminho`) correspondente à linha `indice`."""
        resultado = self._resultado_protocolos or {}
        for chave in ("pendentes", "processados", "ignorados"):
            for registro in resultado.get(chave, []) or []:
                if registro.get("indice_linha") == indice:
                    return registro
        return None

    def _preencher_grade_protocolos(self):
        """Reescreve a grade a partir de `ctx["linhas"]`, aplicando o filtro."""
        tabela = getattr(self, "_grade_protocolos", None)
        if tabela is None:
            return
        ctx = getattr(self, "_ctx_protocolos", None) or {}
        linhas = ctx.get("linhas", []) or []
        filtro = "Tudo"
        if getattr(self, "_filtro_protocolos", None) is not None:
            filtro = self._filtro_protocolos.get() or "Tudo"

        selecionado_antes = None
        if tabela.selection():
            selecionado_antes = tabela.selection()[0]

        tabela.delete(*tabela.get_children())
        self._linha_grade_por_iid = {}

        for indice, linha in enumerate(linhas):
            estado = self._estado_da_linha_protocolo(indice)
            if filtro == "Pendentes" and estado != "pendente":
                continue
            if filtro == "Calculados" and estado != "calculado":
                continue
            if filtro == "Ignorados" and estado != "ignorado":
                continue

            iid = f"lin{indice}"
            self._linha_grade_por_iid[iid] = indice
            tabela.insert("", "end", iid=iid,
                          tags=(estado,) if estado != "calculado" else (),
                          values=self._valores_visiveis_da_linha(linha))

        #  Mantém a linha selecionada depois de regravar (edição não deve
        #  jogar o foco pro topo da grade).
        if selecionado_antes and tabela.exists(selecionado_antes):
            tabela.selection_set(selecionado_antes)
        self._ao_selecionar_linha_protocolo()

    def _valores_visiveis_da_linha(self, linha):
        """Formata a linha da planilha para exibição (dinheiro em reais)."""
        valores = list(linha)
        for coluna in (4, COL_VALOR):
            valores[coluna] = ("" if valores[coluna] in (None, "")
                               else formatar_reais(valores[coluna]))
        if valores[COL_UNIDADES] in (None, ""):
            valores[COL_UNIDADES] = ""
        return valores

    def _linha_selecionada_protocolo(self):
        """Registro do painel da linha selecionada na grade, ou None."""
        tabela = getattr(self, "_grade_protocolos", None)
        if tabela is None:
            return None
        selecao = tabela.selection()
        if not selecao:
            return None
        indice = self._linha_grade_por_iid.get(selecao[0])
        return None if indice is None else self._registro_da_linha_protocolo(indice)

    def _ao_selecionar_linha_protocolo(self, _evento=None):
        """Atualiza os botões de ação e a prévia do documento."""
        dados = self._linha_selecionada_protocolo()
        tema = self.tema_atual

        #  Cada botão liga só para o que falta NAQUELA linha (regra herdada do
        #  painel anterior: oferecer os dois sempre fazia a pessoa tentar o
        #  que não resolve). Botão cobalto desabilitado precisa apagar o
        #  fg_color, senão fica azul e parece clicável.
        falta_valor = dados is not None and dados.get("valor") is None
        falta_cond = dados is not None and self._falta_id_sl(dados)
        for botao, ligado in ((self._botao_valor_protocolo, falta_valor),
                              (self._botao_cond_protocolo, falta_cond)):
            botao.configure(state="normal" if ligado else "disabled",
                            fg_color=tema["acento"] if ligado else tema["borda"])
        self._botao_abrir_protocolo.configure(
            state="normal" if dados and dados.get("caminho") else "disabled")

        self._mostrar_previa(dados)

    def _mostrar_previa(self, dados):
        """Renderiza (com cache) a 1ª página do documento da linha."""
        label = getattr(self, "_label_previa", None)
        if label is None:
            return
        if not dados or not dados.get("caminho"):
            label.configure(image=None, text="Selecione uma linha para ver o documento.")
            return

        caminho = dados["caminho"]
        if caminho not in self._cache_previa:
            self._cache_previa[caminho] = renderizar_previa_pdf(
                caminho, largura=LARGURA_PREVIA // 2)
        imagem, motivo = self._cache_previa[caminho]

        if imagem is None:
            label.configure(image=None, text=motivo)
            return

        foto = ctk.CTkImage(light_image=imagem, dark_image=imagem,
                            size=(imagem.width, imagem.height))
        #  Referência viva: sem isso o coletor de lixo leva a imagem e o
        #  label aparece vazio.
        self._foto_previa = foto
        label.configure(image=foto, text="")

    def _ao_duplo_clique_grade(self, evento):
        """
        Duplo clique numa célula editável abre um campo por cima dela.

        Colunas somente leitura não abrem editor nenhum: "Arquivo" é a
        identidade da linha, "Condomínio" muda junto com o código (não
        sozinha) e "Tarifa" é do lote inteiro. Clicar nelas apenas seleciona
        a linha, como antes.
        """
        tabela = self._grade_protocolos
        iid = tabela.identify_row(evento.y)
        coluna_id = tabela.identify_column(evento.x)
        if not iid or not coluna_id:
            return
        coluna = int(coluna_id[1:]) - 1          # "#3" -> 2
        if coluna not in COLUNAS_EDITAVEIS:
            return
        self._abrir_editor_celula(iid, coluna)

    def _abrir_editor_celula(self, iid, coluna):
        """Entry sobreposto à célula, confirmando com Enter e cancelando com Esc."""
        tabela = self._grade_protocolos
        caixa = tabela.bbox(iid, f"#{coluna + 1}")
        if not caixa:                            # linha fora da área visível
            return
        x, y, largura, altura = caixa

        indice = self._linha_grade_por_iid.get(iid)
        if indice is None:
            return
        ctx = getattr(self, "_ctx_protocolos", None) or {}
        linha = (ctx.get("linhas") or [])[indice]

        #  O editor recebe o valor CRU (não o formatado em reais): quem edita
        #  digita "46,20", não "R$ 46,20" — e converter_valor_digitado aceita
        #  as duas formas de qualquer jeito.
        bruto = linha[coluna]
        texto_inicial = "" if bruto in (None, "") else str(bruto)

        editor = tk.Entry(tabela, borderwidth=1, relief="solid")
        editor.place(x=x, y=y, width=largura, height=altura)
        editor.insert(0, texto_inicial)
        editor.select_range(0, "end")
        editor.focus_set()

        def fechar():
            try:
                editor.destroy()
            except Exception:
                pass

        def confirmar(_evento=None):
            texto = editor.get()
            fechar()
            self._aplicar_edicao_grade(indice, coluna, texto)

        editor.bind("<Return>", confirmar)
        editor.bind("<Escape>", lambda _e: fechar())
        #  Perder o foco confirma, como numa planilha de verdade: clicar fora
        #  depois de digitar não pode descartar o que foi escrito em silêncio.
        editor.bind("<FocusOut>", confirmar)

    def _aplicar_edicao_grade(self, indice, coluna, texto):
        """
        Valida a edição, recarimba o PDF quando ela muda o papel, e só então
        grava a linha e a planilha.

        A ordem importa: código, unidades e valor estão IMPRESSOS no
        documento (carimbo lateral e bloco do Paybox), e o Superlógica lê o
        papel por OCR. Aplicar na planilha sem recarimbar criaria divergência
        silenciosa — o Paybox não acharia o lançamento e o arquivo cairia na
        fila manual sem erro nenhum.
        """
        ctx = getattr(self, "_ctx_protocolos", None)
        if not ctx:
            messagebox.showerror(
                "Erro", "O contexto do processamento se perdeu. Rode o lote de novo.",
                parent=self._janela_resultado)
            return

        linha = ctx["linhas"][indice]
        if str(linha[coluna] if linha[coluna] is not None else "") == texto.strip():
            return                                # nada mudou

        tarifa = ctx.get("tarifa")
        valor, erro = validar_edicao_protocolo(coluna, texto, self.cadastro, tarifa)
        if erro:
            messagebox.showwarning("Edição recusada", erro,
                                    parent=self._janela_resultado)
            return

        nova_linha = aplicar_edicao_na_linha(linha, coluna, valor, self.cadastro, tarifa)
        registro = self._registro_da_linha_protocolo(indice)

        if coluna in COLUNAS_QUE_RECARIMBAM:
            if registro is None or not registro.get("caminho"):
                messagebox.showwarning(
                    "Sem o documento",
                    "Não dá para recarimbar: o arquivo desta linha não está "
                    "mais acessível. A edição não foi aplicada.",
                    parent=self._janela_resultado)
                return
            if not self._recarimbar_linha(registro, nova_linha, ctx):
                return                            # falhou: edição não se aplica

        ctx["linhas"][indice] = nova_linha
        self._sincronizar_registro_com_linha(registro, nova_linha)

        try:
            salvar_planilha_protocolo(ctx["destino"], ctx["linhas"])
        except Exception as e:
            messagebox.showwarning(
                "Planilha não gravada",
                f"A edição foi aplicada e o PDF recarimbado, mas a planilha "
                f"não pôde ser gravada:\n{e}\n\nFeche o arquivo no Excel e "
                f'use "Abrir planilha" para gravar de novo.',
                parent=self._janela_resultado)

        self._atualizar_painel_protocolos()

    def _recarimbar_linha(self, registro, nova_linha, ctx):
        """
        Regrava o PDF com os dados da linha editada. Devolve True se deu
        certo; em caso de falha avisa e devolve False, para o chamador
        descartar a edição (mesma regra de `resolver_protocolo_manual`: não
        faz sentido dar por resolvido um arquivo que não saiu carimbado).
        """
        registro_dados = {"codigo": nova_linha[COL_CODIGO],
                          "condominio": f"{nova_linha[COL_CODIGO]} {nova_linha[1]}".strip()}
        unidades = nova_linha[COL_UNIDADES]
        valor = nova_linha[COL_VALOR]
        try:
            self._carimbar_protocolo(
                registro["caminho"],
                registro.get("pasta_destino") or ctx["pasta_saida"],
                registro["arquivo"], registro_dados, unidades,
                ctx.get("tarifa") if unidades else None,
                ctx["config"], ctx["codigos"],
                valor_manual=None if unidades else Decimal(str(valor)),
                vencimento=ctx.get("vencimento"))
            return True
        except Exception as e:
            messagebox.showerror(
                "Erro ao recarimbar",
                f"A edição NÃO foi aplicada, porque o PDF não pôde ser "
                f"recarimbado:\n{e}\n\nPapel e planilha precisam continuar "
                f"iguais — por isso nada foi alterado.",
                parent=self._janela_resultado)
            return False

    def _sincronizar_registro_com_linha(self, registro, nova_linha):
        """
        Reflete a edição no registro do painel, para os botões de ação e a
        classificação pendente/calculado enxergarem o novo estado sem
        reprocessar o lote.
        """
        if registro is None:
            return
        registro["codigo"] = nova_linha[COL_CODIGO] or None
        registro["condominio"] = (f"{nova_linha[COL_CODIGO]} {nova_linha[1]}".strip()
                                   if nova_linha[COL_CODIGO] else nova_linha[1] or "")
        registro["unidades"] = nova_linha[COL_UNIDADES]
        registro["valor"] = (None if nova_linha[COL_VALOR] in (None, "")
                             else Decimal(str(nova_linha[COL_VALOR])))
        registro["motivo"] = nova_linha[COL_OBSERVACAO] or ""

        #  Uma linha editada pode deixar de ser "ignorada": se alguém deu
        #  código e valor a um arquivo que o programa não reconheceu como
        #  protocolo, ele passa a ser cobrável e precisa sair da lista de
        #  ignorados, senão não entraria na planilha de despesas.
        resultado = self._resultado_protocolos or {}
        ignorados = resultado.get("ignorados") or []
        if registro in ignorados and registro.get("valor") is not None:
            ignorados.remove(registro)
            resultado.setdefault("processados", []).append(registro)

    def _janela_para_dialogo(self):
        """
        Janela que deve ser a "mãe" de uma janelinha modal: o painel de
        resultado quando ele está aberto, senão a janela principal.

        Existe porque uma mesma pergunta pode ser feita antes de o painel
        existir. Fixar `self._janela_resultado` como pai quebrava com
        AttributeError quando a pergunta vinha antes do processamento.
        """
        janela = getattr(self, "_janela_resultado", None)
        try:
            if janela is not None and janela.winfo_exists():
                return janela
        except Exception:
            pass
        return self

    def _listas_do_painel_protocolos(self, resultado):
        """
        Divide o `resultado` do lote no que cada seção do painel mostra:
        (calculados, pendentes, total).

        Protocolo com valor mas sem o ID do Superlógica também precisa de
        ação — só que de outra: dizer de qual condomínio ele é. Aparece junto
        dos pendentes para que tudo que exige atenção fique num lugar só. A
        separação é de EXIBIÇÃO: no `resultado` ele continua em `processados`,
        porque `lancamentos_de_despesa` precisa distinguir "sem valor" de "sem
        condomínio" para dizer onde cada um se resolve.
        """
        processados = resultado.get("processados", []) or []
        pendentes = list(resultado.get("pendentes", []) or [])
        pendentes += [p for p in processados if self._falta_id_sl(p)]
        processados = [p for p in processados if not self._falta_id_sl(p)]
        return processados, pendentes, resultado.get("total_valor", Decimal("0.00"))

    def _atualizar_painel_protocolos(self):
        """
        Atualiza o painel depois de resolver ou editar uma linha, SEM
        remontá-lo.

        O que evita a piscada não é a velocidade — é NÃO destruir widget
        nenhum. Remontar recria todos os widgets CustomTkinter e a janela
        fica visivelmente vazia nesse intervalo.

        Com a grade única não existe mais "mudança de estrutura" (antes, uma
        seção que passava de texto para tabela obrigava a remontagem): a
        grade está sempre lá, mesmo vazia. Só volta a remontar se a janela
        tiver sido fechada.
        """
        resultado = self._resultado_protocolos
        janela = getattr(self, "_janela_resultado", None)
        if janela is None or not janela.winfo_exists():
            self.mostrar_resultado_protocolos(resultado)
            return

        processados, pendentes, total_valor = self._listas_do_painel_protocolos(resultado)

        cartoes = getattr(self, "_cartoes_protocolos", None)
        for label, texto in zip(getattr(cartoes, "labels_valor", []),
                                 (str(len(processados)), str(len(pendentes)),
                                  formatar_reais(total_valor))):
            label.configure(text=texto)

        self._preencher_grade_protocolos()

    def _falta_id_sl(self, item):
        """Se este protocolo ainda não tem como virar lançamento de despesa:
        sem código lido, ou com código que não está no cadastro, ou cadastrado
        sem o campo "ID SL"."""
        codigo = (item or {}).get("codigo") or ""
        if not codigo:
            return True
        cnpj = _codigos_do_cadastro(self.cadastro).get(codigo)
        registro = self.cadastro.get(cnpj) if cnpj else None
        return not (registro or {}).get("id_sl")

    def _acao_abrir_pdf_protocolo(self, dados):
        if not dados:
            return
        try:
            os.startfile(dados["caminho"])
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Não foi possível abrir o PDF:\n{e}",
                parent=self._janela_resultado)

    def _acao_abrir_planilha_protocolos(self):
        """Abre a planilha de cobrança gerada por este lote (o entregável
        final do fluxo — sem isso o painel não teria nenhum jeito de chegar
        até ela). A gravação inicial pode ter falhado (ver `planilha_salva`
        em `_processar_protocolos_em_thread`) — nesse caso o arquivo ainda
        não existe. Antes, esse caso só instruía "resolva um pendente", o
        que virava um beco sem saída quando o lote inteiro tinha sido
        calculado com sucesso (nenhum pendente pra resolver) e só a
        gravação tinha falhado: agora o próprio botão tenta gravar de novo
        — todos os dados já estão em `_ctx_protocolos["linhas"]`, não
        precisa reprocessar o lote (caro: cada protocolo é relido na
        melhor qualidade)."""
        ctx = self._ctx_protocolos or {}
        destino = ctx.get("destino")
        if not destino:
            return
        if not os.path.isfile(destino):
            try:
                salvar_planilha_protocolo(destino, ctx.get("linhas", []))
            except Exception as e:
                messagebox.showinfo(
                    "Planilha ainda não salva",
                    f"Não foi possível gravar a planilha agora:\n{e}\n\n"
                    "Se ela estiver aberta no Excel, feche o arquivo e "
                    "clique de novo neste botão.",
                    parent=self._janela_resultado)
                return
        try:
            os.startfile(destino)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Não foi possível abrir a planilha:\n{e}",
                parent=self._janela_resultado)

    def _acao_gerar_despesas(self):
        """
        Gera a planilha de importação de despesas do Superlógica a partir
        deste lote: uma linha por protocolo cobrável, com o condomínio
        preenchido pelo ID SL do cadastro.

        Recusa gerar enquanto houver protocolo cobrável incompleto — e diz
        quais são e onde resolver cada caso, porque "não gera nada" sem a
        lista vira adivinhação.
        """
        resultado = getattr(self, "_resultado_protocolos", None)
        if not resultado:
            return

        lancamentos, travas = lancamentos_de_despesa(resultado, self.cadastro)

        if travas["sem_valor"] or travas["sem_id_sl"]:
            partes = []
            if travas["sem_valor"]:
                partes.append(
                    'Sem valor — resolva no botão "Informar valor":\n  '
                    + "\n  ".join(travas["sem_valor"]))
            if travas["sem_id_sl"]:
                partes.append(
                    'Sem o condomínio identificado — use "Escolher condomínio" '
                    'na lista de pendentes:\n  ' + "\n  ".join(travas["sem_id_sl"]))
            messagebox.showwarning(
                "Ainda não dá para gerar",
                "A planilha de despesas não foi gerada porque estes protocolos "
                "estão incompletos:\n\n" + "\n\n".join(partes),
                parent=self._janela_resultado)
            return

        if not lancamentos:
            messagebox.showinfo(
                "Nada para lançar",
                "Nenhum protocolo deste lote vira despesa.",
                parent=self._janela_resultado)
            return

        #  Vencimento é dado do lote, como a tarifa — perguntado aqui em vez
        #  de digitado no modelo, que é onde a data virava número.
        #  O vencimento vem do que foi perguntado no início do lote e
        #  carimbado nos PDFs. Perguntar de novo aqui abriria a porta para
        #  digitar uma data diferente da que está no papel — e o Paybox só
        #  anexa o documento quando as duas batem.
        vencimento = (self._ctx_protocolos or {}).get("vencimento")
        if vencimento is None:
            vencimento = self._pedir_vencimento_despesas()
            if vencimento is None:
                return
        chave = (self._ctx_protocolos or {}).get("chave")

        #  O modelo vai junto do executável, como o cadastro: quem quiser
        #  trocar fornecedor, categoria ou forma de pagamento edita o arquivo
        #  ao lado do .exe. Só pede pra escolher quando ele não está lá.
        modelo = os.path.join(pasta_base(), NOME_MODELO_DESPESAS)
        if not os.path.isfile(modelo):
            modelo = filedialog.askopenfilename(
                title="Escolha o modelo de despesas do Superlógica",
                filetypes=[("Excel", "*.xlsx")],
                parent=self._janela_resultado)
            if not modelo:
                return

        ctx = self._ctx_protocolos or {}
        pasta_sugerida = os.path.dirname(ctx.get("destino", "")) or None
        destino = filedialog.asksaveasfilename(
            title="Salvar planilha de despesas como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="despesas_superlogica.xlsx",
            initialdir=pasta_sugerida,
            parent=self._janela_resultado)
        if not destino:
            return

        try:
            gerar_planilha_despesas(modelo, destino, lancamentos,
                                    vencimento=vencimento, chave=chave)
        except Exception as e:
            messagebox.showerror(
                "Erro ao gerar",
                f"Não foi possível gerar a planilha de despesas:\n{e}",
                parent=self._janela_resultado)
            return

        plural = "lançamento" if len(lancamentos) == 1 else "lançamentos"
        if messagebox.askyesno(
            "Planilha de despesas pronta",
            f"{len(lancamentos)} {plural} gravado(s) em:\n{destino}\n\nAbrir agora?",
            parent=self._janela_resultado,
        ):
            try:
                os.startfile(destino)
            except Exception as e:
                messagebox.showerror(
                    "Erro", f"Não foi possível abrir a planilha:\n{e}",
                    parent=self._janela_resultado)

    def _acao_escolher_condominio(self, dados, redesenhar=True):
        """
        Diz a qual condomínio um protocolo pertence, quando o programa não
        conseguiu descobrir sozinho — código ilegível (manuscrito por cima) ou
        fora do cadastro. Recarimba o PDF com o código na lateral, atualiza a
        linha da planilha e destrava a geração da planilha de despesas.

        Sem isso, esses protocolos ficavam num beco sem saída: entravam na
        planilha com valor, travavam a geração por falta do ID do Superlógica,
        e nenhum botão da tela resolvia — "Informar valor" não serve, porque o
        valor eles já têm.
        """
        if not dados:
            return
        ctx = getattr(self, "_ctx_protocolos", None)
        if not ctx:
            messagebox.showerror(
                "Erro",
                "O contexto do processamento se perdeu. Rode o lote de novo.",
                parent=self._janela_resultado)
            return

        escolhido = self._pedir_condominio(dados)
        if escolhido is None:
            return

        if not escolhido.get("id_sl"):
            messagebox.showwarning(
                "Condomínio sem o código do Superlógica",
                f"{escolhido['codigo']} {escolhido['nome']} está no cadastro, "
                "mas sem o campo \"ID SL\" preenchido — sem ele o lançamento "
                "não tem como ser importado.\n\nPreencha o ID SL na aba de "
                "Cadastro e escolha o condomínio de novo.",
                parent=self._janela_resultado)
            return

        dados["codigo"] = escolhido["codigo"]
        dados["condominio"] = f"{escolhido['codigo']} {escolhido['nome']}"

        #  Recarimba a partir do PDF original: o arquivo de saída foi gravado
        #  sem o carimbo lateral, porque na hora não havia código. O papel
        #  arquivado é o que vai para o Superlógica, então não pode ficar sem
        #  identificação.
        registro_dados = {"codigo": escolhido["codigo"],
                          "condominio": dados["condominio"]}
        try:
            self._carimbar_protocolo(
                dados["caminho"], dados.get("pasta_destino") or ctx["pasta_saida"],
                dados["arquivo"], registro_dados, dados.get("unidades"),
                ctx.get("tarifa"), ctx["config"], ctx["codigos"],
                valor_manual=dados.get("valor"),
                vencimento=ctx.get("vencimento"))
        except Exception as e:
            messagebox.showerror(
                "Erro ao carimbar",
                f"O condomínio foi anotado, mas o PDF não pôde ser recarimbado:"
                f"\n{e}",
                parent=self._janela_resultado)

        #  Refaz a linha da planilha com o condomínio agora conhecido.
        ctx["linhas"][dados["indice_linha"]] = linha_planilha_protocolo(
            dados["arquivo"], registro_dados, self.cadastro,
            valor_manual=dados.get("valor"))

        try:
            salvar_planilha_protocolo(ctx["destino"], ctx["linhas"])
        except Exception as e:
            messagebox.showwarning(
                "Planilha não atualizada",
                f"O condomínio foi anotado, mas a planilha não pôde ser "
                f"regravada:\n{e}\n\nSe ela estiver aberta no Excel, feche o "
                f"arquivo — a planilha é regravada inteira na próxima correção.",
                parent=self._janela_resultado)

        if redesenhar:
            self._atualizar_painel_protocolos()

    def _pedir_condominio(self, dados):
        """
        Janelinha para escolher o condomínio: campo de busca por nome ou
        código e a lista do cadastro. Devolve o dict do condomínio escolhido,
        ou None se o usuário cancelar.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        janela = ctk.CTkToplevel(self._janela_resultado)
        janela.title("Escolher condomínio")
        janela.geometry("520x520")
        janela.configure(fg_color=tema["fundo"])
        janela.transient(self._janela_resultado)
        janela.grab_set()

        escolha = {"condominio": None}

        ctk.CTkLabel(janela, text="De qual condomínio é este protocolo?",
                     font=(fonte, 15), text_color=tema["texto"], anchor="w").pack(
            fill="x", padx=24, pady=(24, 4))
        ctk.CTkLabel(janela, text=dados.get("arquivo", ""), font=(fonte, 12),
                     text_color=tema["texto_secundario"], anchor="w").pack(
            fill="x", padx=24)

        entrada = ctk.CTkEntry(janela, corner_radius=0,
                               placeholder_text="Buscar por nome ou código",
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(fill="x", padx=24, pady=(12, 8))
        entrada.focus_set()

        area = ctk.CTkScrollableFrame(janela, corner_radius=0,
                                      fg_color=tema["fundo"])
        area.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        def confirmar(condominio):
            escolha["condominio"] = condominio
            janela.destroy()

        def redesenhar(_evento=None):
            for widget in area.winfo_children():
                widget.destroy()
            achados = buscar_condominios(entrada.get(), self.cadastro, limite=60)
            if not achados:
                ctk.CTkLabel(area, text="Nenhum condomínio com esse nome ou código.",
                             font=(fonte, 13), text_color=tema["texto_secundario"],
                             anchor="w").pack(fill="x", pady=8)
                return
            for condominio in achados:
                #  Sem ID SL aparece, mas marcado: escolher esse não destrava,
                #  e a pessoa precisa saber disso antes de clicar.
                sufixo = "" if condominio["id_sl"] else "   (sem ID do Superlógica)"
                cor = tema["texto"] if condominio["id_sl"] else tema["texto_terciario"]
                ctk.CTkButton(
                    area,
                    text=f"{condominio['codigo']}  {condominio['nome']}{sufixo}",
                    corner_radius=0, anchor="w", height=30,
                    fg_color="transparent", hover_color=tema["superficie"],
                    text_color=cor, font=(fonte, 13),
                    command=lambda c=condominio: confirmar(c),
                ).pack(fill="x")

        entrada.bind("<KeyRelease>", redesenhar)
        redesenhar()

        ctk.CTkButton(janela, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(padx=24, pady=(0, 24), anchor="e")

        self.wait_window(janela)
        return escolha["condominio"]

    def _pedir_vencimento_despesas(self):
        """
        Janelinha que pede o vencimento do lote de despesas. Devolve datetime,
        ou None se o usuário cancelar.

        A data é do lote, não do modelo: perguntar aqui é o que impede alguém
        de digitá-la na célula do modelo, onde ela virava número e fazia o
        Superlógica gravar 01/01/1970 e recusar os lançamentos.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        #  Esta janelinha é aberta em DOIS momentos: no início do lote (junto
        #  da tarifa, para o vencimento entrar no carimbo) e, como reserva, a
        #  partir do painel de resultado. No primeiro, o painel ainda não
        #  existe — daí o pai ser resolvido em vez de fixo.
        pai = self._janela_para_dialogo()
        janela = ctk.CTkToplevel(pai)
        janela.title("Vencimento")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(pai)
        janela.grab_set()

        escolha = {"data": None}

        ctk.CTkLabel(janela, text="Qual o vencimento destas despesas?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text="Ex.: 21/08/2026", font=(fonte, 12),
                     text_color=tema["texto_terciario"]).pack(padx=24, anchor="w")

        entrada = ctk.CTkEntry(janela, corner_radius=0, width=200,
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(padx=24, pady=(12, 4), anchor="w")
        entrada.focus_set()

        aviso = ctk.CTkLabel(janela, text="", font=(fonte, 12),
                             text_color=tema["acento"])
        aviso.pack(padx=24, pady=(0, 8), anchor="w")

        def confirmar():
            data, mensagem = converter_data_digitada(entrada.get())
            if data is None:
                aviso.configure(text=mensagem)
                return
            escolha["data"] = data
            janela.destroy()

        botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        botoes.pack(padx=24, pady=(0, 24), anchor="e")
        ctk.CTkButton(botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(botoes, text="Continuar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return escolha["data"]

    def _pedir_chave_despesas(self):
        """
        Pede a chave da importação de despesas. Devolve int, ou None se o
        usuário cancelar.

        Existe para que ninguém precise abrir a planilha gerada no Excel só
        para trocar esse número: foi abrindo o arquivo para isso que uma vez a
        coluna `vencimento` perdeu o formato de data e o Superlógica gravou
        01/01/1970 nos lançamentos, sem acusar erro nenhum.

        Campo vazio de propósito, sem valor padrão — pelo mesmo motivo da
        tarifa: a chave muda a cada importação, e um padrão herdado passaria
        batido.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        pai = self._janela_para_dialogo()
        janela = ctk.CTkToplevel(pai)
        janela.title("Chave da importação")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(pai)
        janela.grab_set()

        escolha = {"chave": None}

        ctk.CTkLabel(janela, text="Qual a chave desta importação?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text="Só números. Ex.: 46",
                     font=(fonte, 12), text_color=tema["texto_terciario"]).pack(
            padx=24, anchor="w")

        entrada = ctk.CTkEntry(janela, corner_radius=0, width=200,
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(padx=24, pady=(12, 4), anchor="w")
        entrada.focus_set()

        aviso = ctk.CTkLabel(janela, text="", font=(fonte, 12),
                             text_color=tema["acento"])
        aviso.pack(padx=24, pady=(0, 8), anchor="w")

        def confirmar():
            texto = entrada.get().strip()
            if not texto.isdigit():
                aviso.configure(text="Informe só números, sem pontos nem letras.")
                return
            escolha["chave"] = int(texto)
            janela.destroy()

        botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        botoes.pack(padx=24, pady=(0, 24), anchor="e")
        ctk.CTkButton(botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(botoes, text="Continuar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return escolha["chave"]

    def _pedir_valor_protocolo(self, dados):
        """
        Janelinha que pede o valor em reais de um protocolo pendente. Devolve
        Decimal, ou None se o usuário cancelar. Mesma validação da tarifa,
        pela mesma função.
        """
        tema = self.tema_atual
        fonte = familia_fonte()
        janela = ctk.CTkToplevel(self._janela_resultado)
        janela.title("Informar valor")
        janela.configure(fg_color=tema["fundo"])
        janela.resizable(False, False)
        janela.transient(self._janela_resultado)
        janela.grab_set()

        escolha = {"valor": None}

        ctk.CTkLabel(janela, text="Qual o valor deste protocolo?",
                     font=(fonte, 15), text_color=tema["texto"]).pack(
            padx=24, pady=(24, 4), anchor="w")
        ctk.CTkLabel(janela, text=dados.get("arquivo", ""), font=(fonte, 12),
                     text_color=tema["texto_secundario"]).pack(padx=24, anchor="w")
        ctk.CTkLabel(janela, text=dados.get("condominio", ""), font=(fonte, 12),
                     text_color=tema["texto_terciario"]).pack(padx=24, anchor="w")

        entrada = ctk.CTkEntry(janela, corner_radius=0, width=200,
                               fg_color=tema["superficie"], border_width=1,
                               border_color=tema["borda"], text_color=tema["texto"],
                               font=(fonte, 14))
        entrada.pack(padx=24, pady=(12, 4), anchor="w")
        entrada.focus_set()

        aviso = ctk.CTkLabel(janela, text="", font=(fonte, 12),
                             text_color=tema["acento"])
        aviso.pack(padx=24, pady=(0, 8), anchor="w")

        def confirmar():
            valor, mensagem = converter_valor_digitado(entrada.get())
            if valor is None:
                aviso.configure(text=mensagem)
                return
            escolha["valor"] = valor
            janela.destroy()

        botoes = ctk.CTkFrame(janela, corner_radius=0, fg_color=tema["fundo"])
        botoes.pack(padx=24, pady=(0, 24), anchor="e")
        ctk.CTkButton(botoes, text="Cancelar", corner_radius=0, width=100,
                      fg_color="transparent", hover_color=tema["superficie"],
                      border_width=1, border_color=tema["borda_forte"],
                      text_color=tema["texto"], font=(fonte, 13),
                      command=janela.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(botoes, text="Confirmar", corner_radius=0, width=120,
                      fg_color=tema["acento"], hover_color=tema["acento_hover"],
                      text_color=tema["sobre_acento"], border_width=0,
                      font=(fonte, 13), command=confirmar).pack(side="left")

        entrada.bind("<Return>", lambda _e: confirmar())
        self.wait_window(janela)
        return escolha["valor"]

    def _acao_informar_valor(self, dados):
        """
        Pede o valor em reais e resolve a pendência: carimba o PDF, move a
        linha para os calculados e regrava a planilha.
        """
        if not dados:
            return
        ctx = getattr(self, "_ctx_protocolos", None)
        if not ctx:
            messagebox.showerror(
                "Erro",
                "O contexto do processamento se perdeu. Rode o lote de novo.",
                parent=self._janela_resultado)
            return

        valor = self._pedir_valor_protocolo(dados)
        if valor is None:
            return

        #  Toda a decisão de pasta/lote, o carimbo (via callback pra não
        #  virar dependência de PDF de verdade em `logica.py`) e a
        #  atualização de `ctx` (linha da planilha e contador de carimbados)
        #  vivem em `resolver_protocolo_manual` — extraído pra ficar
        #  testável sem instanciar a janela. Se o carimbo falhar, a exceção
        #  sobe daqui sem `ctx` ter sido tocado (não faz sentido marcar como
        #  resolvido um arquivo que não saiu carimbado).
        def carimbar(pasta_destino, registro_dados):
            self._carimbar_protocolo(
                dados["caminho"], pasta_destino, dados["arquivo"],
                registro_dados, None, None, ctx["config"], ctx["codigos"],
                valor_manual=valor, vencimento=ctx.get("vencimento"))

        try:
            linha_atualizada, motivo_painel, _pasta_destino = resolver_protocolo_manual(
                ctx, dados, valor, self.cadastro, carimbar)
        except Exception as e:
            messagebox.showerror(
                "Erro ao carimbar", f"Não foi possível carimbar o PDF:\n{e}",
                parent=self._janela_resultado)
            return

        #  Move do painel de pendentes para os calculados. O motivo que
        #  aparece aqui é o MESMO texto que acabou de ir para a planilha
        #  (última coluna de `linha_atualizada`) — inclui, por exemplo,
        #  "Código não cadastrado" quando for o caso, em vez de deixar a
        #  linha do painel parecer resolvida sem ressalva nenhuma.
        resultado = self._resultado_protocolos
        resultado["pendentes"] = [p for p in resultado["pendentes"] if p is not dados]
        #  O registro é montado em `logica.py` porque ele precisa carregar
        #  adiante `caminho`/`indice_linha`/`pasta_destino`: este item pode
        #  VOLTAR para a tabela de pendentes (todo processado sem ID SL
        #  aparece lá, ver `_listas_do_painel_protocolos`), e "Escolher
        #  condomínio" lê justamente esses campos.
        registro_processado = registro_painel_resolvido(
            dados, valor, motivo_painel, _pasta_destino)
        resultado["processados"].append(registro_processado)
        resultado["total_valor"] = resultado.get("total_valor", Decimal("0.00")) + valor

        #  2. Regrava a planilha inteira. Se falhar (tipicamente porque está
        #     aberta no Excel), o valor NÃO se perde: já está em ctx["linhas"]
        #     e no resultado, e vai junto na próxima regravação.
        try:
            salvar_planilha_protocolo(ctx["destino"], ctx["linhas"])
        except Exception as e:
            messagebox.showwarning(
                "Planilha não atualizada",
                f"O valor foi carimbado no PDF, mas a planilha não pôde ser "
                f"regravada:\n{e}\n\nSe ela estiver aberta no Excel, feche e "
                f"informe o próximo valor — a planilha é regravada inteira a "
                f"cada correção.",
                parent=self._janela_resultado)

        self._atualizar_painel_protocolos()

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
        """Grava as linhas da sessão de processamento no arquivo .log e, se ele
        tiver passado do limite de tamanho, descarta as sessões mais antigas
        (ver rotacionar_log)."""
        try:
            with open(self.caminho_log, "a", encoding="utf-8") as f:
                f.write("\n".join(linhas_sessao) + "\n\n")
            rotacionar_log(self.caminho_log)
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
        separar_em_lotes = self.separar_em_lotes.get()
        tamanho_lote = self._tamanho_lote_valido()

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
                codigo_protocolo = None

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
                            candidatos_regiao = extrair_cnpj_tomador(texto_regiao, cnpj_emitente_norm, self.cadastro)
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

                    # 2a.1) "Protocolo de Recebimento de Documento" dos Correios
                    #       (Imodata): sem CNPJ nenhum, mas com o código do
                    #       condomínio já pronto no texto — tenta antes do
                    #       caminho de CNPJ, que não acharia nada nesse formato.
                    codigo_protocolo = extrair_codigo_protocolo_correio(texto)
                    if codigo_protocolo is not None:
                        cnpj_do_protocolo = _codigos_do_cadastro(self.cadastro).get(codigo_protocolo)
                        if cnpj_do_protocolo:
                            candidatos = [cnpj_do_protocolo]
                            sufixo_origem = " (via código do protocolo dos Correios)"
                            origem_humana = "pelo código do protocolo dos Correios"
                        else:
                            candidatos = []
                    else:
                        candidatos = extrair_cnpj_tomador(texto, cnpj_emitente_norm, self.cadastro)

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
                                candidatos_retry = extrair_cnpj_tomador(texto_retry, cnpj_emitente_norm, self.cadastro)
                                if candidatos_retry:
                                    texto = texto_retry
                                    candidatos = candidatos_retry
                                    sufixo_origem = f" (via OCR, re-tentativa DPI {proximo})"
                                    origem_humana = "pelo CNPJ (releitura)"

                if len(candidatos) == 0 and codigo_protocolo is not None:
                    detalhe = f"Código {codigo_protocolo} (protocolo dos Correios) não cadastrado"
                    pendentes.append((nome, detalhe))
                    msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — {detalhe}"
                    res_pendentes.append({
                        "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "nao_cadastrado",
                        "motivo": "Código não cadastrado", "cnpj": None, "nome_sugerido": None,
                        "candidatos": None})
                elif len(candidatos) == 0:
                    candidatos_nome = candidatos_por_nome(nome, texto, self.cadastro)
                    if candidatos_nome:
                        pendentes.append((nome, "Nome parecido encontrado" + sufixo_origem))
                        msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — nome parecido encontrado{sufixo_origem}"
                        res_pendentes.append({
                            "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "ambiguo",
                            "motivo": "Nome parecido encontrado", "cnpj": None, "nome_sugerido": None,
                            "candidatos": candidatos_nome})
                    else:
                        pendentes.append((nome, "CNPJ do tomador não encontrado no PDF" + sufixo_origem))
                        msg = f"[{idx}/{len(arquivos)}] ⚠ {nome} — CNPJ não encontrado{sufixo_origem}"
                        res_pendentes.append({
                            "arquivo": nome, "caminho": caminho_entrada_pdf, "tipo": "nao_lido",
                            "motivo": "Não foi possível ler", "cnpj": None, "nome_sugerido": None,
                            "candidatos": None})
                elif len(candidatos) > 1:
                    lista = ", ".join(formatar_documento(c) for c in candidatos)
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
                        detalhe = f"CNPJ {formatar_documento(cnpj)} não cadastrado"
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
                        if codigo_protocolo is not None:
                            # Protocolo dos Correios: carimbo lateral rotacionado,
                            # automático — não usa modo_texto/tipo_servico.
                            texto_pdf = montar_texto_protocolo_correio(codigo, condominio, cnpj)
                            config_arquivo = {**config, "angulo": 90}
                        elif modo == "rodape":
                            texto_pdf = f"{codigo} {condominio} - {tipo_servico}".strip()
                            config_arquivo = config
                        else:
                            texto_pdf = codigo
                            config_arquivo = config
                        nome_saida = nome_saida_com_codigo(nome, codigo) if renomear else nome
                        # `sucesso` ainda não foi incrementado: é o índice
                        # 0-based deste arquivo entre os codificados, então os
                        # lotes saem cheios (pendentes não ocupam vaga).
                        pasta_destino = (caminho_do_lote(saida, sucesso, tamanho_lote)
                                          if separar_em_lotes else saida)
                        caminho_saida_pdf = os.path.join(pasta_destino, nome_saida)
                        processar_pdf(caminho_entrada_pdf, caminho_saida_pdf, texto_pdf, config_arquivo)
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
