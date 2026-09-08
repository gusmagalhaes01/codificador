# -*- mode: python ; coding: utf-8 -*-
"""
Receita de build do Codificador (PyInstaller).

Uso:  pyinstaller --noconfirm codificador.spec
      (ou duplo clique em gerar_exe.bat, que confere o ambiente, roda os
      testes e monta o CODIFICADOR.zip de distribuição)

Gera dist/Codificador/ — uma PASTA com o .exe e o _internal/, que é
distribuída zipada como CODIFICADOR.zip. Modo --onedir de propósito, não
--onefile: o --onefile extrairia ~40 MB no temp a cada abertura (app de
uso diário, abertura lenta incomoda) e é mais suscetível a alarme falso de
antivírus/SmartScreen. O preço é o usuário ter de manter a pasta inteira
junta, o que o zip resolve.

pasta_base() usa sys.executable quando frozen, então config.json, os logs
e o cadastro_condominios.xlsx ficam AO LADO do .exe, dentro dessa pasta —
por isso a planilha vai junto no zip (ver gerar_exe.bat).

Pontos que já quebraram o exe antes e por isso estão explícitos aqui:

- icone.ico precisa ir em `datas` na raiz do pacote, porque
  caminho_recurso() o procura em sys._MEIPASS. Passar só `icon=` embute o
  ícone no executável, mas NÃO o disponibiliza em tempo de execução —
  _aplicar_icone() não acharia o arquivo.
- customtkinter carrega temas e fontes de arquivos .json/.otf dentro do
  próprio pacote; sem copiar a pasta inteira, o app abre e quebra na
  primeira janela.
- winocr importa os bindings winrt.* dinamicamente. Sem os hiddenimports
  abaixo o PyInstaller não os enxerga no grafo de imports, e o exe sai com
  OCR_DISPONIVEL=False — a leitura de boletos escaneados morre em silêncio,
  sem nenhum erro visível.
- os dados do Tcl/Tk (as pastas _tcl_data e _tk_data dentro do _internal).
  Normalmente quem os copia é o hook de tkinter do próprio PyInstaller, mas
  já saiu build sem eles, e aí o exe morre logo na abertura com
  `FileNotFoundError: Tcl data directory ... _tcl_data not found` — antes de
  qualquer janela, então nem o handler global de erros pega. A rede está
  logo abaixo de Analysis: se o hook não trouxe, a receita copia à mão.
"""

import os

import customtkinter

pasta_customtkinter = os.path.dirname(customtkinter.__file__)

datas = [
    ("icone.ico", "."),                       # lido via caminho_recurso/sys._MEIPASS
    (pasta_customtkinter, "customtkinter"),   # temas .json e fontes .otf
]

hiddenimports = [
    "darkdetect",
    # Bindings nativos usados pelo winocr (motor de OCR do Windows)
    "winocr",
    "winrt.windows.media.ocr",
    "winrt.windows.graphics.imaging",
    "winrt.windows.storage.streams",
    "winrt.windows.globalization",
    "winrt.windows.foundation",
    "winrt.windows.foundation.collections",
]

a = Analysis(
    ["identificacao_por_cnpj_6_0.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Reduz o tamanho do exe: nada aqui é usado pelo app.
    #  lxml fica de fora de propósito. Ele não é dependência do projeto (não
    #  está no requirements.txt), mas se estiver instalado por acaso na
    #  máquina do build o PyInstaller o empacota — e aí o openpyxl passa a
    #  ESCREVER o XML das planilhas por ele, em vez do parser da biblioteca
    #  padrão. Todas as releases até a v6.13.0 saíram sem lxml; trocar o
    #  gerador de XML sem querer, num programa cujo entregável é uma planilha
    #  lida por outro sistema, é risco sem contrapartida. Também economiza
    #  ~4 MB no zip.
    excludes=["tkinter.test", "test", "unittest", "pydoc_data", "lxml"],
    noarchive=False,
    optimize=0,
)

#  ---- Rede: dados do Tcl/Tk ----------------------------------------------
#  Sem a pasta _tcl_data dentro do _internal, o exe nem chega a abrir a
#  janela: o runtime hook do tkinter levanta FileNotFoundError e o programa
#  morre antes de qualquer tratamento de erro nosso. O hook do PyInstaller
#  costuma copiar essas pastas sozinho, mas já falhou em máquina de build
#  real (ver o cabeçalho), e o build "passa" sem aviso nenhum — o problema
#  só aparece no primeiro duplo clique do usuário. Por isso: se os dados não
#  estiverem no pacote, procura-se o Tcl/Tk pelo próprio interpretador que
#  está rodando o build e copia-se na mão.


def _tem_dados(analysis, destino):
    """A pasta `destino` já foi para dentro do pacote?"""
    return any(nome.replace("\\", "/").startswith(destino + "/")
               for nome, _, _ in analysis.datas)


def _dados_do_tcl_tk():
    """Pastas de dados do Tcl e do Tk da instalação do Python do build.

    `info library` é o próprio Tcl dizendo onde ficam seus arquivos, o que
    funciona igual em Python do python.org, da Microsoft Store ou embutido —
    palpitar o caminho a partir de sys.base_prefix, não.
    """
    import tkinter

    biblioteca_tcl = tkinter.Tcl().eval("info library")     # .../lib/tcl8.6
    pasta_lib = os.path.dirname(biblioteca_tcl)
    versao = os.path.basename(biblioteca_tcl).replace("tcl", "")
    biblioteca_tk = os.path.join(pasta_lib, "tk" + versao)  # .../lib/tk8.6
    return biblioteca_tcl, biblioteca_tk


if not _tem_dados(a, "_tcl_data") or not _tem_dados(a, "_tk_data"):
    pasta_tcl, pasta_tk = _dados_do_tcl_tk()
    if not _tem_dados(a, "_tcl_data"):
        a.datas += Tree(pasta_tcl, prefix="_tcl_data")
    if not _tem_dados(a, "_tk_data"):
        a.datas += Tree(pasta_tk, prefix="_tk_data")
    print("[codificador.spec] dados do Tcl/Tk copiados à mão "
          "(o hook do tkinter não os trouxe)")


pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # binários vão para o _internal/ via COLLECT
    name="Codificador",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # app gráfico: sem janela de console atrás
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="icone.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Codificador",     # resulta em dist/Codificador/
)
