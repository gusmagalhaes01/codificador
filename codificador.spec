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
