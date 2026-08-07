# -*- mode: python ; coding: utf-8 -*-
"""
Receita de build do Codificador (PyInstaller).

Uso:  pyinstaller --noconfirm codificador.spec
      (ou duplo clique em gerar_exe.bat, que também confere o ambiente)

Gera um único Codificador.exe em dist/. Escolhido --onefile por ser um
arquivo só para distribuir; pasta_base() usa sys.executable quando
frozen, então config.json, os logs e cadastro_condominios.xlsx continuam
ficando AO LADO do .exe (e não dentro do pacote), como o app espera.

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
    excludes=["tkinter.test", "test", "unittest", "pydoc_data"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Codificador",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # app gráfico: sem janela de console atrás
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="icone.ico",
)
