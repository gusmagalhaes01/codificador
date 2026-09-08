@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Gerando Codificador.exe
echo ============================================
echo.

echo [1/5] Conferindo dependencias...
python -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 goto erro

echo.
echo [2/5] Rodando os testes antes de empacotar...
python -m unittest discover -s tests -p "test_*.py"
if errorlevel 1 (
    echo.
    echo *** TESTES FALHARAM - build cancelado. ***
    goto erro
)

echo.
echo [3/5] Empacotando...
rem --clean descarta o cache do build anterior. Sem isso, um build feito com
rem outra versao do PyInstaller (ou interrompido no meio) e reaproveitado, e
rem ja saiu pacote faltando arquivo por causa disso.
python -m PyInstaller --noconfirm --clean codificador.spec
if errorlevel 1 goto erro

echo.
echo [4/5] Conferindo o que foi empacotado...
rem Sem os dados do Tcl/Tk o exe morre no duplo clique, ANTES de abrir a
rem janela ("Tcl data directory ... _tcl_data not found") - e o build passa
rem sem aviso nenhum. Melhor falhar aqui, com o motivo na tela, do que
rem entregar um zip que nao abre.
if not exist "dist\Codificador\Codificador.exe" (
    echo *** O executavel nao foi gerado. ***
    goto erro
)
if not exist "dist\Codificador\_internal\_tcl_data" (
    echo *** Faltou _internal\_tcl_data - o exe nao abriria. ***
    goto erro
)
if not exist "dist\Codificador\_internal\_tk_data" (
    echo *** Faltou _internal\_tk_data - o exe nao abriria. ***
    goto erro
)
echo OK: executavel e dados do Tcl/Tk no lugar.

echo.
echo [5/5] Montando o CODIFICADOR.zip...
rem A planilha vai DENTRO da pasta do exe: pasta_base() resolve para o
rem diretorio do executavel, entao e ali que o app procura o cadastro e
rem grava config.json e os logs.
copy /Y cadastro_condominios.xlsx "dist\Codificador\" >nul
if errorlevel 1 goto erro
copy /Y modelo_despesas.xlsx "dist\Codificador\" >nul
if errorlevel 1 goto erro
if exist "dist\CODIFICADOR.zip" del /Q "dist\CODIFICADOR.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\Codificador' -DestinationPath 'dist\CODIFICADOR.zip' -Force"
if errorlevel 1 goto erro

echo.
echo ============================================
echo   Pronto: dist\CODIFICADOR.zip
echo ============================================
echo.
echo O zip contem a pasta Codificador\ com o executavel,
echo o _internal\, o cadastro_condominios.xlsx e o modelo_despesas.xlsx.
echo Extrair a pasta INTEIRA - o exe nao roda sozinho.
echo.
pause
exit /b 0

:erro
echo.
echo *** Falhou. Veja as mensagens acima. ***
pause
exit /b 1
