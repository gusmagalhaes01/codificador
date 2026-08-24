@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Gerando Codificador.exe
echo ============================================
echo.

echo [1/3] Conferindo dependencias...
python -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 goto erro

echo.
echo [2/3] Rodando os testes antes de empacotar...
python -m unittest discover -s tests -p "test_*.py"
if errorlevel 1 (
    echo.
    echo *** TESTES FALHARAM - build cancelado. ***
    goto erro
)

echo.
echo [3/4] Empacotando...
python -m PyInstaller --noconfirm codificador.spec
if errorlevel 1 goto erro

echo.
echo [4/4] Montando o CODIFICADOR.zip...
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
