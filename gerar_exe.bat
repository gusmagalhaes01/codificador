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
echo [3/3] Empacotando...
python -m PyInstaller --noconfirm codificador.spec
if errorlevel 1 goto erro

echo.
echo ============================================
echo   Pronto: dist\Codificador.exe
echo ============================================
echo.
echo Lembre-se: cadastro_condominios.xlsx precisa ficar
echo NA MESMA PASTA do .exe (o app grava config.json e os
echo logs ao lado dele).
echo.
pause
exit /b 0

:erro
echo.
echo *** Falhou. Veja as mensagens acima. ***
pause
exit /b 1
