@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m unittest discover -s tests -p "test_*.py" -v
echo.
pause
