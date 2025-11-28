@echo off
cd /d "%~dp0"

REM Verifica se a venv existe
if not exist ".venv\Scripts\activate.bat" (
    echo ========================================
    echo ERRO: Ambiente nao configurado!
    echo ========================================
    echo.
    echo Execute primeiro o arquivo:
    echo INSTALADOR_COMPLETO.bat
    echo.
    echo Ele ira instalar o Python e configurar tudo.
    echo ========================================
    pause
    exit /b 1
)

REM Ativa o ambiente virtual e executa o programa
call .venv\Scripts\activate.bat
python gerador_ponto.py
