@echo off
cd /d "%~dp0"

echo ========================================
echo INSTALADOR AUTOMATICO - GERADOR DE PONTO
echo ========================================
echo.

REM Verifica se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado!
    echo Baixando e instalando Python automaticamente...
    echo.
    
    REM Cria pasta temporária
    if not exist "temp_install" mkdir temp_install
    
    REM Baixa o instalador do Python (versão 3.12.0 - 64 bits)
    echo Baixando Python 3.12.0...
    echo Isso pode levar alguns minutos dependendo da internet...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe' -OutFile 'temp_install\python_installer.exe'}"
    
    if not exist "temp_install\python_installer.exe" (
        echo.
        echo ERRO: Nao foi possivel baixar o Python.
        echo Verifique sua conexao com a internet.
        echo.
        pause
        exit /b 1
    )
    
    echo.
    echo Instalando Python...
    echo IMPORTANTE: A instalacao sera automatica e silenciosa.
    echo.
    
    REM Instala Python silenciosamente com Add to PATH
    temp_install\python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    
    REM Aguarda instalação finalizar
    timeout /t 10 /nobreak >nul
    
    REM Limpa arquivos temporários
    rd /s /q temp_install
    
    echo.
    echo Python instalado com sucesso!
    echo.
    echo IMPORTANTE: Feche esta janela e execute o arquivo novamente.
    pause
    exit /b 0
)

echo Python encontrado!
python --version
echo.

REM Verifica se a venv existe, se não, cria automaticamente
if not exist ".venv\Scripts\activate.bat" (
    echo ========================================
    echo Criando ambiente virtual...
    echo ========================================
    python -m venv .venv
    echo.
    echo Instalando bibliotecas necessarias...
    echo Isso pode levar alguns minutos...
    echo ========================================
    call .venv\Scripts\activate.bat
    pip install pandas tkcalendar reportlab openpyxl
    echo.
    echo ========================================
    echo Instalacao concluida com sucesso!
    echo ========================================
    echo.
    echo Agora use o arquivo ABRIR_PROGRAMA.bat
    echo para abrir o programa.
    echo ========================================
) else (
    echo ========================================
    echo Sistema ja esta instalado!
    echo ========================================
    echo.
    echo Use o arquivo ABRIR_PROGRAMA.bat
    echo para abrir o programa.
    echo ========================================
)

echo.
pause
