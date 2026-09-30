@echo off
set "PROJECT_DIR=%~dp0"
set "PYTHON=%PROJECT_DIR%.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo Ambiente virtual nao encontrado. Crie-o com: python -m venv .venv
    pause
    exit /b 1
)

"%PYTHON%" -c "import flet, win32com.client" >nul 2>&1
if errorlevel 1 (
    echo Dependencias do aplicativo nao encontradas.
    echo Execute: .venv\Scripts\python.exe -m pip install -r requirements-app.txt
    pause
    exit /b 1
)

"%PYTHON%" "%PROJECT_DIR%app.py"
if errorlevel 1 pause