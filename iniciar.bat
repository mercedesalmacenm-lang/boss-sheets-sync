@echo off
title BOSS Sheets Sync
color 0F

echo ============================================
echo    BOSS Sheets Sync - Monitor de Inventario
echo ============================================
echo.

:: Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se encontro Python instalado.
    echo.
    echo Descargalo desde: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)

echo [OK] Python detectado.
echo.

:: Verificar credenciales
if not exist "%~dp0credentials.json" (
    echo [ERROR] No se encontro credentials.json en esta carpeta.
    echo.
    echo Copia el archivo credentials.json junto a este .bat
    echo.
    pause
    exit /b 1
)

echo [OK] Credenciales detectadas.
echo.

:: Instalar dependencias si es necesario
if not exist "%~dp0.venv" (
    echo [INFO] Primera ejecucion - instalando dependencias...
    python -m venv "%~dp0.venv"
    call "%~dp0.venv\Scripts\activate.bat"
    pip install -r "%~dp0requirements.txt" --quiet
    echo [OK] Dependencias instaladas.
    echo.
) else (
    call "%~dp0.venv\Scripts\activate.bat"
)

:: Ejecutar monitor
echo [INFO] Iniciando monitor...
echo [INFO] Los archivos .xlsx del Escritorio se subiran automaticamente.
echo [INFO] Presiona Ctrl+C para detener.
echo.
python "%~dp0monitor.py"

pause
