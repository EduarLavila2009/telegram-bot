@echo off
setlocal
title Bot financiero Telegram

rem Subir a bot_financiero_telegram (donde esta .venv)
cd /d "%~dp0.." || (
    echo No se pudo entrar a la carpeta del bot.
    pause
    exit /b 1
)
set "PY=%CD%\.venv\Scripts\python.exe"

rem Raiz del repo (padre de bot_financiero_telegram)
cd /d ".." || (
    echo No se pudo entrar a la raiz del repositorio.
    pause
    exit /b 1
)
if not exist "%PY%" (
    echo Falta el entorno virtual. Una sola vez, en PowerShell:
    echo   cd "%BOT_DIR%"
    echo   python -m venv .venv
    echo   .\.venv\Scripts\activate
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

"%PY%" -m bot_financiero_telegram
if errorlevel 1 pause

endlocal
