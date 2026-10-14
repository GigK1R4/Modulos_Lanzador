@echo off
title Lanzador (Modo Consola / Debug)
cd /d "%~dp0"

set "PYTHON_EXE=..\.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" goto CHECK

set "PYTHON_EXE=.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" goto CHECK

set "PYTHON_EXE=python.exe"

:CHECK
"%PYTHON_EXE%" -c "import customtkinter, docxtpl, pypdf" 2>nul
if %errorlevel% neq 0 (
    echo ========================================================
    echo Configurando dependencias del Lanzador...
    echo ========================================================
    "%PYTHON_EXE%" -m pip install -r requirements.txt
)

:RUN
"%PYTHON_EXE%" "Lanzador.py"
pause
