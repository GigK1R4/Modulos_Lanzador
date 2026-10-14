@echo off
title Lanzador de Automatizacion de Archivos
cd /d "%~dp0"

:: Si existe el ejecutable compilado, iniciarlo directamente
if exist "Lanzador.exe" (
    start "" "Lanzador.exe"
    exit
)

:: Alternativa: Buscar interprete de Python en el entorno virtual
set "PYTHON_EXE=..\.venv\Scripts\pythonw.exe"
if exist "%PYTHON_EXE%" goto RUN

set "PYTHON_EXE=.venv\Scripts\pythonw.exe"
if exist "%PYTHON_EXE%" goto RUN

set "PYTHON_EXE=pythonw.exe"

:RUN
start "" "%PYTHON_EXE%" "Lanzador.py"
exit
