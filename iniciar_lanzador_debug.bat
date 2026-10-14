@echo off
title Lanzador (Modo Consola / Debug)
cd /d "%~dp0"

set "PYTHON_EXE=..\.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" goto RUN

set "PYTHON_EXE=.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" goto RUN

set "PYTHON_EXE=python.exe"

:RUN
"%PYTHON_EXE%" "Lanzador.py"
pause
