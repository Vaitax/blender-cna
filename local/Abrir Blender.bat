@echo off
chcp 65001 >nul
title Blender CNA
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -c "import pyodbc, keyring, openpyxl" 2>nul || (
  echo Instalando componentes de Python por unica vez...
  %PY% -m pip install --user pyodbc keyring openpyxl
)
%PY% blender_local.py
pause
