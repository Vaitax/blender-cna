@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -c "import keyring" 2>nul || %PY% -m pip install --user keyring
%PY% guardar_clave.py
pause
