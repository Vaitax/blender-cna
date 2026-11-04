@echo off
chcp 65001 >nul
title Blender CNA (con registro)
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% blender_local.py
pause
