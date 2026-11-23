@echo off
cd /d "%~dp0"
chcp 65001 >nul
title WFM
set PY=python
where py >nul 2>nul && set PY=py
%PY% py\wfm_menu.py
if errorlevel 1 pause
