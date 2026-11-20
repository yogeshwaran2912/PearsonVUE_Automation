@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist main.py (python main.py reviews %*) else (python main.pyc reviews %*)
pause
