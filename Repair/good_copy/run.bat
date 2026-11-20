@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist main.py (python main.py pvue %*) else (python main.pyc pvue %*)
pause
