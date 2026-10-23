@echo off
cd /d "%~dp0"
py -3.13 -m pip install --quiet -r requirements.txt
py -3.13 monitor_laudo.py
if errorlevel 1 python monitor_laudo.py
pause
