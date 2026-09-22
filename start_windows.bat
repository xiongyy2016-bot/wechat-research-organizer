@echo off
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
start "" http://127.0.0.1:8765
.venv\Scripts\python.exe app.py
pause
