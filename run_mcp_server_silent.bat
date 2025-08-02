@echo off
setlocal
cd /d "%~dp0"

if not exist "venv" (
    python -m venv venv 2>nul
    call venv\Scripts\activate.bat 2>nul
    pip install -r requirements.txt 2>nul
) else (
    call venv\Scripts\activate.bat 2>nul
)

python ppt_mcp_server.py