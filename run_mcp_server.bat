@echo off
REM MCP Server batch file for Windows
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "venv" (
    echo Virtual environment not found. Creating it...
    python -m venv venv
    call venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)

echo Starting PPT MCP Server...
python ppt_mcp_server.py