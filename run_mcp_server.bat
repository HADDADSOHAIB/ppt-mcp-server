@echo off
REM MCP Server batch file for Windows
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "venv" (
    python -m venv venv >nul 2>&1
    call venv\Scripts\activate.bat >nul 2>&1
    pip install -r requirements.txt >nul 2>&1
) else (
    call venv\Scripts\activate.bat >nul 2>&1
)

REM Run the MCP server (no echo statements to avoid JSON parsing issues)
python ppt_mcp_server.py