#!/bin/bash

# Activate virtual environment and start the MCP server
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    echo "Virtual environment not found. Creating it..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

echo "Starting PPT MCP Server..."
python ppt_mcp_server.py