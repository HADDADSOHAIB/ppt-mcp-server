#!/bin/bash

# MCP Server wrapper script that ensures proper environment setup
cd "/Users/haddad/ppt-mcp-server"

# Activate virtual environment
source venv/bin/activate

# Run the MCP server
exec python ppt_mcp_server.py