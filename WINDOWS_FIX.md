# Windows JSON Error Fix

## Problem: "Unexpected token 'S', 'Starting P'... is not valid JSON"

This error occurs because the batch file or server is outputting text to stdout, which interferes with the JSON communication protocol that MCP requires.

## Solution

Use the basic MCP server version that doesn't output any banners or text to stdout.

### Step 1: Use the Fixed Configuration

**Config Location**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "C:\\Users\\YourName\\ppt-mcp-server\\venv\\Scripts\\python.exe",
      "args": ["C:\\Users\\YourName\\ppt-mcp-server\\ppt_mcp_server_basic.py"],
      "env": {}
    }
  }
}
```

### Step 2: Alternative with Batch File

If you prefer using a batch file:

```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "C:\\Users\\YourName\\ppt-mcp-server\\run_mcp_server_basic.bat",
      "args": [],
      "env": {}
    }
  }
}
```

### Step 3: Test the Basic Server

Test manually first:
```cmd
cd C:\Users\YourName\ppt-mcp-server
venv\Scripts\activate
python ppt_mcp_server_basic.py
```

This should NOT show any startup banner or messages - it should just wait silently for input.

### What Was Fixed

1. **Removed echo statements** from batch files that were outputting to stdout
2. **Created basic MCP server** (`ppt_mcp_server_basic.py`) without FastMCP banner
3. **Redirected all logging to stderr** instead of stdout
4. **Suppressed installation messages** using `>nul 2>&1`

### Files for Windows Fix

- `ppt_mcp_server_basic.py` - Clean server without banners
- `run_mcp_server_basic.bat` - Silent batch file
- `claude_desktop_config_windows_fixed.json` - Fixed configuration

### Important Notes

- **Use absolute paths** (no ~/ shortcuts)
- **Use double backslashes** in JSON paths (`\\`)
- **Replace** `C:\\Users\\YourName\\ppt-mcp-server` with your actual path
- **Restart Claude Desktop** completely after config changes

The basic server provides the same functionality as the FastMCP version but without any stdout pollution that causes JSON parsing errors.