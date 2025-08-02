# MCP Server Troubleshooting Guide

## Connection Issues

### Problem: "Failed to reconnect to ppt-processor"

This error typically occurs due to configuration issues. Here are the solutions:

### Solution 1: Use the Wrapper Script (Recommended)

1. **Copy the configuration**:
   ```json
   {
     "mcpServers": {
       "ppt-processor": {
         "command": "/Users/haddad/ppt-mcp-server/run_mcp_server.sh",
         "args": [],
         "env": {
           "PATH": "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"
         }
       }
     }
   }
   ```

2. **Add to Claude Desktop Config**:
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

3. **Restart Claude Desktop** completely (quit and reopen)

### Solution 2: Use Python Direct Path

If Solution 1 doesn't work, try this alternative:

```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "/Users/haddad/ppt-mcp-server/venv/bin/python",
      "args": ["/Users/haddad/ppt-mcp-server/ppt_mcp_server.py"],
      "env": {}
    }
  }
}
```

### Solution 3: Debug Steps

1. **Test server manually**:
   ```bash
   cd /Users/haddad/ppt-mcp-server
   ./run_mcp_server.sh
   ```
   Should show the FastMCP startup screen.

2. **Check file permissions**:
   ```bash
   ls -la run_mcp_server.sh
   # Should show: -rwxr-xr-x (executable)
   ```

3. **Verify Python environment**:
   ```bash
   /Users/haddad/ppt-mcp-server/venv/bin/python --version
   # Should show Python 3.x
   ```

### Solution 4: Fresh Installation

If all else fails:

1. **Recreate virtual environment**:
   ```bash
   cd /Users/haddad/ppt-mcp-server
   rm -rf venv
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Test again**:
   ```bash
   ./run_mcp_server.sh
   ```

## Common Error Messages

- **"command not found"**: Check the command path in your config
- **"Permission denied"**: Run `chmod +x run_mcp_server.sh`
- **"Module not found"**: Recreate virtual environment
- **"Connection refused"**: Restart Claude Desktop completely

## Available Tools

Once connected, you'll have access to these MCP tools:
- `extract_ppt_info` - Extract PowerPoint content
- `analyze_word_structure` - Analyze Word document structure  
- `combine_ppt_and_word` - Combine documents
- `create_presentation_from_structure` - Generate new presentations

## Support

If you still have issues:
1. Check the Claude Desktop logs
2. Verify all paths are absolute (no ~/ shortcuts)
3. Ensure Claude Desktop has file system permissions
4. Try restarting your computer