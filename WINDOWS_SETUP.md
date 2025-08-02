# Windows Setup Guide for PPT MCP Server

## Quick Setup

### Step 1: Copy Project to Windows

Copy the entire `ppt-mcp-server` folder to your Windows machine, for example:
```
C:\Users\YourName\ppt-mcp-server\
```

### Step 2: Install Dependencies

Open Command Prompt or PowerShell and navigate to the project folder:
```cmd
cd C:\Users\YourName\ppt-mcp-server
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Step 3: Test the Server

Test that the server works:
```cmd
run_mcp_server.bat
```

You should see the FastMCP startup screen. Press Ctrl+C to stop.

### Step 4: Configure Claude Desktop

**Config File Location**: `%APPDATA%\Claude\claude_desktop_config.json`

**Option 1 - Using Batch File (Recommended)**:
```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "C:\\Users\\YourName\\ppt-mcp-server\\run_mcp_server.bat",
      "args": [],
      "env": {}
    }
  }
}
```

**Option 2 - Direct Python**:
```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "C:\\Users\\YourName\\ppt-mcp-server\\venv\\Scripts\\python.exe",
      "args": ["C:\\Users\\YourName\\ppt-mcp-server\\ppt_mcp_server.py"],
      "env": {}
    }
  }
}
```

### Step 5: Update Paths

**IMPORTANT**: Replace `C:\\Users\\YourName\\ppt-mcp-server` with your actual path.

To find your path:
```cmd
cd C:\path\to\your\ppt-mcp-server
echo %CD%
```

### Step 6: Restart Claude Desktop

- Close Claude Desktop completely
- Reopen Claude Desktop
- Open a new conversation and the MCP server should connect

## Windows-Specific Files Created

- `run_mcp_server.bat` - Windows batch startup script
- `claude_desktop_config_windows.json` - Windows config template
- `claude_desktop_config_windows_alternative.json` - Alternative config

## Available MCP Tools

Once connected, you'll have access to:
- `extract_ppt_info(file_path)` - Extract PowerPoint content  
- `analyze_word_structure(file_path)` - Analyze Word document structure
- `combine_ppt_and_word(ppt_file, word_file, output_file)` - Combine documents
- `create_presentation_from_structure(structure_data, output_file)` - Generate presentations

## Example Usage

```
User: "Extract information from C:\Documents\presentation.pptx"
Claude: [Uses extract_ppt_info tool]

User: "Combine C:\Documents\presentation.pptx with C:\Documents\structure.docx"
Claude: [Uses combine_ppt_and_word tool]
```

## Troubleshooting Windows Issues

### Common Problems:

1. **"'python' is not recognized"**
   - Install Python from python.org
   - Make sure Python is in your PATH

2. **"Access denied"**
   - Run Command Prompt as Administrator
   - Check folder permissions

3. **Path issues**
   - Use full absolute paths (C:\...)
   - Use double backslashes in JSON (\\)
   - No spaces in folder names (or use quotes)

4. **Virtual environment issues**
   - Delete `venv` folder and recreate:
     ```cmd
     rmdir /s venv
     python -m venv venv
     venv\Scripts\activate
     pip install -r requirements.txt
     ```

### PowerShell Alternative

If you prefer PowerShell, create `run_mcp_server.ps1`:
```powershell
Set-Location $PSScriptRoot
if (!(Test-Path "venv")) {
    python -m venv venv
    & "venv\Scripts\Activate.ps1"
    pip install -r requirements.txt
} else {
    & "venv\Scripts\Activate.ps1"
}
python ppt_mcp_server.py
```

Then use in config:
```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "powershell.exe",
      "args": ["-ExecutionPolicy", "Bypass", "-File", "C:\\path\\to\\ppt-mcp-server\\run_mcp_server.ps1"],
      "env": {}
    }
  }
}
```