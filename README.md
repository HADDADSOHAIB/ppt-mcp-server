# PPT MCP Server

An MCP (Model Context Protocol) server for processing PowerPoint presentations and Word documents.

## Features

- **Extract PPT Information**: Extract content, structure, and metadata from PowerPoint files
- **Analyze Word Structure**: Analyze Word document structure to understand templates and organization
- **Intelligent Combination**: Combine PPT content with Word document structure
- **Generate Presentations**: Create new PowerPoint presentations from combined data

## Installation

1. Create a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Running the Server

```bash
python ppt_mcp_server.py
```

### Available Tools

1. **extract_ppt_info(file_path)**: Extract all content from a PowerPoint file
2. **analyze_word_structure(file_path)**: Analyze Word document structure
3. **combine_ppt_and_word(ppt_file, word_file, output_file)**: Combine PPT and Word content
4. **create_presentation_from_structure(structure_data, output_file)**: Generate PPT from structured data

### Configuration for Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "ppt-processor": {
      "command": "python",
      "args": ["/path/to/ppt-mcp-server/ppt_mcp_server.py"],
      "env": {}
    }
  }
}
```

## Example Workflow

1. Extract content from an existing PowerPoint presentation
2. Analyze the structure of a Word document template
3. Combine the extracted content with the Word structure
4. Generate a new PowerPoint presentation following the Word document's organization

## Dependencies

- fastmcp: MCP server framework
- python-pptx: PowerPoint file processing
- python-docx: Word document processing# ppt-mcp-server
