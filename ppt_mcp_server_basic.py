#!/usr/bin/env python3
"""
Basic MCP Server for PowerPoint processing without FastMCP banner issues.
"""

import asyncio
import json
import logging
import os
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Optional

# Configure logging to stderr only
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stderr
)
logger = logging.getLogger(__name__)

# Import the classes from the main server
from ppt_mcp_server import PPTExtractor, WordStructureAnalyzer, DocumentCombiner, PPTGenerator

class BasicMCPServer:
    """Basic MCP server implementation without banners"""
    
    def __init__(self, name: str):
        self.name = name
        self.tools = {}
    
    def tool(self, func):
        """Register a tool function"""
        self.tools[func.__name__] = func
        return func
    
    async def handle_request(self, request):
        """Handle MCP requests"""
        try:
            if request.get("method") == "tools/list":
                return {
                    "tools": [
                        {
                            "name": name,
                            "description": func.__doc__ or "",
                            "inputSchema": {
                                "type": "object",
                                "properties": {},
                                "required": []
                            }
                        }
                        for name, func in self.tools.items()
                    ]
                }
            
            elif request.get("method") == "tools/call":
                tool_name = request["params"]["name"]
                arguments = request["params"].get("arguments", {})
                
                if tool_name in self.tools:
                    result = await self.tools[tool_name](**arguments)
                    return {"content": [{"type": "text", "text": json.dumps(result, indent=2)}]}
                else:
                    return {"error": f"Unknown tool: {tool_name}"}
            
            elif request.get("method") == "initialize":
                return {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {
                        "tools": {}
                    },
                    "serverInfo": {
                        "name": self.name,
                        "version": "1.0.0"
                    }
                }
            
            return {"error": "Unknown method"}
            
        except Exception as e:
            logger.error(f"Error handling request: {e}")
            return {"error": str(e)}
    
    async def run_stdio(self):
        """Run server with stdio transport"""
        while True:
            try:
                line = await asyncio.get_event_loop().run_in_executor(None, sys.stdin.readline)
                if not line:
                    break
                
                request = json.loads(line.strip())
                response = await self.handle_request(request)
                
                # Send response
                print(json.dumps(response))
                sys.stdout.flush()
                
            except json.JSONDecodeError:
                continue
            except Exception as e:
                logger.error(f"Error in main loop: {e}")
                break

# Create server instance
server = BasicMCPServer("PPT Document Processor")

@server.tool
async def extract_ppt_info(file_path: str) -> Dict[str, Any]:
    """Extract information from a PowerPoint file."""
    try:
        if not os.path.exists(file_path):
            return {"error": f"File not found: {file_path}"}
        
        if not file_path.lower().endswith(('.pptx', '.ppt')):
            return {"error": "File must be a PowerPoint presentation (.pptx or .ppt)"}
        
        extractor = PPTExtractor()
        result = extractor.extract_ppt_content(file_path)
        
        return {
            "success": True,
            "data": result,
            "message": f"Successfully extracted content from {result['total_slides']} slides"
        }
    
    except Exception as e:
        logger.error(f"Error in extract_ppt_info: {e}")
        return {"error": str(e)}

@server.tool
async def analyze_word_structure(file_path: str) -> Dict[str, Any]:
    """Analyze the structure of a Word document."""
    try:
        if not os.path.exists(file_path):
            return {"error": f"File not found: {file_path}"}
        
        if not file_path.lower().endswith(('.docx', '.doc')):
            return {"error": "File must be a Word document (.docx or .doc)"}
        
        analyzer = WordStructureAnalyzer()
        result = analyzer.analyze_word_structure(file_path)
        
        return {
            "success": True,
            "data": result,
            "message": f"Successfully analyzed structure with {len(result['sections'])} sections"
        }
    
    except Exception as e:
        logger.error(f"Error in analyze_word_structure: {e}")
        return {"error": str(e)}

@server.tool
async def combine_ppt_and_word(ppt_file: str, word_file: str, output_file: Optional[str] = None) -> Dict[str, Any]:
    """Combine PowerPoint content with Word document structure."""
    try:
        # Extract PPT content
        ppt_result = await extract_ppt_info(ppt_file)
        if not ppt_result.get("success"):
            return ppt_result
        
        # Analyze Word structure
        word_result = await analyze_word_structure(word_file)
        if not word_result.get("success"):
            return word_result
        
        # Combine the data
        combiner = DocumentCombiner()
        combined_data = combiner.combine_documents(ppt_result["data"], word_result["data"])
        
        # Generate output file path
        if not output_file:
            output_file = os.path.join(tempfile.gettempdir(), "combined_presentation.pptx")
        
        # Generate the presentation
        generator = PPTGenerator()
        output_path = generator.generate_presentation(combined_data, output_file)
        
        return {
            "success": True,
            "output_file": output_path,
            "slides_created": len(combined_data["slides"]),
            "combined_data": combined_data,
            "message": f"Successfully created combined presentation with {len(combined_data['slides'])} slides"
        }
    
    except Exception as e:
        logger.error(f"Error in combine_ppt_and_word: {e}")
        return {"error": str(e)}

@server.tool
async def create_presentation_from_structure(structure_data: Dict[str, Any], output_file: Optional[str] = None) -> Dict[str, Any]:
    """Create a PowerPoint presentation from structured data."""
    try:
        if not output_file:
            output_file = os.path.join(tempfile.gettempdir(), "generated_presentation.pptx")
        
        generator = PPTGenerator()
        output_path = generator.generate_presentation(structure_data, output_file)
        
        return {
            "success": True,
            "output_file": output_path,
            "slides_created": len(structure_data.get("slides", [])),
            "message": f"Successfully created presentation at {output_path}"
        }
    
    except Exception as e:
        logger.error(f"Error in create_presentation_from_structure: {e}")
        return {"error": str(e)}

if __name__ == "__main__":
    asyncio.run(server.run_stdio())