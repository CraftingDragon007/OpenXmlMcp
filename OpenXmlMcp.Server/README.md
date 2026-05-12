# OpenXmlMcp Server

MCP server built with the official C# MCP SDK and the Open XML SDK.

## Available tools

- `create_docx_base64(title, body)` creates a DOCX and returns a base64 string.
- `extract_docx_text(base64Docx)` extracts plain text from a base64 DOCX.
- `count_docx_paragraphs(base64Docx)` counts paragraphs in a base64 DOCX.
- `extract_docx_text_from_path(filePath)` extracts plain text from a DOCX file path.
- `count_docx_paragraphs_from_path(filePath)` counts paragraphs in a DOCX file path.

## Run locally

```bash
dotnet run --project OpenXmlMcp.Server
```

## Configure in MCP client (stdio)

```json
{
  "servers": {
    "openxml-mcp": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "<PATH_TO_REPO>/OpenXmlMcp.Server"
      ]
    }
  }
}
```

## Test

```bash
dotnet test OpenXmlMcp.slnx
```
