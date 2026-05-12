# OpenXmlMcp Server

MCP server built with the official C# MCP SDK and the Open XML SDK.

## Available tools

- `create_docx_base64(title, body)` creates a DOCX and returns a base64 string.
- `extract_docx_text(base64Docx)` extracts plain text from a base64 DOCX.
- `count_docx_paragraphs(base64Docx)` counts paragraphs in a base64 DOCX.
- `extract_docx_text_from_path(filePath)` extracts plain text from a DOCX file path.
- `count_docx_paragraphs_from_path(filePath)` counts paragraphs in a DOCX file path.

## Phase 1 Office session tools

- `open_document(filePath, readOnly)` opens a `.docx`, `.xlsx`, or `.pptx` file.
- `create_document(filePath, documentType)` creates `docx`, `xlsx`, or `pptx` and opens a session.
- `save_document(sessionId)` validates an active session and persists in-memory changes.
- `close_document(sessionId)` closes an active session.
- `get_document_info(sessionId)` returns metadata and file details for a session.
- `list_structure(sessionId)` returns structural summary (paragraphs/tables, sheets, slide count).
- `find_text(sessionId, query)` performs a case-insensitive text search.
- `validate_operation(sessionId, operationName)` dry-runs session constraints before execution.
- `word_append_paragraph(sessionId, text)` appends text to a Word document.
- `excel_set_cell_value(sessionId, sheetName, cellReference, value)` writes a string cell value.
- `excel_get_cell_value(sessionId, sheetName, cellReference)` reads a string cell value.
- `powerpoint_add_slide(sessionId, title, body)` adds a simple title/body slide.

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
