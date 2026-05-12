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

## Phase 2 tools

- `word_add_table(sessionId, rows, columns)` inserts a basic table in Word.
- `excel_add_worksheet(sessionId, sheetName)` adds a worksheet to Excel.
- `powerpoint_add_bullet_slide(sessionId, title, bulletLines)` adds a bullet-style slide.
- `batch_execute(sessionId, operationsJson)` executes a JSON array of operations and returns per-operation results.

## Phase 3 tools

- `get_operation_history(sessionId)` returns an operation audit trail.
- `undo_last_change(sessionId)` rolls back to the latest checkpoint.
- Safety guard: opening files larger than 20 MB is blocked.

## Phase 4 (Word) tools

- `word_insert_paragraph_at(sessionId, index, text)` inserts a paragraph at a 1-based index.
- `word_replace_text(sessionId, find, replace, matchCase)` replaces text and returns replacement count.
- `word_add_heading(sessionId, level, text)` appends heading style paragraphs (levels 1-6).
- `word_add_bulleted_list(sessionId, lines)` appends newline-separated bullet items.

## Phase 4 (Excel) tools

- `excel_get_used_range(sessionId, sheetName)` returns used range bounds and dimensions.
- `excel_set_range_values(sessionId, sheetName, startCell, valuesJson)` sets a 2D value matrix from JSON.
- `excel_set_formula(sessionId, sheetName, cellReference, formula)` writes a formula.
- `excel_get_formula(sessionId, sheetName, cellReference)` reads a formula.

## Phase 4 (PowerPoint) tools

- `powerpoint_insert_slide_at(sessionId, index, title, body)` inserts at 1-based index.
- `powerpoint_set_slide_title(sessionId, slideIndex, title)` updates slide title text.
- `powerpoint_set_slide_body(sessionId, slideIndex, body)` updates slide body text.
- `powerpoint_reorder_slide(sessionId, fromIndex, toIndex)` reorders slides.
- `powerpoint_delete_slide(sessionId, slideIndex)` deletes a slide by 1-based index.

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
