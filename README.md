# OpenXmlMcp Server

OpenXmlMcp is an MCP server for creating, reading, and editing Office documents (`.docx`, `.xlsx`, `.pptx`) using the Open XML SDK.

It provides both low-level document helpers and high-level session-based editing tools, including validation, batch execution, undo checkpoints, structured list formatting, and style management.

## Features

- Session-based editing for Word, Excel, and PowerPoint documents
- Unified operation validation with read-only and document-type safety checks
- Batch execution with structured per-operation error reporting
- Undo/checkpoint history for write operations
- Cross-suite style presets and text presets
- Rich Word tooling for headings, spacing, lists, anchor-based insertion, and custom styles
- Excel helpers for ranges, formulas, and cell styling
- PowerPoint helpers for slide management and text styling

## Tool Categories

## Base DOCX Utilities

- `create_docx_base64(title, body)`
- `extract_docx_text(base64Docx)`
- `count_docx_paragraphs(base64Docx)`
- `extract_docx_text_from_path(filePath)`
- `count_docx_paragraphs_from_path(filePath)`

## Session Management

- `open_document(filePath, readOnly)`
- `create_document(filePath, documentType)`
- `save_document(sessionId)`
- `close_document(sessionId)`
- `get_document_info(sessionId)`
- `list_structure(sessionId)`
- `find_text(sessionId, query)`
- `validate_operation(sessionId, operationName)`
- `batch_execute(sessionId, operationsJson)`
- `get_operation_history(sessionId)`
- `undo_last_change(sessionId)`

## Word

- Content: `word_append_paragraph`, `word_insert_paragraph_at`, `word_insert_paragraph_after_text`, `word_insert_text_after_text`, `word_replace_text`, `word_add_heading`, `word_add_table`
- Lists: `word_add_bulleted_list`, `word_add_numbered_list`, `word_add_structured_list`
- Formatting: `word_set_paragraph_style`, `word_set_paragraph_spacing`, `word_set_document_spacing_preset`
- Styles: `word_list_styles`, `word_apply_style_by_name`, `word_create_or_update_style`

## Excel

- Data: `excel_set_cell_value`, `excel_get_cell_value`, `excel_set_range_values`, `excel_get_used_range`
- Formula: `excel_set_formula`, `excel_get_formula`
- Structure: `excel_add_worksheet`
- Formatting: `excel_set_cell_style`

## PowerPoint

- Slides: `powerpoint_add_slide`, `powerpoint_add_bullet_slide`, `powerpoint_insert_slide_at`, `powerpoint_reorder_slide`, `powerpoint_delete_slide`
- Text: `powerpoint_set_slide_title`, `powerpoint_set_slide_body`, `powerpoint_set_text_style`

## Cross-Suite Presets

- `apply_style_preset(sessionId, preset)`
- `apply_text_preset(sessionId, preset, targetIndex)`

## Behavior Guarantees

- Operation aliases are normalized to canonical names (e.g. `powerpoint_*` and `power_point_*`).
- Write operations are rejected in read-only sessions.
- Operations are validated against the active session's document type.
- Batch failures include `operation`, `index`, `errorCode`, and `error`.
- PPTX defaults are generated programmatically (no embedded template dependency).

## Run Locally

```bash
dotnet run --project OpenXmlMcp.Server
```

## MCP Client Configuration (stdio)

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

## Development

Run the test suite:

```bash
dotnet test OpenXmlMcp.slnx
```

## License

GPL-3.0. See `LICENSE`.
