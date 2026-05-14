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
- Excel helpers for ranges, formulas, cell styling, and cell inspection
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

`batch_execute` accepts either:

- a JSON array of operation objects, or
- a JSON object with an `operations` array.

Each operation object can use `operation` (preferred) or `operationName` (legacy alias).

## Word

- Content: `word_append_paragraph`, `word_insert_paragraph_at`, `word_insert_paragraph_after_text`, `word_insert_text_after_text`, `word_replace_text`, `word_add_heading`, `word_add_table`
- Table cells: `word_set_table_cell`, `word_get_table_cell`
- Readback: `word_get_paragraph_info`
- Lists: `word_add_bulleted_list`, `word_add_numbered_list`, `word_add_structured_list`
- Formatting: `word_set_paragraph_style`, `word_set_paragraph_spacing`, `word_set_document_spacing_preset`
- Styles: `word_list_styles`, `word_apply_style_by_name`, `word_create_or_update_style`

## Excel

- Data: `excel_set_cell_value`, `excel_get_cell_value`, `excel_get_cell_info`, `excel_get_cell_style`, `excel_set_range_values`, `excel_get_used_range`
- Formula: `excel_set_formula`, `excel_get_formula`
- Structure: `excel_add_worksheet`
- Formatting: `excel_set_cell_style`

## PowerPoint

- Slides: `powerpoint_add_slide` (supports `bodyType`: `text`, `bulleted`, `numbered`), `powerpoint_add_bullet_slide` (compat wrapper), `powerpoint_insert_slide_at`, `powerpoint_reorder_slide`, `powerpoint_delete_slide`
- Text: `powerpoint_set_slide_title`, `powerpoint_set_slide_body` (supports `bodyType`: `text`, `bulleted`, `numbered`), `powerpoint_set_text_style`
- Text readback: `powerpoint_get_text_style`
- Notes: `powerpoint_set_slide_notes`, `powerpoint_get_slide_notes`

`powerpoint_set_text_style` uses `slot=0` for title and `slot=1` for the entire body region (all body paragraphs, including bulleted/numbered lines).

## Cross-Suite Presets

- `apply_style_preset(sessionId, preset)`
- `list_style_presets(sessionId)`
- `apply_text_preset(sessionId, preset, targetIndex)`

## Behavior Guarantees

- Operation aliases are normalized to canonical names (e.g. `powerpoint_*` and `power_point_*`).
- Write operations are rejected in read-only sessions.
- Operations are validated against the active session's document type.
- Batch failures include `operation`, `index`, `errorCode`, and `error`.
- PPTX defaults are generated programmatically (no embedded template dependency).
- Mutating tools return a structured result payload (`ok`, `changed`, `operation`, `target`).

Excel note: formulas are stored and retrievable, but not calculated server-side; formula cells may not have a cached value until recalculated by a spreadsheet client. The server now marks workbook calculation properties to force recalculation on open in spreadsheet clients.

## Indexing And Structure

- Word paragraph write tools (`word_set_paragraph_style`, `word_set_paragraph_spacing`, `word_apply_style_by_name`) use **1-based body paragraph indexes** (direct body paragraphs only; table-cell paragraphs are excluded).
- `find_text` for Word returns rich match metadata:
  - `index` (1-based over all Word paragraphs, including table cells)
  - `bodyParagraphIndex` (nullable; set only for body paragraphs)
  - `addressableByParagraphTools` (boolean)
  - `tableIndex`/`rowIndex`/`columnIndex` for table-cell matches
- `list_structure` now returns rich structures:
  - Word: `elements` and `tables` summaries in addition to counts
  - PowerPoint: `slides` metadata in addition to `slideCount`

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

## Examples

Batch execute with wrapped payload:

```json
{
  "operations": [
    { "operation": "word_append_paragraph", "text": "Hello" },
    { "operationName": "word_add_heading", "level": 2, "text": "Section" }
  ]
}
```

Set and read a Word table cell:

```text
word_set_table_cell(sessionId, tableIndex=1, rowIndex=2, columnIndex=1, text="R2C1")
word_get_table_cell(sessionId, tableIndex=1, rowIndex=2, columnIndex=1)
```

Set Excel range values with strict JSON:

```json
[["A", "B"], ["=SUM(C1:D1)", 42]]
```

Discover and apply style presets:

```text
list_style_presets(sessionId)
apply_style_preset(sessionId, "default")
```

## License

GPL-3.0. See `LICENSE`.
